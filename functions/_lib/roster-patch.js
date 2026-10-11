import { rosterImportDigest } from './roster-import-batches.js';
import { reserveRosterMaintenanceBudget } from './roster-maintenance-budget.js';
import { derivedRosterFileUpsertStatement,upsertRosterFileStatusSummaryStatement,facilitySmsMembershipStatement } from './d1-calendar.js';
import { facilityRefreshStatements } from './facility-refresh-queue.js';

// Fixed sections with indexed keyset pagination; callers cannot choose SQL.
export const PATCH_SECTIONS={
  doctors:{table:'roster_file_doctors',key:'doctor_key',columns:'doctor_key,display_name,seniority,membership_source,provider_staff_id',limit:512},
  events:{table:'roster_events',key:'rowid',columns:'id,doctor_key,display_name,start_date,end_date,start_ts,end_ts,title,raw_value,seniority,provider_staff_id,location,all_day,time_label,event_json',limit:25000},
  issues:{table:'roster_issues',key:'rowid',columns:'id,doctor_key,display_name,start_date,raw_value,seniority,status,message,resolution_type,suggested_title,time_label,issue_json',limit:5000},
  staff:{table:'facility_term_staff_contributions',key:"term_start||'|'||doctor_key",columns:'*',limit:750},
  catalog:{table:'facility_stream_catalog_contributions',key:"term_start||'|'||catalog_key",columns:'*',limit:750},
};
export async function beginRosterPatch(db,run,fileId) {
  const old=await db.prepare('SELECT * FROM roster_patch_jobs WHERE run_id=?').bind(run.id).first();
  if(old) return {mode:'patch',fileId:old.file_id,baselineRevision:old.baseline_revision,baselineParsedAt:old.baseline_parsed_at,activated:Boolean(old.activated),patchRevision:old.plan_revision};
  const baseline=await db.prepare(`SELECT f.parsed_at,f.source_type,c.content_revision FROM roster_files f JOIN roster_file_coverage c ON c.file_id=f.id
    JOIN roster_file_status_summaries s ON s.file_id=f.id WHERE f.id=? AND f.source_id=? AND f.active=1 AND s.derived_state='ready'`).bind(fileId,run.sourceId).first();
  if(!baseline?.content_revision) throw Error('Patch requires an active, fully prepared roster.');
  if(!await reserveRosterMaintenanceBudget(db,32,256)) return {deferred:true};
  await db.prepare('INSERT INTO roster_patch_jobs(run_id,file_id,source_id,source_type,baseline_revision,baseline_parsed_at) VALUES(?,?,?,?,?,?)')
    .bind(run.id,fileId,run.sourceId,baseline.source_type,baseline.content_revision,baseline.parsed_at).run();
  return {mode:'patch',fileId,baselineRevision:baseline.content_revision,baselineParsedAt:baseline.parsed_at};
}
async function jobFor(db,run) {
  const job=await db.prepare('SELECT * FROM roster_patch_jobs WHERE run_id=?').bind(run.id).first();
  if(!job || job.source_id!==run.sourceId) throw Error('Patch does not match its queued source.');
  return job;
}
export async function readRosterPatchPage(db,run,input) {
  const job=await jobFor(db,run), section=PATCH_SECTIONS[input.section];
  if(!section || String(input.after||'').length>500) throw Error('Invalid patch inspection section.');
  const after = section.key==='rowid' ? Number(input.after || 0) : String(input.after || '');
  if(section.key==='rowid' && (!Number.isSafeInteger(after) || after<0)) throw Error('Invalid indexed patch cursor.');
  if(!await reserveRosterMaintenanceBudget(db,0,2048)) return {deferred:true};
  await patchGuard(db,job).first();
  const rows=(await db.prepare(`SELECT ${section.columns},${section.key} AS patch_cursor FROM ${section.table} WHERE file_id=? AND ${section.key}>? ORDER BY ${section.key} LIMIT 33`)
    .bind(job.file_id,after).all()).results;
  return {rows:rows.slice(0,32),next:rows.length>32?rows[31].patch_cursor:''};
}
export async function pinRosterPatchPlan(db,run,input) {
  const job=await jobFor(db,run), plan=input.manifest;
  validatePatchManifest(plan,job);
  const revision=await rosterImportDigest(plan);
  if(revision!==input.patchRevision) throw Error('Patch plan hash differs from its wire representation.');
  if(job.plan_revision) {
    if(job.plan_revision!==revision) throw Error('Patch plan changed; a new queued run is required.');
    return {nextBatch:job.next_batch,activated:Boolean(job.activated)};
  }
  if(!await reserveRosterMaintenanceBudget(db,32,256)) return {deferred:true};
  await db.batch([patchGuard(db,job),db.prepare("UPDATE roster_patch_jobs SET manifest_json=?,plan_revision=? WHERE run_id=? AND plan_revision=''").bind(JSON.stringify(plan),revision,run.id)]);
  return {nextBatch:0};
}
export async function stageRosterPatchBatch(db,run,input) {
  const job=await jobFor(db,run), plan=JSON.parse(job.manifest_json||'null'),batch=input.batch;
  if(!plan || input.patchRevision!==job.plan_revision) throw Error('Unpinned patch batch.');
  const descriptor=plan.batches[batch?.index];
  if(!descriptor || !Array.isArray(batch.entries) || batch.entries.length!==descriptor.count || JSON.stringify(batch).length>24576
    || await rosterImportDigest(batch)!==descriptor.hash) throw Error('Patch batch differs from its pinned plan.');
  for(const entry of batch.entries) validatePatchEntry(entry,job,plan);
  const receipt=await db.prepare('SELECT payload_hash FROM roster_patch_batches WHERE run_id=? AND batch_index=?').bind(run.id,batch.index).first();
  if(receipt) {
    if(receipt.payload_hash!==descriptor.hash) throw Error('Conflicting patch receipt.');
    return {duplicate:true,nextBatch:job.next_batch};
  }
  if(batch.index!==job.next_batch) throw Error('Patch batches must arrive in order.');
  if(!await reserveRosterMaintenanceBudget(db,64,256)) return {deferred:true};
  await db.batch([patchGuard(db,job),db.prepare(`SELECT CASE WHEN EXISTS(SELECT 1 FROM roster_patch_jobs WHERE run_id=? AND plan_revision=? AND next_batch=?) THEN 1 ELSE json('patch-cursor-conflict') END`).bind(run.id,job.plan_revision,batch.index),
    db.prepare('INSERT INTO roster_patch_batches(run_id,batch_index,payload_hash,entries_json) VALUES(?,?,?,?)').bind(run.id,batch.index,descriptor.hash,JSON.stringify(batch.entries)),
    db.prepare('UPDATE roster_patch_jobs SET next_batch=next_batch+1 WHERE run_id=? AND next_batch=?').bind(run.id,batch.index),
  ]);
  return {nextBatch:batch.index+1};
}
function patchGuard(db,job) {
  return db.prepare(`SELECT CASE WHEN EXISTS(SELECT 1 FROM roster_files f JOIN roster_file_coverage c ON c.file_id=f.id
    JOIN roster_file_status_summaries s ON s.file_id=f.id WHERE f.id=? AND f.source_id=? AND f.active=1 AND f.parsed_at=?
      AND c.content_revision=? AND s.derived_state='ready') THEN 1 ELSE json('roster-patch-baseline-conflict') END`)
    .bind(job.file_id,job.source_id,job.baseline_parsed_at,job.baseline_revision);
}
function validatePatchManifest(plan,job) {
  if(!plan || plan.fileId!==job.file_id || plan.baselineRevision!==job.baseline_revision || plan.baselineParsedAt!==job.baseline_parsed_at
    || !/^[a-f0-9]{64}$/.test(plan.contentRevision) || !Number.isInteger(plan.doctorCount) || plan.doctorCount<1 || plan.doctorCount>512
    || !Number.isInteger(plan.eventCount) || plan.eventCount<1 || plan.eventCount>25000 || !Number.isInteger(plan.issueCount) || plan.issueCount<0 || plan.issueCount>5000
    || !Number.isInteger(plan.changedFacts) || plan.changedFacts<0 || plan.changedFacts>1250 || !Array.isArray(plan.dates) || plan.dates.length>180
    || plan.dates.some(d=>!/^\d{4}-\d{2}-\d{2}$/.test(d)) || !Array.isArray(plan.batches) || plan.batches.length>250
    || JSON.stringify(plan).length>49152) throw Error('Invalid bounded patch manifest.');
  if(!plan.coverage || !/^[a-f0-9]{64}$/.test(plan.coverage.staffDigest) || !/^[a-f0-9]{64}$/.test(plan.coverage.dailyDigest)
    || !/^\d{4}-\d{2}-\d{2}$/.test(plan.coverage.start) || !/^\d{4}-\d{2}-\d{2}$/.test(plan.coverage.end)) throw Error('Invalid patch coverage.');
  let count=0;
  for(const [i,b] of plan.batches.entries()) {
    if(b.index!==i || !/^[a-f0-9]{64}$/.test(b.hash) || !Number.isInteger(b.count) || b.count<1 || b.count>16) throw Error('Invalid patch descriptor.');
    count+=b.count;
  }
  if(count>4000 || !Number.isInteger(plan.reservedWrites) || plan.reservedWrites<64 || plan.reservedWrites>50000) throw Error('Patch cost bound exceeded.');
}
function validatePatchEntry(entry,job,plan) {
  const sizes={doctor:7,event:17,issue:14,presence:5,staff:12,catalog:16,removeDoctor:1,removeEvent:1,removeIssue:1,removeStaff:2,removeCatalog:2};
  const row=entry?.row, kind=entry?.kind;
  if(!sizes[kind] || !Array.isArray(row) || row.length!==sizes[kind] || row.some(v=>typeof v!=='string' && typeof v!=='number')) throw Error('Invalid typed patch fact.');
  const sourceIndex={doctor:1,event:2,issue:2,presence:1,staff:0,catalog:0}[kind];
  if(sourceIndex!==undefined && row[sourceIndex]!==job.source_type) throw Error('Patch fact belongs to another source.');
  if(['event','issue'].includes(kind) && (row[1]!==job.file_id || !row[0].startsWith(job.file_id+':'))) throw Error('Patch event belongs to another file.');
  if(kind==='doctor' && row[0]!==job.file_id || ['staff','catalog'].includes(kind) && row[3-(kind==='catalog'?1:0)]!==job.file_id) throw Error('Patch fact belongs to another file.');
  if(['removeEvent','removeIssue'].includes(kind) && !row[0].startsWith(job.file_id+':') || kind==='presence' && !row[4].startsWith(job.file_id+':')) throw Error('Patch removal or presence belongs to another file.');
}

export async function activateRosterPatch(db,run,input,{file,deliveryGuard,publishFacility}={}) {
  const job=await jobFor(db,run);
  if(job.activated) return {activated:true,fileId:job.file_id};
  const plan=JSON.parse(job.manifest_json||'null');
  if(!plan || input.patchRevision!==job.plan_revision || job.next_batch!==plan.batches.length) throw Error('Patch has unfinished batches.');
  const total=plan.batches.reduce((n,b)=>n+b.count,0);
  const from=`FROM roster_patch_batches b,json_each(b.entries_json) e WHERE b.run_id=?`;
  const kind=k=>`${from} AND json_extract(e.value,'$.kind')='${k}'`;
  const col=i=>`json_extract(e.value,'$.row[${i}]')`;
  const affected=`SELECT ${col(0)} ${from} AND json_extract(e.value,'$.kind') IN ('event','removeEvent')`;
  if(!await reserveRosterMaintenanceBudget(db,0,160000)) return {deferred:true};
  const oldPresence=await db.prepare(`SELECT COUNT(*) AS n FROM roster_daily_presence INDEXED BY idx_roster_daily_presence_event_id WHERE event_id IN (${affected})`).bind(run.id).first();
  const writes=total*32+1024+plan.doctorCount*4+Number(oldPresence.n)*5;
  if(writes>50000) return {deferred:true,reason:'patch-write-envelope'};
  if(!await reserveRosterMaintenanceBudget(db,writes,100000)) return {deferred:true};
  const now=new Date().toISOString(), statements=[...(deliveryGuard?[deliveryGuard()]:[]),patchGuard(db,job),
    db.prepare(`SELECT CASE WHEN (SELECT COUNT(*) FROM roster_patch_batches WHERE run_id=?)=?
      AND (SELECT COALESCE(SUM(json_array_length(entries_json)),0) FROM roster_patch_batches WHERE run_id=?)=?
      AND EXISTS(SELECT 1 FROM roster_patch_jobs WHERE run_id=? AND plan_revision=? AND next_batch=? AND activated=0)
      THEN 1 ELSE json('patch-incomplete') END`).bind(run.id,plan.batches.length,run.id,total,run.id,job.plan_revision,plan.batches.length),
    db.prepare(`DELETE FROM roster_daily_presence WHERE event_id IN (${affected})`).bind(run.id),
  ];
  for(const [table,k,key,ids] of [['roster_events','removeEvent','id',col(0)],['roster_issues','removeIssue','id',col(0)],['roster_file_doctors','removeDoctor','doctor_key',col(0)]]) {
    statements.push(db.prepare(`DELETE FROM ${table} WHERE file_id=? AND ${key} IN (SELECT ${ids} ${kind(k)})`).bind(job.file_id,run.id));
  }
  for(const [table,k,key] of [['facility_term_staff_contributions','removeStaff','doctor_key'],['facility_stream_catalog_contributions','removeCatalog','catalog_key']]) {
    statements.push(db.prepare(`DELETE FROM ${table} WHERE file_id=? AND (term_start,${key}) IN (SELECT ${col(0)},${col(1)} ${kind(k)})`).bind(job.file_id,run.id));
  }
  const insert=(table,columns,k,conflict,updates,extra='')=>{
    const names=columns.split(',');
    const select=names.map((_,i)=>col(i)).join(',');
    const action=updates?`DO UPDATE SET ${updates.split(',').map(c=>`${c}=excluded.${c}`).join(',')}`:'DO NOTHING';
    statements.push(db.prepare(`INSERT INTO ${table}(${columns}) SELECT ${select} ${kind(k)}
      ON CONFLICT(${conflict}) ${action} ${extra}`).bind(run.id));
  };
  insert('roster_file_doctors','file_id,source_type,doctor_key,display_name,seniority,membership_source,provider_staff_id','doctor','file_id,source_type,doctor_key','display_name,seniority,membership_source,provider_staff_id');
  statements.push(db.prepare(`INSERT INTO roster_doctors(source_type,doctor_key,display_name,updated_at)
    SELECT ${col(1)},${col(2)},${col(3)},? ${kind('doctor')} ON CONFLICT(source_type,doctor_key) DO UPDATE SET display_name=excluded.display_name,updated_at=excluded.updated_at`).bind(now,run.id));
  insert('roster_events','id,file_id,source_type,doctor_key,display_name,start_date,end_date,start_ts,end_ts,title,raw_value,seniority,provider_staff_id,location,all_day,time_label,event_json','event','id','doctor_key,display_name,start_date,end_date,start_ts,end_ts,title,raw_value,seniority,provider_staff_id,location,all_day,time_label,event_json');
  insert('roster_issues','id,file_id,source_type,doctor_key,display_name,start_date,raw_value,seniority,status,message,resolution_type,suggested_title,time_label,issue_json','issue','id','doctor_key,display_name,start_date,raw_value,seniority,status,message,resolution_type,suggested_title,time_label,issue_json');
  insert('roster_daily_presence','date,source_type,doctor_key,display_name,event_id','presence','date,source_type,doctor_key,event_id',null);
  insert('facility_term_staff_contributions','source_type,term_start,doctor_key,file_id,display_name,seniority,membership_source,provider_staff_id,first_applicable_date,last_applicable_date,fact_digest,updated_at','staff','source_type,term_start,doctor_key,file_id','display_name,seniority,membership_source,provider_staff_id,first_applicable_date,last_applicable_date,fact_digest,updated_at');
  insert('facility_stream_catalog_contributions','source_type,term_start,file_id,catalog_key,seniority,title,raw_value,location,all_day,time_label,start_time,end_time,first_date,last_date,fact_digest,updated_at','catalog','source_type,term_start,file_id,catalog_key','seniority,title,raw_value,location,all_day,time_label,start_time,end_time,first_date,last_date,fact_digest,updated_at');
  statements.push(db.prepare(`INSERT INTO facility_term_visibility(source_type,term_start,visible_from,revision,updated_at)
    SELECT DISTINCT ${col(0)},${col(1)},date(${col(1)},'start of month','-1 month'),'',? ${kind('staff')}
    ON CONFLICT(source_type,term_start) DO NOTHING`).bind(now,run.id));
  // SMS membership is intentionally persistent. A corrected roster never deletes
  // a consultant's department membership or rewrites historical term grades.
  statements.push(facilitySmsMembershipStatement(db,job.file_id));
  statements.push(db.prepare(`SELECT CASE WHEN (SELECT COUNT(*) FROM roster_events WHERE file_id=?)=?
    AND (SELECT COUNT(*) FROM roster_issues WHERE file_id=?)=? AND (SELECT COUNT(*) FROM roster_file_doctors WHERE file_id=?)=?
    THEN 1 ELSE json('patch-final-count-conflict') END`).bind(job.file_id,plan.eventCount,job.file_id,plan.issueCount,job.file_id,plan.doctorCount));
  statements.push(db.prepare(`UPDATE roster_file_coverage SET coverage_start=?,coverage_end=?,content_revision=?,staff_digest=?,daily_digest=?,updated_at=? WHERE file_id=?`)
    .bind(plan.coverage.start,plan.coverage.end,plan.contentRevision,plan.coverage.staffDigest,plan.coverage.dailyDigest,now,job.file_id));
  const desiredFile={...file,id:job.file_id,active:true};
  statements.push(derivedRosterFileUpsertStatement(db,desiredFile,file.sourceType,now),upsertRosterFileStatusSummaryStatement(db,{...desiredFile,derivedState:'ready',expectedDoctorCount:plan.doctorCount,indexedDoctorCount:plan.doctorCount,eventCount:plan.eventCount,contentRevision:plan.contentRevision}));
  if(publishFacility && plan.dates.length) statements.push(...facilityRefreshStatements(db,file.sourceType,plan.dates,plan.contentRevision));
  statements.push(db.prepare('UPDATE roster_patch_jobs SET activated=1 WHERE run_id=? AND plan_revision=?').bind(run.id,job.plan_revision));
  await db.batch(statements);
  return {activated:true,fileId:job.file_id,doctorCount:plan.doctorCount,eventCount:plan.eventCount};
}
