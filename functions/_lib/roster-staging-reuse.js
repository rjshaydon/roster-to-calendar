import { reserveRosterMaintenanceBudget } from './roster-maintenance-budget.js';
import { PATCH_SECTIONS, validatePatchEntry } from './roster-patch.js';
import { rosterImportDigest } from './roster-import-batches.js';
import { derivedRosterFileUpsertStatement, upsertRosterFileStatusSummaryStatement } from './d1-calendar.js';

// Inspect only a fixed number of indexed, inactive candidates. Never borrow
// active clinical data or a different provider file/term as a staging baseline.
export async function reusableRosterStage(db,run,manifest) {
 const existing=await db.prepare('SELECT file_id,prepared FROM roster_rebase_jobs WHERE run_id=?').bind(run.id).first();
 const own=await db.prepare('SELECT run_id FROM roster_import_jobs WHERE run_id=?').bind(run.id).first();
 if(existing&&(existing.prepared||!own))return existing.file_id;
 if(own)return '';
 const rows=(await db.prepare(`WITH candidates AS (
  SELECT run_id,file_id,manifest_json,event_count,next_batch FROM roster_import_jobs INDEXED BY idx_roster_import_reuse
  WHERE source_id=? AND activated=0 AND initialized=1 AND json_extract(manifest_json,'$.startDate')=?
  AND json_extract(manifest_json,'$.rosterEndDate')=? ORDER BY event_count DESC LIMIT 8
 ) SELECT c.file_id FROM candidates c JOIN roster_files f ON f.id=c.file_id
 JOIN roster_sync_runs r ON r.id=c.run_id
 JOIN raw_roster_files raw ON raw.file_id=COALESCE(NULLIF(r.source_file_id,''),r.file_id)
 JOIN raw_roster_files incoming ON incoming.file_id=?
 WHERE f.active=0 AND c.event_count=json_extract(c.manifest_json,'$.eventCount')
 AND json_type(c.manifest_json,'$.compact')='object'
 AND c.next_batch=json_array_length(json_extract(c.manifest_json,'$.batches'))
 AND LOWER(raw.name)=LOWER(incoming.name)
 AND json_extract(c.manifest_json,'$.startDate')=?
 AND json_extract(c.manifest_json,'$.rosterEndDate')=? LIMIT 1`)
 .bind(run.sourceId,manifest.startDate,manifest.rosterEndDate,run.sourceFileId||run.fileId,manifest.startDate,manifest.rosterEndDate).all()).results;
 return rows[0]?.file_id||'';
}

export async function beginRosterRebase(db,run,fileId,range,cachedPresence=false) {
 const old=await db.prepare('SELECT * FROM roster_rebase_jobs WHERE run_id=?').bind(run.id).first();
 if(old)return {...baseline(old),presenceMode:old.plan_revision?JSON.parse(old.manifest_json).presenceMode:(cachedPresence?'cached':undefined)};
 const candidate=await reusableRosterStage(db,run,range);
 if(!candidate||candidate!==fileId)throw Error('Reusable inactive roster does not match this provider file/term.');
 const parent=await db.prepare(`SELECT j.*,f.parsed_at,f.source_type FROM roster_import_jobs j
  JOIN roster_files f ON f.id=j.file_id WHERE j.file_id=? AND j.source_id=? AND j.activated=0 AND j.initialized=1
  AND f.active=0 ORDER BY j.event_count DESC LIMIT 1`).bind(fileId,run.sourceId).first();
 if(!parent)throw Error('Reusable roster is no longer inactive.');
 if(!await reserveRosterMaintenanceBudget(db,64,128))return {deferred:true};
 await db.batch([
  db.prepare(`SELECT CASE WHEN EXISTS(SELECT 1 FROM roster_import_jobs j JOIN roster_files f ON f.id=j.file_id
   WHERE j.run_id=? AND j.plan_revision=? AND j.rebase_owner=? AND j.prepared_batch=? AND j.next_batch=?
   AND j.activated=0 AND f.active=0 AND f.parsed_at=?) THEN 1 ELSE json('rebase-claim-conflict') END`)
   .bind(parent.run_id,parent.plan_revision,parent.rebase_owner,parent.prepared_batch,parent.next_batch,parent.parsed_at),
  db.prepare('UPDATE roster_import_jobs SET rebase_owner=? WHERE run_id=?').bind(run.id,parent.run_id),
  db.prepare(`INSERT INTO roster_rebase_jobs(run_id,baseline_run_id,file_id,source_id,baseline_revision,baseline_parsed_at)
   VALUES(?,?,?,?,?,?)`).bind(run.id,parent.run_id,fileId,run.sourceId,parent.plan_revision,parent.parsed_at),
 ]);
 return {...baseline({file_id:fileId,baseline_revision:parent.plan_revision,baseline_parsed_at:parent.parsed_at}),presenceMode:cachedPresence?'cached':undefined};
}
function baseline(job){return {mode:'rebase',fileId:job.file_id,baselineRevision:job.baseline_revision,baselineParsedAt:job.baseline_parsed_at,prepared:Boolean(job.prepared),revision:job.plan_revision,nextBatch:job.next_batch||0,batchCount:job.plan_revision?JSON.parse(job.manifest_json).batches.length:0};}
async function jobFor(db,run){
 const j=await db.prepare(`SELECT j.*,f.source_type FROM roster_rebase_jobs j JOIN roster_files f ON f.id=j.file_id WHERE run_id=?`).bind(run.id).first();
 if(!j||j.source_id!==run.sourceId)throw Error('Staging reuse belongs to another source.');
 return j;
}
function guard(db,j){return db.prepare(`SELECT CASE WHEN EXISTS(SELECT 1 FROM roster_files f
 JOIN roster_import_jobs p ON p.run_id=? WHERE f.id=? AND f.active=0 AND f.parsed_at=?
 AND p.file_id=f.id AND p.plan_revision=? AND p.rebase_owner=? AND p.activated=0)
 THEN 1 ELSE json('rebase-baseline-conflict') END`)
 .bind(j.baseline_run_id,j.file_id,j.baseline_parsed_at,j.baseline_revision,j.run_id);}

export async function readRosterRebasePage(db,run,input){
 const j=await jobFor(db,run),section=PATCH_SECTIONS[input.section];
 if(!section&&input.section!=='presence')throw Error('Invalid staging inspection section.');
 if(String(input.after||'').length>500)throw Error('Invalid staging cursor.');
 if(!await reserveRosterMaintenanceBudget(db,0,4096))return {deferred:true};
 await guard(db,j).first();
 const numeric=input.section==='presence'||section.key==='rowid';
 const after=numeric?Number(input.after||0):String(input.after||'');
 if(numeric&&(!Number.isSafeInteger(after)||after<0))throw Error('Invalid staging cursor.');
 // Presence probes use the exact event-id index, at most 32 events per page.
 const sql=input.section==='presence'?`SELECT e.rowid AS patch_cursor,e.id,
  (SELECT json_group_array(json_array(date,source_type,doctor_key,display_name,event_id))
   FROM (SELECT * FROM roster_daily_presence INDEXED BY idx_roster_daily_presence_event_id WHERE event_id=e.id LIMIT 121)) AS rows_json
  FROM roster_events e WHERE e.file_id=? AND e.rowid>? ORDER BY e.rowid LIMIT 33`
  :`SELECT ${section.columns},${section.key} AS patch_cursor FROM ${section.table} WHERE file_id=? AND ${section.key}>? ORDER BY ${section.key} LIMIT 33`;
 const rows=(await db.prepare(sql).bind(j.file_id,after).all()).results;
 if(input.section==='presence'&&rows.some(r=>JSON.parse(r.rows_json).length>120))throw Error('Staged presence exceeds its event bound.');
 return {rows:rows.slice(0,32),next:rows.length>32?rows[31].patch_cursor:''};
}
export async function pinRosterRebasePlan(db,run,input){
 const j=await jobFor(db,run),p=input.manifest;
 if(j.plan_revision){if(j.plan_revision!==input.patchRevision)throw Error('Staging reuse plan changed.');return {nextBatch:j.next_batch,prepared:Boolean(j.prepared)};}
 const parent=await db.prepare('SELECT manifest_json FROM roster_import_jobs WHERE run_id=?').bind(j.baseline_run_id).first();
 const old=JSON.parse(parent.manifest_json);
 if(!p||p.fileId!==j.file_id||p.baselineRevision!==j.baseline_revision||p.baselineParsedAt!==j.baseline_parsed_at
  ||p.coverage?.start!==old.startDate||p.rosterEndDate!==old.rosterEndDate
  ||!Array.isArray(p.doctors)||p.doctors.length!==p.doctorCount||p.doctorCount<1||p.doctorCount>512
  ||new Set(p.doctors.map(d=>d.key)).size!==p.doctorCount
  ||!Number.isInteger(p.eventCount)||p.eventCount<1||p.eventCount>25000
  ||!Number.isInteger(p.issueCount)||p.issueCount<0||p.issueCount>5000
  ||(p.presenceMode==='cached'&&p.presenceCount!==0)
  ||!Number.isInteger(p.presenceCount)||p.presenceCount<(p.presenceMode==='cached'?0:p.eventCount)||p.presenceCount>50000
  ||!Number.isInteger(p.changedFacts)||p.changedFacts<0||p.changedFacts>1250
  ||![p.contentRevision,p.coverage.staffDigest,p.coverage.dailyDigest].every(v=>/^[a-f0-9]{64}$/.test(v))
  ||!Array.isArray(p.batches)||p.batches.length>800||JSON.stringify(p).length>262144
  ||p.batches.some((b,i)=>b.index!==i||!Number.isInteger(b.count)||b.count<1||b.count>16||!/^[a-f0-9]{64}$/.test(b.hash))
  ||await rosterImportDigest(p)!==input.patchRevision)throw Error('Invalid bounded staging reuse plan.');
 if(!await reserveRosterMaintenanceBudget(db,64,128))return {deferred:true};
 await db.batch([guard(db,j),db.prepare("UPDATE roster_rebase_jobs SET manifest_json=?,plan_revision=? WHERE run_id=? AND plan_revision=''").bind(JSON.stringify(p),input.patchRevision,run.id)]);
 return {nextBatch:0};
}

const inserts={
 doctor:['roster_file_doctors','file_id,source_type,doctor_key,display_name,seniority,membership_source,provider_staff_id','file_id,source_type,doctor_key'],
 event:['roster_events','id,file_id,source_type,doctor_key,display_name,start_date,end_date,start_ts,end_ts,title,raw_value,seniority,provider_staff_id,location,all_day,time_label,event_json','id'],
 issue:['roster_issues','id,file_id,source_type,doctor_key,display_name,start_date,raw_value,seniority,status,message,resolution_type,suggested_title,time_label,issue_json','id'],
 presence:['roster_daily_presence','date,source_type,doctor_key,display_name,event_id','date,source_type,doctor_key,event_id'],
 staff:['facility_term_staff_contributions','source_type,term_start,doctor_key,file_id,display_name,seniority,membership_source,provider_staff_id,first_applicable_date,last_applicable_date,fact_digest,updated_at','source_type,term_start,doctor_key,file_id'],
 catalog:['facility_stream_catalog_contributions','source_type,term_start,file_id,catalog_key,seniority,title,raw_value,location,all_day,time_label,start_time,end_time,first_date,last_date,fact_digest,updated_at','source_type,term_start,file_id,catalog_key'],
};
function batchKey(j,index){return `roster-rebase/v1/${encodeURIComponent(j.run_id)}/${j.plan_revision}/${index}.json`;}
export async function saveRosterRebaseBatch(db,run,input,bucket){
 const j=await jobFor(db,run),p=JSON.parse(j.manifest_json||'null'),b=input.batch,d=p?.batches[b?.index];
 if(!bucket?.put||!d||input.patchRevision!==j.plan_revision||!Array.isArray(b.entries)||b.entries.length!==d.count
  ||JSON.stringify(b).length>24576||await rosterImportDigest(b)!==d.hash)throw Error('Invalid retained staging batch.');
 for(const e of b.entries)validatePatchEntry(e,j,p);
 await guard(db,j).first();
 await bucket.put(batchKey(j,b.index),JSON.stringify(b),{httpMetadata:{contentType:'application/json'}});
 return {saved:true};
}
export async function applyRosterRebaseBatch(db,run,input,bucket){
 const j=await jobFor(db,run),p=JSON.parse(j.manifest_json||'null');
 if(!Number.isInteger(input.batchIndex)||!p?.batches[input.batchIndex])throw Error('Invalid retained staging cursor.');
 const object=await bucket?.get(batchKey(j,input.batchIndex));
 const b=object?await object.json():null;
 const descriptor=p?.batches[b?.index];
 if(!p||b?.index!==input.batchIndex||input.patchRevision!==j.plan_revision||!descriptor||!Array.isArray(b.entries)||b.entries.length!==descriptor.count
  ||JSON.stringify(b).length>24576||await rosterImportDigest(b)!==descriptor.hash)throw Error('Staging reuse batch differs from its pinned plan.');
 if(b.index<j.next_batch)return {duplicate:true,nextBatch:j.next_batch};
 if(b.index!==j.next_batch||j.prepared)throw Error('Staging reuse batches must be ordered.');
 for(const e of b.entries)validatePatchEntry(e,j,p);
 const affected=b.entries.filter(e=>['event','removeEvent'].includes(e.kind)).map(e=>e.row[0]);
 const oldPresence=affected.length?Number((await db.prepare(`SELECT COUNT(*) AS n FROM roster_daily_presence INDEXED BY idx_roster_daily_presence_event_id WHERE event_id IN (${affected.map(()=>'?')})`).bind(...affected).first()).n):0;
 if(oldPresence>1920)throw Error('Staging presence deletion exceeds its bound.');
 if(!await reserveRosterMaintenanceBudget(db,b.entries.length*32+oldPresence*5+128,4096))return {deferred:true};
 const statements=[guard(db,j),db.prepare(`SELECT CASE WHEN EXISTS(SELECT 1 FROM roster_rebase_jobs WHERE run_id=? AND plan_revision=? AND next_batch=? AND prepared=0) THEN 1 ELSE json('rebase-cursor-conflict') END`).bind(run.id,j.plan_revision,b.index)];
 if(affected.length)statements.push(db.prepare(`DELETE FROM roster_daily_presence WHERE event_id IN (${affected.map(()=>'?')})`).bind(...affected));
 for(const {kind,row} of b.entries){
  if(inserts[kind]){
   const [table,columns,conflict]=inserts[kind],names=columns.split(','),keys=conflict.split(',');
   const update=names.filter(n=>!keys.includes(n)).map(n=>`${n}=excluded.${n}`).join(',');
   statements.push(db.prepare(`INSERT INTO ${table}(${columns}) VALUES(${row.map(()=>'?')}) ON CONFLICT(${conflict}) DO UPDATE SET ${update}`).bind(...row));
  }else{
   const removal={removeDoctor:['roster_file_doctors','doctor_key'],removeEvent:['roster_events','id'],removeIssue:['roster_issues','id'],removeStaff:['facility_term_staff_contributions','doctor_key'],removeCatalog:['facility_stream_catalog_contributions','catalog_key']}[kind];
   const composite=['removeStaff','removeCatalog'].includes(kind);
   statements.push(db.prepare(`DELETE FROM ${removal[0]} WHERE file_id=? AND ${composite?'term_start=? AND ':''}${removal[1]}=?`).bind(j.file_id,...row));
  }
 }
 statements.push(db.prepare('UPDATE roster_rebase_jobs SET next_batch=next_batch+1 WHERE run_id=? AND next_batch=?').bind(run.id,b.index));
 await db.batch(statements);
 return {nextBatch:b.index+1};
}

export async function prepareRosterRebase(db,run,input,file){
 const j=await jobFor(db,run),p=JSON.parse(j.manifest_json||'null');
 if(!p||input.patchRevision!==j.plan_revision||j.next_batch!==p.batches.length)throw Error('Staging reuse is incomplete.');
 if(j.prepared)return {prepared:true,fileId:j.file_id,revision:j.plan_revision};
 if(!await reserveRosterMaintenanceBudget(db,512+p.doctorCount*6,160000))return {deferred:true};
 const active=(await db.prepare('SELECT id,parsed_at FROM roster_files INDEXED BY idx_roster_files_source_active WHERE source_type=? AND active=1 ORDER BY id LIMIT 33').bind(j.source_type).all()).results;
 if(active.length>32)throw Error('Staging promotion exceeds its file ceiling.');
 const now=new Date().toISOString(),desired={...file,id:j.file_id,active:false};
 const manifest={doctors:p.doctors,eventCount:p.eventCount,issueCount:p.issueCount};
 await db.batch([guard(db,j),
  db.prepare(`SELECT CASE WHEN (SELECT COUNT(*) FROM roster_events WHERE file_id=?)=?
   AND (SELECT COUNT(*) FROM roster_issues WHERE file_id=?)=? AND (SELECT COUNT(*) FROM roster_file_doctors WHERE file_id=?)=?
   ${p.presenceMode==='cached'?'':`AND (SELECT COUNT(*) FROM roster_events e JOIN roster_daily_presence d INDEXED BY idx_roster_daily_presence_event_id ON d.event_id=e.id WHERE e.file_id=?)=?`}
   THEN 1 ELSE json('rebase-final-count-conflict') END`).bind(j.file_id,p.eventCount,j.file_id,p.issueCount,j.file_id,p.doctorCount,...(p.presenceMode==='cached'?[]:[j.file_id,p.presenceCount])),
  db.prepare(`INSERT INTO roster_file_coverage(file_id,source_type,coverage_start,coverage_end,content_revision,staff_digest,daily_digest,updated_at)
   VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(file_id) DO UPDATE SET coverage_start=excluded.coverage_start,coverage_end=excluded.coverage_end,content_revision=excluded.content_revision,staff_digest=excluded.staff_digest,daily_digest=excluded.daily_digest,updated_at=excluded.updated_at`)
   .bind(j.file_id,j.source_type,p.coverage.start,p.coverage.end,j.plan_revision,p.coverage.staffDigest,p.coverage.dailyDigest,now),
  derivedRosterFileUpsertStatement(db,desired,j.source_type,now),
  upsertRosterFileStatusSummaryStatement(db,{...desired,derivedState:'ready',expectedDoctorCount:p.doctorCount,indexedDoctorCount:p.doctorCount,eventCount:p.eventCount,contentRevision:j.plan_revision}),
  db.prepare(`INSERT INTO roster_import_jobs(run_id,file_id,source_id,plan_revision,manifest_json,initialized,compact_ready,promotion_fence_json) VALUES(?,?,?,?,?,1,1,?)`)
   .bind(run.id,j.file_id,run.sourceId,j.plan_revision,JSON.stringify(manifest),JSON.stringify(active)),
  db.prepare(`INSERT INTO roster_doctors(source_type,doctor_key,display_name,updated_at)
   SELECT ?,json_extract(value,'$.key'),json_extract(value,'$.displayName'),? FROM json_each(?) WHERE 1
   ON CONFLICT(source_type,doctor_key) DO UPDATE SET display_name=excluded.display_name,updated_at=excluded.updated_at`)
   .bind(j.source_type,now,JSON.stringify(p.doctors)),
  db.prepare(`INSERT INTO facility_term_visibility(source_type,term_start,visible_from,revision,updated_at)
   SELECT DISTINCT source_type,term_start,date(term_start,'start of month','-1 month'),'',? FROM facility_term_staff_contributions WHERE file_id=?
   ON CONFLICT(source_type,term_start) DO NOTHING`).bind(now,j.file_id),
  db.prepare('UPDATE roster_rebase_jobs SET prepared=1 WHERE run_id=? AND next_batch=?').bind(run.id,p.batches.length),
 ]);
 return {prepared:true,fileId:j.file_id,revision:j.plan_revision};
}
