import { rosterImportDigest } from './roster-import-batches.js';
import { australianTermStartForDate } from './d1-calendar.js';
import { rosterTermAvailableFrom } from '../../public/static/roster-term-policy.js';
import { reserveRosterMaintenanceBudget } from './roster-maintenance-budget.js';

// Only typed, hash-pinned facts from the authenticated background planner are
// accepted. No executable SQL or whole-term event payload crosses this boundary.
export async function stageRosterCompactMetadata(db,job,manifest,batch) {
  const compact=manifest.compact, now=new Date().toISOString();
  const source=await db.prepare('SELECT source_type,active FROM roster_files WHERE id=?').bind(job.file_id).first();
  if(!source || Number(source.active)!==0) throw Error('Compact staging requires an inactive roster.');
  if(!batch) {
    const receipts=await db.prepare('SELECT batch_index,payload_hash,plan_revision FROM roster_import_compact_receipts WHERE run_id=? ORDER BY batch_index LIMIT 48').bind(job.run_id).all();
    if(receipts.results.length!==compact.batches.length || receipts.results.some((r,i)=>r.batch_index!==i || r.payload_hash!==compact.batches[i].hash || r.plan_revision!==job.plan_revision)) throw Error('Compact preparation has unfinished batches.');
    if(!await reserveRosterMaintenanceBudget(db,64,128)) return {deferred:true};
    await db.batch([
      inactiveGuard(db,job),
      db.prepare(`INSERT INTO roster_file_coverage(file_id,source_type,coverage_start,coverage_end,content_revision,staff_digest,daily_digest,updated_at)
        VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(file_id) DO UPDATE SET coverage_start=excluded.coverage_start,coverage_end=excluded.coverage_end,
        content_revision=excluded.content_revision,staff_digest=excluded.staff_digest,daily_digest=excluded.daily_digest,updated_at=excluded.updated_at`)
        .bind(job.file_id,source.source_type,compact.coverageStart,compact.coverageEnd,job.plan_revision,compact.staffDigest,compact.dailyDigest,now),
      db.prepare('UPDATE roster_import_jobs SET compact_ready=1 WHERE run_id=? AND plan_revision=?').bind(job.run_id,job.plan_revision),
    ]);
    return {doctorCount:manifest.doctors.length,eventCount:manifest.eventCount};
  }
  const descriptor=compact.batches[batch.index];
  if(!descriptor || !Array.isArray(batch.entries) || batch.entries.length!==descriptor.count
    || batch.entries.length>16 || JSON.stringify(batch).length>24576 || await rosterImportDigest(batch)!==descriptor.hash) throw Error('Compact batch differs from its pinned manifest.');
  const receipt=await db.prepare('SELECT payload_hash,plan_revision FROM roster_import_compact_receipts WHERE run_id=? AND batch_index=?').bind(job.run_id,batch.index).first();
  if(receipt) {
    if(receipt.payload_hash!==descriptor.hash || receipt.plan_revision!==job.plan_revision) throw Error('Conflicting compact receipt.');
    return {duplicate:true};
  }
  const statements=[];
  for(const {kind,fact:f} of batch.entries) {
    if(!f || f.fileId!==job.file_id || f.sourceType!==source.source_type || !/^[a-f0-9]{64}$/.test(f.factDigest)) throw Error('Compact fact belongs to another roster.');
    const first=kind==='staff'?f.firstApplicableDate:f.firstDate, last=kind==='staff'?f.lastApplicableDate:f.lastDate;
    if(first<manifest.startDate || last>manifest.endDate || first>last || f.termStart!==australianTermStartForDate(first)) throw Error('Compact fact dates differ from the pinned term.');
    if(kind==='staff') {
      if(!manifest.doctors.some(d=>d.key===f.doctorKey)) throw Error('Compact staff is not in the pinned roster.');
      statements.push(db.prepare(`INSERT INTO facility_term_staff_contributions
        (source_type,term_start,doctor_key,file_id,display_name,seniority,membership_source,provider_staff_id,first_applicable_date,last_applicable_date,fact_digest,updated_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(source_type,term_start,doctor_key,file_id) DO UPDATE SET
        display_name=excluded.display_name,seniority=excluded.seniority,membership_source=excluded.membership_source,provider_staff_id=excluded.provider_staff_id,
        first_applicable_date=excluded.first_applicable_date,last_applicable_date=excluded.last_applicable_date,fact_digest=excluded.fact_digest,updated_at=excluded.updated_at`)
        .bind(f.sourceType,f.termStart,f.doctorKey,f.fileId,f.displayName,f.seniority,f.membershipSource,f.providerStaffId,first,last,f.factDigest,now));
      statements.push(db.prepare("INSERT INTO facility_term_visibility(source_type,term_start,visible_from,revision,updated_at) VALUES(?,?,?,'',?) ON CONFLICT(source_type,term_start) DO NOTHING")
        .bind(f.sourceType,f.termStart,rosterTermAvailableFrom(f.termStart),now));
    } else if(kind==='catalog') {
      if(!/^[a-f0-9]{64}$/.test(f.catalogKey)) throw Error('Invalid compact catalog key.');
      statements.push(db.prepare(`INSERT INTO facility_stream_catalog_contributions
        (source_type,term_start,file_id,catalog_key,seniority,title,raw_value,location,all_day,time_label,start_time,end_time,first_date,last_date,fact_digest,updated_at)
        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(source_type,term_start,file_id,catalog_key) DO UPDATE SET
        seniority=excluded.seniority,title=excluded.title,raw_value=excluded.raw_value,location=excluded.location,all_day=excluded.all_day,time_label=excluded.time_label,
        start_time=excluded.start_time,end_time=excluded.end_time,first_date=excluded.first_date,last_date=excluded.last_date,fact_digest=excluded.fact_digest,updated_at=excluded.updated_at`)
        .bind(f.sourceType,f.termStart,f.fileId,f.catalogKey,f.seniority,f.title,f.rawValue,f.location,f.allDay,f.timeLabel,f.startTime,f.endTime,first,last,f.factDigest,now));
    } else throw Error('Unknown compact fact kind.');
  }
  if(!await reserveRosterMaintenanceBudget(db,statements.length*6+64,256)) return {deferred:true};
  await db.batch([inactiveGuard(db,job),...statements,
    db.prepare('INSERT INTO roster_import_compact_receipts(run_id,batch_index,plan_revision,payload_hash) VALUES(?,?,?,?)').bind(job.run_id,batch.index,job.plan_revision,descriptor.hash),
  ]);
  return {preparedCompactBatch:batch.index+1};
}
function inactiveGuard(db,job) {
  return db.prepare(`SELECT CASE WHEN EXISTS(SELECT 1 FROM roster_files WHERE id=? AND active=0)
    AND EXISTS(SELECT 1 FROM roster_import_jobs WHERE run_id=? AND plan_revision=? AND compact_ready=0 AND rebase_owner='')
    THEN 1 ELSE json('compact-staging-conflict') END`).bind(job.file_id,job.run_id,job.plan_revision);
}
