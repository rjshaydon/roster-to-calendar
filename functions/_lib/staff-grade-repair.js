import {loadCachedSnapshot,storeCachedSnapshot,australianTermEndForStart} from './d1-calendar.js';
import {knownRosterGrade,resolveRosterGrades} from './roster-grade-evidence.js';
import {historyDigest,staffHistorySourceTerms} from './staff-history-cache.js';
import {identityMaintenancePreflight} from './identity-maintenance-preflight.js';
import {reserveRosterMaintenanceBudget} from './roster-maintenance-budget.js';
import {publishFacilityStaffMetadata} from './facility-overview-cache.js';

// Eight indexed person/file reads per checkpoint. Original shifts are never
// changed. Durable cursor + dirty marker make lost responses safe to resume.
export async function repairStaffGrades(context,input){
 const {ROSTER_DB:db,ROSTER_FILES:r2}=context.env;
 const source=await staffHistorySourceTerms(r2,input.sourceType);
 const term=source.terms.find(t=>t.termStart===input.termStart);
 if(!term||term.staffRevision!==input.staffRevision)return {status:'source-repaired'};
 const key=`staff-history/v1/grade-repair-v1/${term.sourceType}/${term.termStart}/${term.staffRevision}.json`;
 const previous=await r2.get(key),state=previous?await previous.json():{cursor:['',''],changed:0};
 if(state.done)return {status:'complete',changed:0};
 const preflight=await identityMaintenancePreflight(db,{accountAware:context.env.ROSTER_ACCOUNT_BUDGET_ENABLED==='true'});
 if(preflight.active_import||Number(preflight.unfinished_count)>=3||!Number(preflight.current_grant))return {status:'deferred'};
 // Include expiry-cache deletion triggers and their index writes. Reject an
 // unexpectedly large cache before reserving; no unbounded delete estimate.
 const sessions=(await db.prepare('SELECT subject_email FROM facility_access_sessions LIMIT 513').all()).results;
 if(sessions.length>512||!await reserveRosterMaintenanceBudget(db,4096,65536))return {status:'deferred'};
 const rows=(await db.prepare(`SELECT c.* FROM facility_term_staff_contributions c JOIN roster_files f ON f.id=c.file_id AND f.active=1
  WHERE c.source_type=? AND c.term_start=? AND (c.doctor_key,c.file_id)>(?,?) AND lower(trim(c.seniority)) IN ('','unknown','not recorded')
  ORDER BY c.doctor_key,c.file_id LIMIT 9`).bind(term.sourceType,term.termStart,...state.cursor).all()).results;
 const page=rows.slice(0,8),end=australianTermEndForStart(term.termStart);
 if(page.length){
  const evidence=await db.batch(page.map(row=>db.prepare(`SELECT seniority FROM roster_events INDEXED BY idx_roster_events_file_doctor
   WHERE file_id=? AND doctor_key=? AND start_date>=? AND start_date<=? LIMIT 513`).bind(row.file_id,row.doctor_key,term.termStart,end)));
  const updates=[];
  for(let i=0;i<page.length;i++){
   if(evidence[i].results.length>512)throw Error('Term grade evidence exceeds bounded checkpoint.');
   const grade=resolveRosterGrades(evidence[i].results.map(r=>r.seniority));
   if(!knownRosterGrade(grade))continue;
   const row=page[i],fact={sourceType:row.source_type,termStart:row.term_start,doctorKey:row.doctor_key,fileId:row.file_id,displayName:row.display_name,seniority:grade,membershipSource:row.membership_source,providerStaffId:row.provider_staff_id,firstApplicableDate:row.first_applicable_date,lastApplicableDate:row.last_applicable_date};
   updates.push(db.prepare(`UPDATE facility_term_staff_contributions SET seniority=?,fact_digest=?,updated_at=?
    WHERE source_type=? AND term_start=? AND doctor_key=? AND file_id=? AND fact_digest=? AND lower(trim(seniority)) IN ('','unknown','not recorded') AND EXISTS(SELECT 1 FROM roster_files WHERE id=? AND active=1)`)
    .bind(grade,await historyDigest(fact),new Date().toISOString(),row.source_type,row.term_start,row.doctor_key,row.file_id,row.fact_digest,row.file_id));
  }
  // Persist dirty before D1: an interrupted/lost response must still republish.
  if(updates.length){state.changed+=updates.length;await r2.put(key,JSON.stringify(state));await db.batch(updates);}
  state.cursor=[page.at(-1).doctor_key,page.at(-1).file_id];
 }
 if(rows.length>8){await r2.put(key,JSON.stringify(state));return {status:'continue'};}
 if(state.changed){
  // Publish only this term's staff metadata. Existing day/calendar artifacts
  // and source files do not need regenerating for a summary-grade correction.
  await publishFacilityStaffMetadata(context,[term.sourceType],{termStart:term.termStart});
  const refreshed=(await staffHistorySourceTerms(r2,term.sourceType)).terms.find(t=>t.termStart===term.termStart);
  if(!refreshed||refreshed.staffRevision===term.staffRevision)throw Error('Grade publication not yet visible.');
  await r2.put(`staff-history/v1/grade-repair-v1/${term.sourceType}/${term.termStart}/${refreshed.staffRevision}.json`,JSON.stringify({done:true,changed:0}));
 }
 state.done=true;await r2.put(key,JSON.stringify(state));
 return {status:state.changed?'source-repaired':'complete',changed:state.changed};
}
