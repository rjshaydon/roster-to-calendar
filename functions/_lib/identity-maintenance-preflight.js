import {activeRosterImportsQuery} from './roster-delivery-health.js';

// Independent indexed admission reads share one native D1 call. This is only
// a preflight: the engine lease and budget reservation still commit atomically.
export async function identityMaintenancePreflight(db,{runId='',weekKey='',accountAware=false,now=new Date()}={}) {
 const instant=now.toISOString(),day=instant.slice(0,10);
 const active=activeRosterImportsQuery(['monash-adults','monash-paeds','vhh-active-medical-roster','dandenong-findmyshift']);
 const bindings=[...(weekKey?[weekKey]:[]),...active.bindings,...(runId?[runId,instant]:[]),...(accountAware?[day,day,instant]:[])];
 const result=await db.prepare(`SELECT
  ${weekKey?"(SELECT run_id FROM roster_identity_audit_runs WHERE week_key=? AND status='complete' AND mode='audit' LIMIT 1)":'NULL'} AS completed_run,
  (${active.sql}) AS active_import,
  ${runId?"(SELECT run_id FROM roster_identity_audit_runs WHERE run_id=? AND status='running' AND lease_until>? LIMIT 1)":'NULL'} AS live_lease,
  ${accountAware?"(SELECT COUNT(*) FROM (SELECT request_id FROM roster_maintenance_receipts WHERE utc_day=? AND reconciled_at='' AND (finished_at='' OR metadata_complete<>1) LIMIT 3))":'0'} AS unfinished_count,
  ${accountAware?"EXISTS(SELECT 1 FROM roster_account_budget WHERE utc_day=? AND valid_until>? AND stop_reason='' LIMIT 1)":'1'} AS current_grant`)
  .bind(...bindings).first();
 if(!result || !Number.isInteger(Number(result.unfinished_count)) || ![0,1].includes(Number(result.current_grant))) throw Error('Identity admission inspection failed.');
 return result;
}
