import {accountMaintenanceHeadroom,ACCOUNT_MAINTENANCE_POLICY} from '../functions/_lib/account-maintenance-policy.js';
import {maintenanceReceiptSummary} from '../functions/_lib/maintenance-receipt-summary.js';
import {accountMaintenanceSummaries} from '../functions/_lib/account-maintenance-summaries.js';
import {publicationHeldReads,publicationHeldWrites} from '../functions/_lib/publication-delivery-budget.js';

// Read-only: never recovers/refunds receipts, updates a grant or stops the app.
// Any GO is an advisory estimate; production still atomically admits each step.
export async function maintenanceDiagnosticEvidence(db,previewDb,analytics,{previewEnabled=false}={}) {
 const day=analytics.interval.date,previous=new Date(Date.parse(analytics.interval.start)-86400000).toISOString().slice(0,10);
 const primary=await maintenanceReceiptSummary(db,day,previous,analytics.interval.start,analytics.interval.observedUntil);
 const summary=await accountMaintenanceSummaries({env:{ACCOUNT_MAINTENANCE_PREVIEW_SUMMARY_ENABLED:String(previewEnabled),ROSTER_MAINTENANCE_PREVIEW_DB:previewDb}},analytics,primary,day,previous);
 const holds=await db.prepare(`SELECT ${publicationHeldReads} AS reads,${publicationHeldWrites} AS writes`).first();
 const grant=await db.prepare('SELECT utc_day,valid_until,stop_reason,maximum_reads,maximum_writes,allocated_reads,allocated_writes FROM roster_account_budget WHERE utc_day=?').bind(day).first();
 return {...summary,holds,grant};
}
export function evaluateMaintenanceDiagnostic(analytics,evidence,now=new Date(),estimate={}) {
 const reasons=[];
 let budget=null;
 try {
  if(!evidence||!evidence.outstanding||!evidence.settled||!evidence.holds)throw Error('Maintenance accounting evidence is required.');
  for(const group of [evidence.outstanding,evidence.settled,evidence.holds])for(const k of ['reads','writes'])if(!Number.isSafeInteger(Number(group[k]))||Number(group[k])<0)throw Error('Maintenance accounting evidence is invalid.');
  for(const k of ['reads','writes'])if(evidence.settled[k]>analytics.totals[k==='reads'?'rowsRead':'rowsWritten'])throw Error('Settled maintenance exceeds account analytics.');
  budget=accountMaintenanceHeadroom(analytics,evidence.outstanding,now,evidence.settled);
  if(evidence.grant?.stop_reason)reasons.push(evidence.grant.stop_reason);
  const holds=evidence.holds;
  const bulk={reads:Math.max(0,budget.reads-Number(holds.reads)),writes:Math.max(0,budget.writes-Number(holds.writes))};
  const requested={reads:Number(estimate.rowsRead||0),writes:Number(estimate.rowsWritten||0)};
  for(const k of ['reads','writes'])if(!Number.isSafeInteger(requested[k])||requested[k]<0)throw Error('A non-negative integer estimate is required.');
  if(bulk.reads<Math.max(1,requested.reads))reasons.push('insufficient-uncommitted-reads');
  if(bulk.writes<Math.max(1,requested.writes))reasons.push('insufficient-uncommitted-writes');
  return {decision:reasons.length?'STOP':'GO',advisory:true,reasons,policy:ACCOUNT_MAINTENANCE_POLICY,budget,bulkHeadroom:bulk,
   publicationHolds:holds,outstanding:evidence.outstanding,settledMaintenance:evidence.settled,
   currentGrant:evidence.grant||null,optionalEstimate:requested,note:'Read-only estimate. Production independently reserves each request and protects publication capacity; this report grants no capacity.'};
 } catch(error){return {decision:'STOP',advisory:true,reasons:[error.message],policy:ACCOUNT_MAINTENANCE_POLICY,budget};}
}
