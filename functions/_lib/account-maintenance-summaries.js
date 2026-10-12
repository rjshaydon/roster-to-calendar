import {maintenanceReceiptSummary} from './maintenance-receipt-summary.js';
import inventory from './d1-account-inventory.js';
// Receipt costs are lower bounds after removing the conservative settlement
// allowance. Only confirmed costs settled before the same analytics cutoff
// are excluded from recurring forecasts. Actual account usage is never reduced.
export async function accountMaintenanceSummaries(context,analytics,primary,day,previousDay){
 const result={outstanding:{...primary.outstanding},settled:{...primary.settled}};
 if(context.env.ACCOUNT_MAINTENANCE_PREVIEW_SUMMARY_ENABLED!=='true')return result;
 const raw=context.env.ROSTER_MAINTENANCE_PREVIEW_DB;
 const known=inventory.databases.find(d=>d.name==='roster-identity-batch2-preview');
 if(!raw||!known)throw Error('Preview maintenance accounting unavailable.');
 const db=context.data?.d1Meter?.additionalBinding(raw)||raw;
 const summary=await maintenanceReceiptSummary(db,day,previousDay,analytics.interval.start,analytics.interval.observedUntil);
 const measured=analytics.databases.find(d=>d.databaseId===known.id)||{rowsRead:0,rowsWritten:0};
 for(const [key,field]of [['reads','rowsRead'],['writes','rowsWritten']]){
  if(summary.settled[key]>Number(measured[field]))throw Error('Preview maintenance costs have not reconciled with account analytics.');
  result.settled[key]+=summary.settled[key];result.outstanding[key]+=summary.outstanding[key];
 }
 return result;
}
