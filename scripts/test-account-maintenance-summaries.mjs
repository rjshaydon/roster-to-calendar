import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {readFile} from 'node:fs/promises';
import {accountMaintenanceSummaries} from '../functions/_lib/account-maintenance-summaries.js';
import {createD1Meter} from '../functions/_middleware.js';
import {accountMaintenanceHeadroom} from '../functions/_lib/account-maintenance-policy.js';
const sqlite=new DatabaseSync(':memory:');sqlite.exec('CREATE TABLE roster_import_daily_budget(utc_day TEXT PRIMARY KEY,reserved_reads INTEGER,reserved_writes INTEGER)');
for(const file of ['0035_roster_account_budget.sql','0041_maintenance_recovery.sql'])sqlite.exec(await readFile('migrations/'+file,'utf8'));
const cutoff='2026-10-01T04:45:00.000Z',day='2026-10-01',previous='2026-09-30';
const insert=sqlite.prepare('INSERT INTO roster_maintenance_receipts(request_id,utc_day,reserved_reads,reserved_writes,actual_reads,actual_writes,finished_at,metadata_complete)VALUES(?,?,?,?,?,?,?,?)');
insert.run('settled',day,2000,15000,1024,14024,'2026-10-01T04:00:00.000Z',1);
insert.run('unknown',day,30,40,0,0,'',0);
insert.run('late',day,200,200,100,50,'2026-10-01T04:50:00.000Z',1);
const raw={prepare(sql){
 assert.match(sql,/^WITH today/,'auxiliary binding is used only for the fixed read-only summary');
 return {bind(...values){return {
  async all(){return {results:sqlite.prepare(sql).all(...values),meta:{rows_read:3,rows_written:0}};},
  async first(){return sqlite.prepare(sql).get(...values);}
 };}};
}};
const meter=createD1Meter(raw,4,{firstViaAll:true});
const context={env:{ACCOUNT_MAINTENANCE_PREVIEW_SUMMARY_ENABLED:'true',ROSTER_MAINTENANCE_PREVIEW_DB:raw},data:{d1Meter:meter}};
const primary={settled:{reads:100,writes:3000},outstanding:{reads:10,writes:20}};
const analytics={complete:true,interval:{start:day+'T00:00:00.000Z',observedUntil:cutoff,end:'2026-10-02T00:00:00.000Z'},totals:{rowsRead:20000,rowsWritten:20000},databases:[{databaseId:'741ce534-eafe-4aa7-9305-7b7ed5661979',rowsRead:1000,rowsWritten:14000}]};
const summary=await accountMaintenanceSummaries(context,analytics,primary,day,previous);
assert.deepEqual(summary,{settled:{reads:1100,writes:17000},outstanding:{reads:140,writes:110}});
assert.equal(meter.statementCount,1);assert.equal(meter.rowsRead,3);assert.equal(meter.rowsWritten,0);assert.equal(meter.metadataComplete,true);
assert.equal(accountMaintenanceHeadroom(analytics,primary.outstanding,'2026-10-01T05:00:00Z',primary.settled).writes,0);
assert(accountMaintenanceHeadroom(analytics,summary.outstanding,'2026-10-01T05:00:00Z',summary.settled).writes>40000,'only verified settled one-off costs leave the ordinary forecast');
assert.deepEqual(analytics.totals,{rowsRead:20000,rowsWritten:20000},'actual account usage never changes');
await assert.rejects(accountMaintenanceSummaries(context,{...analytics,databases:[]},primary,day,previous),/not reconciled/);
await assert.rejects(accountMaintenanceSummaries({...context,env:{ACCOUNT_MAINTENANCE_PREVIEW_SUMMARY_ENABLED:'true'}},analytics,primary,day,previous),/unavailable/);
assert.deepEqual(await accountMaintenanceSummaries({env:{}},analytics,primary,day,previous),{settled:primary.settled,outstanding:primary.outstanding});
assert.equal(sqlite.prepare('SELECT COUNT(*) n FROM roster_maintenance_receipts').get().n,3,'unknown and completed receipts remain untouched');
console.log('Account-wide summaries passed verified per-database settlement, unknown/late reservations, shared read metering, read-only access, unchanged actual usage and fail-closed evidence.');
