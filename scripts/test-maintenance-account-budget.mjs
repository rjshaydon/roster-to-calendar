import assert from "node:assert/strict";
import { DatabaseSync } from "node:sqlite";
import { readFile } from "node:fs/promises";
import { outstandingMaintenance, accountMaintenanceHeadroom } from "../functions/_lib/account-maintenance-policy.js";
import { beginMaintenanceAccounting, reserveRosterMaintenanceBudget, finishMaintenanceAccounting, recoverAbandonedMaintenance, optionalMaintenanceAvailable } from "../functions/_lib/roster-maintenance-budget.js";
import { createD1Meter, onRequest as middleware } from "../functions/_middleware.js";
import { onRequestPost as admit } from "../functions/api/automation/account-budget.js";

const { default: inventory } = await import("../functions/_lib/d1-account-inventory.js");
assert.deepEqual(inventory, JSON.parse(await readFile("config/d1-database-inventory.json", "utf8")), "runtime inventory must match the verified account inventory");
const cutoff = "2026-10-01T04:45:00.000Z";
assert.deepEqual(outstandingMaintenance([
  { metadata_complete: 1, finished_at: "2026-10-01T04:00:00.000Z", actual_reads: 100, actual_writes: 10 },
  { metadata_complete: 1, finished_at: "2026-10-01T04:50:00.000Z", actual_reads: 200, actual_writes: 20 },
  { metadata_complete: 0, reserved_reads: 300, reserved_writes: 30 },
], cutoff), { reads: 500, writes: 50 }, "settled work must not be counted twice; unknown work stays reserved");
const analytics = { complete: true, interval: { start: "2026-10-01T00:00:00.000Z", end: "2026-10-02T00:00:00.000Z", observedUntil: cutoff }, totals: { rowsRead: 1000000, rowsWritten: 20000 }, fiveMinuteBuckets: [] };
assert.equal(accountMaintenanceHeadroom(analytics, { reads: 500000, writes: 5000 }, "2026-10-01T05:00:00Z", { reads: 900000, writes: 19000 }).writes > 10000, true, "measured one-off import costs must not be forecast as recurring traffic");
assert.equal(accountMaintenanceHeadroom({ ...analytics, totals: { rowsRead: 4000000, rowsWritten: 80000 } }, { reads: 1, writes: 1 }, "2026-10-01T05:00:00Z").reads, 0);
assert.throws(() => accountMaintenanceHeadroom(analytics, { reads: 0, writes: 0 }, "2026-10-01T06:00:00Z"), /stale/);
assert.equal(accountMaintenanceHeadroom({ ...analytics, fiveMinuteBuckets: [{ observedAt: "2026-10-01T04:40:00Z", rowsRead: 1000000, rowsWritten: 10000 }] }, { reads: 0, writes: 0 }, "2026-10-01T05:00:00Z").reads, 0, "projected traffic requires additional headroom");

const sqlite = new DatabaseSync(":memory:");
sqlite.exec("CREATE TABLE IF NOT EXISTS roster_import_daily_budget (utc_day TEXT PRIMARY KEY,reserved_writes INTEGER NOT NULL DEFAULT 0,reserved_reads INTEGER NOT NULL DEFAULT 0)");
sqlite.exec(await readFile("migrations/0035_roster_account_budget.sql", "utf8"));
sqlite.exec(await readFile("migrations/0041_maintenance_recovery.sql", "utf8"));
function database() {
  const db = { prepare(sql) { return { args: [], bind(...args) { this.args = args; return this; },
    async run() { const r = sqlite.prepare(sql).run(...this.args); return { meta: { changes: Number(r.changes), rows_read: Number(r.changes)+1, rows_written: Number(r.changes) } }; },
    async first() { return sqlite.prepare(sql).get(...this.args) || null; },
    async all() { const results = sqlite.prepare(sql).all(...this.args); return { results, meta: { rows_read: results.length+1, rows_written: 0 } }; }
  }; }, async batch(statements) { sqlite.exec("BEGIN"); try { const results=[]; for (const s of statements) results.push(await s.run()); sqlite.exec("COMMIT"); return results; } catch(e) { sqlite.exec("ROLLBACK"); throw e; } } };
  return db;
}
const day = new Date().toISOString().slice(0,10);
sqlite.prepare("INSERT INTO roster_account_budget(utc_day,maximum_reads,maximum_writes,valid_until) VALUES(?,?,?,?)").run(day, 1000000, 20000, new Date(Date.now()+600000).toISOString());
const first = database(); beginMaintenanceAccounting(first, "first", true);
assert.equal(await reserveRosterMaintenanceBudget(first, 15000, 100), true, "legitimate writes are no longer capped at 10000");
const second = database(); beginMaintenanceAccounting(second, "second", true);
assert.equal(await reserveRosterMaintenanceBudget(second, 5000, 100), false, "concurrent jobs share the same SQL ceiling");
assert.equal(sqlite.prepare("SELECT COUNT(*) n FROM roster_maintenance_receipts").get().n, 1);
assert.equal(sqlite.prepare("SELECT reserved_writes FROM roster_import_daily_budget").get().reserved_writes, 15048);
await finishMaintenanceAccounting(first, { metadataComplete: true, rowsRead: 100, rowsWritten: 500 });
assert.equal(sqlite.prepare("SELECT actual_writes FROM roster_maintenance_receipts").get().actual_writes, 524);
assert.equal(sqlite.prepare("SELECT allocated_writes FROM roster_account_budget").get().allocated_writes, 524, "unused measured reservations release the grant immediately");
await finishMaintenanceAccounting(first, { metadataComplete: true, rowsRead: 100, rowsWritten: 500 });
assert.equal(sqlite.prepare("SELECT allocated_writes FROM roster_account_budget").get().allocated_writes, 524, "settlement replay cannot refund twice");
sqlite.prepare("UPDATE roster_account_budget SET valid_until=''").run();
assert.equal(await reserveRosterMaintenanceBudget(second, 1, 1), false, "expired analytics cannot admit work");
sqlite.prepare("UPDATE roster_account_budget SET valid_until=?,maximum_writes=50000").run(new Date(Date.now()+600000).toISOString());
assert.equal(await reserveRosterMaintenanceBudget(second, 10, 10), true);
await finishMaintenanceAccounting(second, { metadataComplete: false, rowsRead: 0, rowsWritten: 0 });
assert.equal(sqlite.prepare("SELECT metadata_complete FROM roster_maintenance_receipts WHERE request_id='second'").get().metadata_complete, 0);
const third = database(); beginMaintenanceAccounting(third, "third", true);
assert.equal(await reserveRosterMaintenanceBudget(third, 1, 1), true);
await finishMaintenanceAccounting(third, { metadataComplete: true, rowsRead: 1000, rowsWritten: 1000 });
assert.match(sqlite.prepare("SELECT stop_reason FROM roster_account_budget").get().stop_reason, /cost-overrun/);
assert.equal(await reserveRosterMaintenanceBudget(second, 1, 1), false);

const env = { ROSTER_DB: database(), ROSTER_AUTOMATION_TOKEN: "test-only", ROSTER_ACCOUNT_BUDGET_ENABLED: "true", ROSTER_ACCOUNT_ANALYTICS_TOKEN: "test-only" };
const context = { env, request: new Request("https://test/api/automation/account-budget", { method: "POST", headers: { Authorization: "Bearer test-only" } }) };
const originalFetch = globalThis.fetch;
try {
  globalThis.fetch = async () => { throw new Error("injected analytics failure"); };
  assert.equal((await (await admit(context)).json()).deferred, true);
  assert.equal(await reserveRosterMaintenanceBudget(second, 1, 1), false, "analytics failure closes admission");
  assert.match(sqlite.prepare("SELECT stop_reason FROM roster_account_budget").get().stop_reason, /cost-overrun/, "analytics failure cannot erase a cost overrun");
  sqlite.prepare("UPDATE roster_account_budget SET stop_reason='',valid_until=''").run();
  const bucket = new Date(Date.now()-20*60*1000).toISOString();
  globalThis.fetch = async () => Response.json({ data: { viewer: { accounts: [{
    usage: [{ dimensions: { databaseId: "237d0d52-3a7c-4e02-8648-9f4dedbc1cb0" }, sum: { rowsRead: 100, rowsWritten: 10, readQueries: 1, writeQueries: 1 } }],
    timeline: [{ dimensions: { databaseId: "237d0d52-3a7c-4e02-8648-9f4dedbc1cb0", datetimeFiveMinutes: bucket }, sum: { rowsRead: 100, rowsWritten: 10, readQueries: 1, writeQueries: 1 } }], queries: []
  }] } } });
  const admission = await (await admit(context)).json();
  assert.equal(admission.deferred, false, JSON.stringify(admission));
  assert.ok(admission.budget.writes > 10000);
  assert.ok(sqlite.prepare("SELECT maximum_writes-allocated_writes AS headroom FROM roster_account_budget").get().headroom > 10000);
  // A reservation interleaving the snapshot's two reads must prevent stale
  // headroom being applied. The previous grant remains atomically consumed.
  const raceDb = database(), racer = database();
  beginMaintenanceAccounting(racer, "interleaved", true);
  const prepare = raceDb.prepare;
  raceDb.prepare = (sql) => {
    const statement = prepare(sql);
    if (sql.startsWith("SELECT allocated_reads")) {
      const first = statement.first;
      statement.first = async () => { await reserveRosterMaintenanceBudget(racer, 100, 100); return first.call(statement); };
    }
    return statement;
  };
  assert.equal((await (await admit({ ...context, env: { ...env, ROSTER_DB: raceDb } })).json()).deferred, true, "snapshot CAS rejects concurrent allocation");
} finally { globalThis.fetch = originalFetch; }
console.log("Account-aware maintenance accounting passed settled/unsettled usage, write headroom, concurrent reservations, stale admission and cost-overrun checks.");

// Lost D1 responses cannot be treated as zero cost and refunded after settlement.
const broken = createD1Meter({ prepare() { return { async run() { throw new Error("lost response"); } }; } }, 4);
await assert.rejects(() => broken.binding.prepare("write").run(), /lost response/);
assert.equal(broken.metadataComplete, false);

// HTTP middleware applies account-aware mode even without a caller session;
// receipt settlement still works after the route uses its statement allowance.
sqlite.prepare("UPDATE roster_account_budget SET stop_reason='',valid_until=?,maximum_reads=allocated_reads+1000000,maximum_writes=allocated_writes+50000").run(new Date(Date.now()+600000).toISOString());
const middlewareEnv = { ROSTER_DB: database(), ROSTER_ACCOUNT_BUDGET_ENABLED: "true", ROSTER_AUTOMATION_WRITES_ENABLED: "true" };
const originalLog = console.log;
try {
  console.log = () => {};
  const response = await middleware({ env: middlewareEnv,
    request: new Request("https://test/api/automation/facility-refresh", { method: "POST", body: "{}" }),
    async next() {
      assert.equal(await reserveRosterMaintenanceBudget(middlewareEnv.ROSTER_DB, 100, 100), true);
      await middlewareEnv.ROSTER_DB.prepare("SELECT utc_day FROM roster_account_budget LIMIT 1").first();
      return Response.json({ ok: true });
    }
  });
  assert.equal(response.status, 200);
  const latest = sqlite.prepare("SELECT * FROM roster_maintenance_receipts WHERE request_id NOT IN ('first','second','third','interleaved')").get();
  assert.equal(latest.metadata_complete, 1);
  assert.ok(latest.finished_at);
  assert.equal(middlewareEnv.ROSTER_DB.prepare("SELECT 1").constructor, Object);
} finally { console.log = originalLog; }
console.log("HTTP accounting and lost-response reservation retention checks passed.");

// Expiry is only eligible after a enforced execution bound AND a validated
// analytics cutoff. Historical requests have no bound and stay reserved.
sqlite.exec("DELETE FROM roster_maintenance_receipts; UPDATE roster_account_budget SET allocated_reads=9000,allocated_writes=9000,maximum_reads=1000000,maximum_writes=40000,stop_reason=''");
const receipt=sqlite.prepare("INSERT INTO roster_maintenance_receipts(request_id,utc_day,reserved_reads,reserved_writes,recover_after,finished_at,metadata_complete) VALUES(?,?,?,?,?,?,?)");
const recoveryCutoff=new Date(Date.now()-20*60000).toISOString();
receipt.run('legacy',day,1000,1000,'','',0);
receipt.run('bounded-abandoned',day,1000,1000,recoveryCutoff,'',0);
receipt.run('lost-result',day,1000,1000,recoveryCutoff,recoveryCutoff,0);
receipt.run('still-running',day,1000,1000,new Date(Date.now()+60000).toISOString(),'',0);
receipt.run('settled',day,1000,1000,recoveryCutoff,recoveryCutoff,1);
receipt.run('yesterday','2000-01-01',1000,1000,recoveryCutoff,'',0);
assert.equal(await optionalMaintenanceAvailable(database()),false,'three unknown requests pause optional work');
assert.equal(await recoverAbandonedMaintenance(database(),day,recoveryCutoff),2);
assert.equal(await recoverAbandonedMaintenance(database(),day,recoveryCutoff),0,'recovery replay cannot refund twice');
assert.equal(sqlite.prepare('SELECT allocated_writes FROM roster_account_budget').get().allocated_writes,7000);
assert.equal(sqlite.prepare("SELECT reconciled_at FROM roster_maintenance_receipts WHERE request_id='legacy'").get().reconciled_at,'');
assert.equal(sqlite.prepare("SELECT reconciled_at FROM roster_maintenance_receipts WHERE request_id='yesterday'").get().reconciled_at,'','midnight cannot refund the new day using old receipts');
assert.equal(sqlite.prepare("SELECT metadata_complete FROM roster_maintenance_receipts WHERE request_id='lost-result'").get().metadata_complete,0,'recovery never invents actual usage metadata');
assert.equal(await optionalMaintenanceAvailable(database()),true);
const late=database();beginMaintenanceAccounting(late,'bounded-abandoned',true);
// A late settlement for a recovered request cannot release another grant.
assert.equal(await reserveRosterMaintenanceBudget(late,1,1),true);
const beforeLate=sqlite.prepare('SELECT allocated_writes FROM roster_account_budget').get().allocated_writes;
await finishMaintenanceAccounting(late,{metadataComplete:true,rowsRead:0,rowsWritten:0});
assert.equal(sqlite.prepare('SELECT allocated_writes FROM roster_account_budget').get().allocated_writes,beforeLate);
const optional=database();beginMaintenanceAccounting(optional,'optional',true,{purpose:'identity'});
sqlite.prepare('UPDATE roster_account_budget SET maximum_writes=allocated_writes+20000').run();
assert.equal(await reserveRosterMaintenanceBudget(optional,512,8192),false,'optional maintenance preserves routine roster headroom');
const roster=database();beginMaintenanceAccounting(roster,'routine-roster',true,{purpose:'roster',deadlineAt:Date.now()+120000,statementLimit:96});
assert.equal(await reserveRosterMaintenanceBudget(roster,512,8192),true,'routine roster work can use its protected headroom');
assert.ok(sqlite.prepare("SELECT recover_after FROM roster_maintenance_receipts WHERE request_id='routine-roster'").get().recover_after);
let underlyingCalls=0;
const expired=createD1Meter({prepare(){return {async run(){underlyingCalls++;},async all(){underlyingCalls++;}};},async batch(){underlyingCalls++;}},96,{deadlineAt:Date.now()-1});
await assert.rejects(()=>expired.binding.prepare('SELECT 1').run(),e=>e.code==='d1-request-deadline-exceeded');
await assert.rejects(()=>expired.binding.batch([expired.binding.prepare('SELECT 1')]),e=>e.code==='d1-request-deadline-exceeded');
assert.equal(underlyingCalls,0,'expired requests cannot issue more SQL');
console.log('Bounded recovery, legacy retention, midnight isolation, replay, late settlement, circuit breaker and protected roster headroom passed.');

// Cross-midnight in-flight batches remain covered in the new day's headroom,
// while reconciliation can only release the original day's allocated grant.
// Keep today's unrelated fixtures outside this simulated cutoff, regardless
// of the real time of day at which this test is run.
sqlite.prepare("UPDATE roster_maintenance_receipts SET recover_after=? WHERE utc_day=? AND reconciled_at='' AND recover_after<>''").run(day+'T23:59:00.000Z',day);
const RealDate=Date;
let instant=RealDate.parse(day+'T00:30:00.000Z');
const previousDay=new RealDate(instant-86400000).toISOString().slice(0,10);
try {
 globalThis.Date=class extends RealDate {constructor(...args){super(...(args.length?args:[instant]));}static now(){return instant;}};
 globalThis.fetch=async()=>Response.json({data:{viewer:{accounts:[{
  usage:[{dimensions:{databaseId:inventory.databases[0].id||'237d0d52-3a7c-4e02-8648-9f4dedbc1cb0'},sum:{rowsRead:100,rowsWritten:10,readQueries:1,writeQueries:1}}],
  timeline:[{dimensions:{databaseId:'237d0d52-3a7c-4e02-8648-9f4dedbc1cb0',datetimeFiveMinutes:new Date(instant-20*60000).toISOString()},sum:{rowsRead:100,rowsWritten:10,readQueries:1,writeQueries:1}}],queries:[]
 }]}}});
 const before=await (await admit(context)).json();
 assert.equal(before.deferred,false,JSON.stringify(before));
 receipt.run('cross-midnight',previousDay,600,700,day+'T01:00:00.000Z','',0);
 sqlite.prepare('INSERT INTO roster_account_budget(utc_day,allocated_reads,allocated_writes) VALUES(?,600,700)').run(previousDay);
 const held=await (await admit(context)).json();
 assert.equal(held.outstanding.reads-before.outstanding.reads,600);
 assert.equal(held.outstanding.writes-before.outstanding.writes,700);
 const todayAllocated=sqlite.prepare('SELECT allocated_writes FROM roster_account_budget WHERE utc_day=?').get(day).allocated_writes;
 instant=RealDate.parse(day+'T01:25:00.000Z');
 const recovered=await (await admit(context)).json();
 assert.equal(recovered.recovered,1,JSON.stringify(recovered));
 assert.equal(sqlite.prepare('SELECT allocated_writes FROM roster_account_budget WHERE utc_day=?').get(previousDay).allocated_writes,0);
 assert.equal(sqlite.prepare('SELECT allocated_writes FROM roster_account_budget WHERE utc_day=?').get(day).allocated_writes,todayAllocated,'yesterday cannot refund today');
}finally{globalThis.Date=RealDate;globalThis.fetch=originalFetch;}
console.log('Cross-midnight unknown work remains reserved until settled analytics covers its execution bound.');
