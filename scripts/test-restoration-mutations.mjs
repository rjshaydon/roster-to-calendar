import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { onRequest as middleware } from '../functions/_middleware.js';
import { onRequestPost as stateHandler } from '../functions/api/state.js';
import { DatabaseSync } from 'node:sqlite';
import { readFile, readdir } from 'node:fs/promises';
import { makeSessionPatch, applySessionPatch } from '../public/static/session-settings-patch.js';
import { saveSessionSettings } from '../functions/_lib/session-settings.js';
import { planRosterImportBatches } from '../functions/_lib/roster-import-batches.js';
import { handleManualRosterImport, deactivateManualRosterFiles } from '../functions/_lib/manual-roster-management.js';
import { beginMaintenanceAccounting } from '../functions/_lib/roster-maintenance-budget.js';
import { executeBoundedRosterImport } from './roster-import-driver.mjs';
import { loadPublishedDoctorCalendar } from '../functions/_lib/published-doctor-calendar.js';
import { storeCachedSnapshot } from '../functions/_lib/d1-calendar.js';
class LocalD1 {
  constructor(sqlite) { this.sqlite = sqlite; this.rowsWritten = 0; this.sql = []; this.failRunIncludes = ""; }
  prepare(sql) {
    const owner = this;
    owner.sql.push(sql.replace(/\s+/g, " ").trim());
    return {
      args: [],
      bind(...args) { this.args = args; return this; },
      async run() {
        if (owner.failRunIncludes && sql.includes(owner.failRunIncludes)) {
          owner.failRunIncludes = "";
          throw new Error("Injected D1 statement failure");
        }
        const result = owner.sqlite.prepare(sql).run(...this.args);
        owner.rowsWritten += Number(result.changes || 0);
        return { success: true, meta: { changes: Number(result.changes || 0), rows_read: Number(result.changes || 0) + 1, rows_written: Number(result.changes || 0) } };
      },
      // Metadata is simulated here to test reservation/refund logic, not to
      // claim SQLite result counts measure Cloudflare index billing.
      async all() { const results = owner.sqlite.prepare(sql).all(...this.args); return { success: true, results, meta: { rows_read: results.length + 1, rows_written: 0 } }; },
      async first() { return owner.sqlite.prepare(sql).get(...this.args) || null; },
    };
  }
  async batch(statements) {
    if (this.beforeBatch) { const fn = this.beforeBatch; this.beforeBatch = null; fn(statements); }
    const results = [];
    const ownsTransaction = !this.sqlite.isTransaction;
    if (ownsTransaction) this.sqlite.exec("BEGIN");
    try {
      for (const statement of statements) results.push(await statement.run());
      if (ownsTransaction) this.sqlite.exec("COMMIT");
      return results;
    } catch (error) {
      if (ownsTransaction) this.sqlite.exec("ROLLBACK");
      throw error;
    }
  }
}

class LocalR2 {
  constructor() { this.objects = new Map(); this.puts = 0; this.gets = 0; this.version = 0; this.failPointerOnce = false; }
  async put(key, value, options = {}) {
    if (this.failPointerOnce && key.endsWith("/manifest.json")) { this.failPointerOnce = false; throw new Error("Injected manifest failure"); }
    const current = this.objects.get(key);
    if (options.onlyIf?.etagMatches && current?.etag !== options.onlyIf.etagMatches) throw new Error("Precondition failed");
    if (options.onlyIf?.etagDoesNotMatch === "*" && current) throw new Error("Precondition failed");
    this.puts += 1;
    const bytes = value instanceof Uint8Array ? value : new Uint8Array(value);
    this.version += 1;
    this.objects.set(key, { bytes, options, etag: `local-${this.version}` });
  }
  async get(key) {
    this.gets += 1;
    const item = this.objects.get(key);
    if (!item) return null;
    return { etag: item.etag, arrayBuffer: async () => item.bytes.buffer.slice(item.bytes.byteOffset, item.bytes.byteOffset + item.bytes.byteLength) };
  }
}


const sqlite = new DatabaseSync(':memory:');
for (const name of (await readdir(new URL('../migrations', import.meta.url))).filter(name => name.endsWith('.sql')).sort()) sqlite.exec(await readFile(new URL(`../migrations/${name}`, import.meta.url),'utf8'));
const db = new LocalD1(sqlite);
const base = { settings: { includeLocations: true, includeLeave: true } };
const left = makeSessionPatch(base, {settings:{...base.settings,includeLeave:false}});
const right = makeSessionPatch(base, {settings:{...base.settings,includeLocations:false}});
assert.deepEqual(applySessionPatch(applySessionPatch(base,left),right).settings,{includeLocations:false,includeLeave:false});
assert.deepEqual(applySessionPatch(applySessionPatch(base,left),left).settings,{includeLocations:true,includeLeave:false});
assert.throws(()=>applySessionPatch({settings:{includeLeave:'remote'}},left),/another device/);
assert.throws(()=>applySessionPatch({settings:null},left),/another device/);
assert.throws(()=>makeSessionPatch({},JSON.parse('{"settings":{"__proto__":{}}}')),/Invalid/);
assert.equal(makeSessionPatch({customEvents:[{id:'a'},{id:'b'}]},{customEvents:[{id:'b'},{id:'a'}]}).length,0);
const email='fixture@example.test', sanitizeCustomEvents = events => events || [];
const save = changes => saveSessionSettings(db,{email,changes,sanitizeCustomEvents});
sqlite.prepare('INSERT INTO account_states(email,session_json) VALUES(?,?)').run(email,JSON.stringify(base));
await save(left); await save(right); await save(right);
assert.deepEqual(JSON.parse(sqlite.prepare('SELECT session_json FROM account_states WHERE email=?').get(email).session_json).settings,{includeLocations:false,includeLeave:false});
const beforeRace = JSON.parse(sqlite.prepare('SELECT session_json FROM account_states WHERE email=?').get(email).session_json);
db.beforeBatch = () => sqlite.prepare("UPDATE account_states SET settings_revision='concurrent',session_json=? WHERE email=?").run(JSON.stringify({...beforeRace,settings:{...beforeRace.settings,includeLeave:'remote'}}),email);
await assert.rejects(save(makeSessionPatch(beforeRace,{...beforeRace,settings:{...beforeRace.settings,includeLeave:true}})),/during this save/);
assert.equal(JSON.parse(sqlite.prepare('SELECT session_json FROM account_states WHERE email=?').get(email).session_json).settings.includeLeave,'remote');
assert.equal(db.sql.some(sql=>/roster_events|roster_daily_presence/.test(sql)),false,'settings must not query roster history');
const custom={id:'custom-a',title:'Own event',startDate:'2026-10-01',endDate:'2026-10-01',allDay:true,include:true};
const stored=JSON.parse(sqlite.prepare('SELECT session_json FROM account_states WHERE email=?').get(email).session_json);
db.failRunIncludes='INSERT INTO custom_events';
await assert.rejects(save(makeSessionPatch(stored,{...stored,customEvents:[custom]})),/Injected/);
assert.equal(sqlite.prepare('SELECT COUNT(*) AS n FROM custom_events').get().n,0);
assert.equal(sqlite.prepare('SELECT session_json FROM account_states WHERE email=?').get(email).session_json,JSON.stringify(stored));
console.log('Settings: independent edits, conflicts, replay, unsafe paths, CAS races and rollback passed.');

const day=new Date().toISOString().slice(0,10);
sqlite.prepare('INSERT INTO roster_account_budget(utc_day,maximum_reads,maximum_writes,valid_until) VALUES(?,?,?,?)').run(day,50000000,5000000,new Date(Date.now()+3600000).toISOString());
const context={env:{ROSTER_DB:db,FACILITY_AUTOMATIC_SOURCES:'mmc,mch,ddh,vhh',FACILITY_AUTOMATIC_PUBLICATION_ENABLED:'true'}};
let request=0;
const call=async(file,body)=>{beginMaintenanceAccounting(db,'fixture-'+(++request),true);return handleManualRosterImport(context,{file,...body},email)};
const doctor={key:'FIXTURE',displayName:'Fixture',seniority:'Registrar'};
async function plan(from,to,count=1600) {
  const file={id:'manual:'+crypto.randomUUID(),name:'Fixture.xlsx',sourceType:'mmc',sourceId:'manual-mmc',parserVersion:'fixture'};
  const events=Array.from({length:count},(_,i)=>({id:file.id+':'+i,source:'MMC',title:'Day',rawValue:'Day',seniority:'Registrar',start:`${i===count-1?to:from}T08:00:00`,end:`${i===count-1?to:from}T16:00:00`}));
  return {file,plan:await planRosterImportBatches({file,doctors:[doctor],eventsByDoctor:{FIXTURE:events},issuesByDoctor:{}})};
}
const old=await plan('2026-08-03','2026-11-01');
await executeBoundedRosterImport(old.plan,body=>call(old.file,body));
const future=await plan('2026-11-02','2027-01-31',2);
await executeBoundedRosterImport(future.plan,body=>call(future.file,body));
const replacement=await plan('2026-08-03','2026-11-01');
await call(replacement.file,{phase:'bounded-begin',revision:replacement.plan.revision,manifest:replacement.plan.manifest});
await call(replacement.file,{phase:'bounded-events',revision:replacement.plan.revision,batch:replacement.plan.batches[0]});
assert.equal(sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(old.file.id).active,1,'interruption retains active roster');
db.failRunIncludes='SET active=1 WHERE id=?';
await assert.rejects(executeBoundedRosterImport(replacement.plan,body=>call(replacement.file,body)),/Injected/);
assert.equal(sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(old.file.id).active,1);
assert.equal(sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(replacement.file.id).active,0);
const result=await executeBoundedRosterImport(replacement.plan,body=>call(replacement.file,body));
assert.deepEqual(result.retiredFileIds,[old.file.id]);
assert.equal(sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(future.file.id).active,1);
assert.equal(sqlite.prepare('SELECT COUNT(*) AS n FROM roster_events WHERE file_id=?').get(old.file.id).n,1600,'retired history retained');
assert.deepEqual((await executeBoundedRosterImport(replacement.plan,body=>call(replacement.file,body))).retiredFileIds,[old.file.id]);
const partial=await plan('2026-10-01','2026-11-01',2);
await assert.rejects(executeBoundedRosterImport(partial.plan,body=>call(partial.file,body)),/full retained term range/);
const raced=await plan('2026-08-03','2026-11-01',2);
await executeBoundedRosterImport(raced.plan,async body=>{
  if(body.phase==='bounded-activate') db.beforeBatch=statements=>{
    // Reservation is the first batch; inject at the actual promotion transaction.
    db.beforeBatch=()=>sqlite.prepare("UPDATE roster_files SET parsed_at='changed' WHERE id=?").run(replacement.file.id);
  };
  if(body.phase==='bounded-activate') return assert.rejects(call(raced.file,body),/Concurrent roster change/);
  return call(raced.file,body);
});
assert.equal(sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(replacement.file.id).active,1);
assert.equal(sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(raced.file.id).active,0);
beginMaintenanceAccounting(db,'removal-fixture',true);
const removed=await deactivateManualRosterFiles(context,[replacement.file.id]);
assert.equal(removed.allDeactivated,true);
assert.equal(sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(future.file.id).active,1);
assert.equal(sqlite.prepare('SELECT COUNT(*) AS n FROM roster_events WHERE file_id=?').get(replacement.file.id).n,1600);
assert.equal(sqlite.prepare('SELECT COUNT(*) AS n FROM roster_events WHERE file_id=?').get(old.file.id).n,1600,'replacement and removal retain retired history');
console.log('Rosters: large replacement, interruption, rollback, replay, future preservation, partial rejection, atomic race and recoverable removal passed.');

const r2=new LocalR2();
await storeCachedSnapshot(r2,'facility-overview/v1/mmc/manifest.json',{terms:[{termStart:'2026-08-03',termEnd:'2026-11-01',visibleFrom:'2026-07-20',staffRevision:'empty'}],months:{},coverage:[]});
const published=await loadPublishedDoctorCalendar(r2,{doctorKey:'FIXTURE',sourceTypes:['mmc'],aliases:[{sourceType:'mmc',key:'FIXTURE'}],state:{session:{}}},{range:{startDate:'2026-01-01',endDate:'2026-12-31'},today:'2026-10-02',previousSnapshot:{preview:{events:[{id:'old',source:'MMC',start:'2026-09-01T08:00:00',end:'2026-09-01T16:00:00',title:'Old'},{id:'historical',source:'MMC',start:'2026-04-01T08:00:00',end:'2026-04-01T16:00:00',title:'Past'}]}}});
assert.equal(published.snapshotAvailable,true);
assert.deepEqual(published.snapshot.preview.events.map(event=>event.id),['historical']);
console.log('Published calendar: removed current shifts disappear while cached historical shifts remain.');

// Exercise authenticated APIs through the production request meter.
const password='fixture-password', salt='fixture-salt';
sqlite.prepare('INSERT INTO account_profiles(email,role,password_salt,password_hash) VALUES(?,?,?,?)').run(email,'user',salt,createHash('sha256').update(`${salt}:${password}`).digest('hex'));
async function api(body, flags={}) {
  const request=new Request('https://fixture.test/api/state',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({email,password,...body})});
  const context={request,env:{...contextEnv,...flags},waitUntil(){}};
  context.next=()=>stateHandler(context);
  const response=await middleware(context);
  return {status:response.status,data:await response.json()};
}
const contextEnv={ROSTER_DB:db,ROSTER_FILES:r2,SESSION_SETTINGS_SAVE_ENABLED:'true',BOUNDED_MANUAL_ROSTER_ENABLED:'true',CREATOR_DIRECTORY_ENABLED:'true'};
assert.equal((await api({action:'saveSessionSettings',changes:[]},{SESSION_SETTINGS_SAVE_ENABLED:'false'})).status,503);
assert.equal((await api({action:'saveSessionSettings',targetEmail:'other@example.test',changes:[]})).status,403);
assert.equal((await api({action:'saveSessionSettings',profile:{id:'another',doctorKey:'FIXTURE',sourceTypes:['mmc']},changes:[]})).status,403);
assert.equal((await api({action:'saveDerivedCalendarFile',file:replacement.file,phase:'bounded-begin'})).status,403);
const apiBefore=JSON.parse(sqlite.prepare('SELECT session_json FROM account_states WHERE email=?').get(email).session_json);
const sqlStart=db.sql.length;
const resultSettings=await api({action:'saveSessionSettings',changes:makeSessionPatch(apiBefore,{...apiBefore,settings:{...apiBefore.settings,includeLocations:true}})});
assert.equal(resultSettings.status,200,JSON.stringify(resultSettings));
assert.equal(db.sql.slice(sqlStart).some(sql=>/roster_events|roster_daily_presence/.test(sql)),false);
sqlite.prepare("UPDATE account_profiles SET role='creator' WHERE email=?").run(email);
const retry=await api({action:'saveDerivedCalendarFile',file:replacement.file,phase:'bounded-begin',revision:replacement.plan.revision,manifest:replacement.plan.manifest});
assert.equal(retry.status,200,JSON.stringify(retry));
assert.equal(retry.data.completed,true);
console.log('Authenticated APIs: feature gates, target/profile ownership, Creator-only imports and bounded settings saves passed.');
