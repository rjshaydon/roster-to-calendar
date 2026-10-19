import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {readFile,readdir} from 'node:fs/promises';
import {runInNewContext} from 'node:vm';
import {buildAutomatedDerivedRosterPayload} from '../functions/_lib/automation-import.js';
import {planRosterImportBatches,rosterImportDigest} from '../functions/_lib/roster-import-batches.js';
import {handleManualRosterImport,manualCachedPresenceEnabled} from '../functions/_lib/manual-roster-management.js';
import {beginMaintenanceAccounting} from '../functions/_lib/roster-maintenance-budget.js';
import {runFacilityPublicationStep,loadPublishedFacilityDays,loadPublishedFacilityRange,loadPublishedFacilityStaff} from '../functions/_lib/facility-overview-cache.js';
import {facilityTermDates} from '../functions/_lib/facility-refresh-queue.js';
import {loadCachedRosterInsights} from '../functions/_lib/cached-roster-insights.js';
import {loadPublishedDoctorCalendar} from '../functions/_lib/published-doctor-calendar.js';
import {onRequestGet as feedHandler} from '../functions/api/feed.js';
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


const app=await readFile(new URL('../public/static/app.js',import.meta.url),'utf8');
const clientFunction=app.slice(app.indexOf('async function saveDerivedCalendarFilePayload('),app.indexOf('async function saveSelectedRosterFilesToD1(',app.indexOf('async function saveDerivedCalendarFilePayload(')));
async function store(cached=true){
 const sqlite=new DatabaseSync(':memory:');
 for(const name of(await readdir(new URL('../migrations',import.meta.url))).filter(n=>n.endsWith('.sql')).sort())sqlite.exec(await readFile(new URL('../migrations/'+name,import.meta.url),'utf8'));
 const db=new LocalD1(sqlite),r2=new LocalR2();
 sqlite.prepare('INSERT INTO roster_account_budget(utc_day,maximum_reads,maximum_writes,valid_until)VALUES(?,?,?,?)').run(new Date().toISOString().slice(0,10),100000000,10000000,new Date(Date.now()+3600000).toISOString());
 const env={ROSTER_DB:db,ROSTER_FILES:r2,FACILITY_AUTOMATIC_PUBLICATION_ENABLED:'true',FACILITY_AUTOMATIC_SOURCES:'mmc,mch,ddh,vhh',FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST:'mmc,mch,ddh,vhh',FACILITY_SHARED_ROLLOUT_ACTIVE:'true',FACILITY_SHARED_EMERGENCY_PAUSED:'false',FACILITY_SHARED_READER_COHORT:'all',FACILITY_SHARED_READER_SOURCE_ALLOWLIST:'mmc,mch,ddh,vhh',FACILITY_SHARED_METADATA_ENABLED:'true',FACILITY_SHARED_DAYS_ENABLED:'true',FACILITY_LEGACY_READS_PAUSED:'true',ROSTER_INSIGHT_READS_ENABLED:'true',SUBSCRIPTION_ARTIFACT_MODE:'ready'};
 if(!cached)env.FACILITY_LEGACY_READS_PAUSED='false';
 let request=0;const calls=[];
 const send=async(file,body)=>{calls.push(body.phase);beginMaintenanceAccounting(db,'manual-cache-'+(++request),true);return handleManualRosterImport({env},{file,...body},'creator@example.test');};
 const vm={planRosterImportBatches,rosterImportDigest,Date,Error,refreshManualRosterBudget:async()=>{},manualRosterBudgetCheckedAt:Date.now(),calendarStoreRequest:async(action,body)=>{assert.equal(action,'saveDerivedCalendarFile');return send(body.file,body);},setRosterSyncState(){},selectedFiles:[],finishManualRosterPublication:async()=>({publicationPending:true})};
 const client=runInNewContext('('+clientFunction.trim()+')',vm);
 return {sqlite,db,r2,env,send,calls,client};
}
const totals=[];
for(const [sourceId,path] of [['monash-adults','fixtures/AdultTerm1.2026.xlsx'],['monash-paeds','fixtures/Paeds_Term_2_2026.xlsx']]){
 const payload=await buildAutomatedDerivedRosterPayload({sourceId,file:new File([await readFile(new URL('../'+path,import.meta.url))],path.split('/').at(-1)),contentHash:'fixture',fileId:'manual:'+crypto.randomUUID()});
 for(const cached of [false,true]){
  const ctx=await store(cached);const file={...payload.file};const input={...payload,file};
  const result=await ctx.client(input,file,[]);assert.equal(result.completed,true);
  const events=ctx.sqlite.prepare('SELECT COUNT(*) n FROM roster_events WHERE file_id=?').get(file.id).n;
  const presence=ctx.sqlite.prepare('SELECT COUNT(*) n FROM roster_daily_presence').get().n;
  assert.equal(events,Object.values(payload.eventsByDoctor).flat().length);
  assert.equal(ctx.sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(file.id).active,1);
  const job=ctx.sqlite.prepare('SELECT manifest_json FROM roster_import_jobs').get();const manifest=JSON.parse(job.manifest_json);
  assert.equal(manifest.presenceMode,cached?'cached':undefined);
  assert.equal(presence,cached?0:manifest.batches.reduce((n,b)=>n+b.presenceRows,0));
  assert.equal(ctx.calls.filter(p=>p==='bounded-presence').length,cached?1:manifest.batches.length);
  const reserved=ctx.sqlite.prepare('SELECT SUM(reserved_writes) n FROM roster_maintenance_receipts').get().n;
  const writes=ctx.db.rowsWritten;await ctx.client(input,file,[]);assert.equal(ctx.db.rowsWritten,writes,'retry of a complete plan changes no rows');
  totals.push({site:payload.file.sourceType,mode:cached?'cached':'legacy',events,presence,logicalRowChanges:writes,reservedWrites:reserved});
 }
}
for(let i=0;i<totals.length;i+=2){const legacy=totals[i],cached=totals[i+1];assert.ok(cached.reservedWrites<legacy.reservedWrites*0.65);assert.ok(cached.logicalRowChanges<legacy.logicalRowChanges);}
console.log(JSON.stringify({note:'Local row-change and conservative reservation comparison; neither is native Cloudflare billed usage. Publication costs are separate.',imports:totals},null,2));
// Exercise cached publication and colleague readers for every supported site.
for(const site of ['mmc','mch','ddh','vhh']){
 const ctx=await store();const source=site.toUpperCase(),key='SYNTHETIC',other='COLLEAGUE';
 const file={id:'manual:'+crypto.randomUUID(),sourceType:site,sourceId:'manual-'+site,name:'Synthetic.xlsx'};
 const night={id:'night',source,title:source+': Night',rawValue:'Night',seniority:'Intern',start:'2026-11-02T23:00:00+11:00',end:'2026-11-03T09:00:00+11:00'};
 const leave={id:'leave',source,title:'Annual Leave',rawValue:'AL',seniority:'Intern',allDay:true,start:'2026-11-04',end:'2026-11-10'};
 const together={...night,id:'colleague',seniority:'Senior Registrar'};
 const payload={file,doctors:[{key,displayName:'Synthetic Doctor',seniority:'Intern'},{key:other,displayName:'Synthetic Colleague',seniority:'Senior Registrar'}],eventsByDoctor:{[key]:[night,leave],[other]:[together]},issuesByDoctor:{}};
 assert.equal((await ctx.client(payload,file,[])).completed,true);
 assert.equal(ctx.sqlite.prepare('SELECT COUNT(*) n FROM roster_daily_presence').get().n,0);
 const dates=facilityTermDates('2026-11-02','2026-11-10');const context={env:ctx.env};
 beginMaintenanceAccounting(ctx.db,'publish-'+site,true);
 const planned=await runFacilityPublicationStep(context,site,{mode:'plan',termStart:'2026-11-02',dates});assert.equal(planned.ok,true,JSON.stringify(planned));
 for(let batchIndex=0;batchIndex<planned.batchCount;batchIndex++)assert.equal((await runFacilityPublicationStep(context,site,{mode:'build-batch',termStart:'2026-11-02',dates,operationRevision:planned.operationRevision,batchIndex})).ok,true);
 for(const month of planned.months)assert.equal((await runFacilityPublicationStep(context,site,{mode:'build-month',termStart:'2026-11-02',dates,operationRevision:planned.operationRevision,month})).ok,true);
 assert.equal((await runFacilityPublicationStep(context,site,{mode:'finalize',termStart:'2026-11-02',dates,operationRevision:planned.operationRevision})).ok,true);
 const today='2026-10-20';
 const day=await loadPublishedFacilityDays(ctx.r2,[site],'2026-11-02',today);assert.ok(day.rows.some(r=>r.event.title.includes('Night')&&r.event.end.startsWith('2026-11-03T09:00')),'On shift preserves the overnight end on its roster start date: '+JSON.stringify(day));
 const range=await loadPublishedFacilityRange(ctx.r2,[site],'2026-11-02','2026-11-10',today);assert.ok(range.events.some(r=>r.event.allDay),'multi-day leave remains in calendar source');
 const staff=await loadPublishedFacilityStaff(ctx.r2,[site],'2026-11-02',today);assert.ok(staff.members.some(m=>m.doctorKey===key&&m.seniority==='Intern'),'historical grade remains unchanged');
 const personal=await loadPublishedDoctorCalendar(ctx.r2,{doctorKey:key,displayName:'Synthetic Doctor',sourceTypes:[site],aliases:[{sourceType:site,key}],state:{session:{}}},{range:{startDate:'2026-11-02',endDate:'2026-11-10'},today});
 assert.equal(personal.snapshotAvailable,true);assert.equal(personal.snapshot.preview.events.length,2);assert.ok(personal.snapshot.preview.events.some(e=>e.seniority==='Intern'));
 ctx.sqlite.prepare("INSERT INTO account_profiles(email)VALUES('synthetic@example.test')").run();
 ctx.sqlite.prepare('INSERT INTO account_claims(email,source_type,doctor_key,display_name)VALUES(?,?,?,?)').run('synthetic@example.test',site,key,'Synthetic Doctor');
 ctx.sqlite.prepare("INSERT INTO subscription_tokens(token,email)VALUES('synthetic-token','synthetic@example.test')").run();
 const feedWrites=ctx.db.rowsWritten;
 const feed=await feedHandler({env:{ROSTER_DB:ctx.db},request:new Request('https://fixture/api/feed?token=synthetic-token')});
 assert.equal(feed.status,200);const ics=await feed.text();assert.ok(ics.includes('Night'));assert.ok(ics.includes('Annual Leave'));assert.equal(ctx.db.rowsWritten,feedWrites,'subscription readers write nothing');
 const beforeQueries=ctx.db.sql.length;
 const insights=await loadCachedRosterInsights(ctx.r2,{sourceTypes:[site],startDate:'2026-11-02',endDate:'2026-11-03',overlapDoctorKeys:[key],excludeDoctorKeys:[key]},today,event=>event.kind==='shift'||event.title.includes('Night'));
 assert.equal(insights.ok,true);assert.ok(insights.doctors.some(d=>d.doctorKey===other),'colleague matching works without the SQL index');assert.equal(ctx.db.sql.length,beforeQueries,'cached colleague reads use no D1');
 // A changed rollout must refuse cached cursor advancement/activation, not
 // silently build a legacy index or expose incomplete colleague results.
 ctx.env.FACILITY_SHARED_READER_COHORT='sample';
 assert.equal(manualCachedPresenceEnabled(ctx.env,site),false);
 await assert.rejects(ctx.send(file,{phase:'bounded-activate',revision:ctx.sqlite.prepare('SELECT plan_revision FROM roster_import_jobs').get().plan_revision}),/readers are unavailable/);
}
console.log('All-site manual cached imports: actual browser driver, zero presence rows, one-step cursor, completed replay, overnight On shift, term grades, leave, colleague matching and rollout guard passed.');

// Lost response/retry keeps the previous active roster until the complete new
// plan activates. Old legacy jobs keep their original hash after rollout.
function smallPayload(label){const file={id:'manual:'+crypto.randomUUID(),sourceId:'manual-mmc',sourceType:'mmc',name:'Synthetic.xlsx'};return {file,doctors:[{key:'DOCTOR',displayName:'Synthetic Doctor',seniority:'Intern'}],eventsByDoctor:{DOCTOR:[{id:'day',source:'MMC',title:label,seniority:'Intern',start:'2026-11-02T08:00:00',end:'2026-11-02T17:00:00'}]},issuesByDoctor:{}};}
const interrupted=await store();const old=smallPayload('Old'),next=smallPayload('New');
await interrupted.client(old,old.file,[]);
const newPlan=await planRosterImportBatches(next);newPlan.manifest={...newPlan.manifest,presenceMode:'cached'};newPlan.revision=await rosterImportDigest(newPlan.manifest);
await interrupted.send(next.file,{phase:'bounded-initialize',revision:newPlan.revision,manifest:newPlan.manifest});
await interrupted.send(next.file,{phase:'bounded-events',revision:newPlan.revision,batch:newPlan.batches[0]});
assert.equal(interrupted.sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(old.file.id).active,1);
assert.equal(interrupted.sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(next.file.id).active,0);
const stagedRows=interrupted.sqlite.prepare('SELECT COUNT(*) n FROM roster_events WHERE file_id=?').get(next.file.id).n;
await interrupted.client(next,next.file,[]);
assert.equal(interrupted.sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(old.file.id).active,0);
assert.equal(interrupted.sqlite.prepare('SELECT COUNT(*) n FROM roster_events WHERE file_id=?').get(next.file.id).n,stagedRows);
assert.equal(interrupted.sqlite.prepare('SELECT COUNT(*) n FROM roster_daily_presence').get().n,0);
const legacy=await store(false),legacyPayload=smallPayload('Legacy');const legacyPlan=await planRosterImportBatches(legacyPayload);
await legacy.send(legacyPayload.file,{phase:'bounded-begin',manifest:legacyPlan.manifest,revision:legacyPlan.revision});
legacy.env.FACILITY_LEGACY_READS_PAUSED='true';
assert.equal((await legacy.client(legacyPayload,legacyPayload.file,[])).completed,true);
assert.equal(JSON.parse(legacy.sqlite.prepare('SELECT manifest_json FROM roster_import_jobs').get().manifest_json).presenceMode,undefined);
assert.ok(legacy.sqlite.prepare('SELECT COUNT(*) n FROM roster_daily_presence').get().n>0);
const disabled=await store(false),forged=smallPayload('Forged'),forgedPlan=await planRosterImportBatches(forged);
forgedPlan.manifest.presenceMode='cached';forgedPlan.revision=await rosterImportDigest(forgedPlan.manifest);
const beforeDisabled=disabled.db.rowsWritten;
await assert.rejects(disabled.send(forged.file,{phase:'bounded-initialize',manifest:forgedPlan.manifest,revision:forgedPlan.revision}),/readers are unavailable/);
assert.equal(disabled.db.rowsWritten,beforeDisabled,'a caller cannot opt into an unapproved cached path');
console.log('Interrupted cached import resumes without duplicate facts; pre-rollout legacy plan remains compatible; forged cached-mode opt-in is rejected before writes.');
