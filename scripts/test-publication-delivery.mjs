import {onRequestPost as artifact} from '../functions/api/automation/doctor-calendars.js';
import {publishDoctorMonth} from './process-doctor-calendars.mjs';
import {publishDoctorSource} from './process-doctor-source.mjs';
import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {readFile,readdir} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {replaceDerivedRosterFile} from '../functions/_lib/d1-calendar.js';
import {facilityRefreshStatements,facilityTermDates} from '../functions/_lib/facility-refresh-queue.js';
import {onRequestPost as refresh} from '../functions/api/automation/facility-refresh.js';
import {runFacilityPublicationStep,facilityMetadataManifestKey,facilityDayObjectKey,loadPublishedFacilityDays} from '../functions/_lib/facility-overview-cache.js';
import {beginMaintenanceAccounting,reserveRosterMaintenanceBudget,configurePublicationReservation,rosterProcessingHeadroom,finishMaintenanceAccounting} from '../functions/_lib/roster-maintenance-budget.js';
import {publicationHoldStatement,publicationTail,publicationHeldWrites,releasePublicationHoldStatement} from '../functions/_lib/publication-delivery-budget.js';
import {doctorMonthIndexKey} from '../functions/_lib/doctor-calendar-artifacts.js';
import {doctorSourceKey,doctorSourceCurrentKey} from '../functions/_lib/doctor-source-artifacts.js';
import {drainRosterPublication} from './roster-publication-driver.mjs';
const sqlite=new DatabaseSync(':memory:');
for(const file of(await readdir('migrations')).filter(f=>f.endsWith('.sql')).sort())sqlite.exec(await readFile('migrations/'+file,'utf8'));
let fail='',sequence=0;const sqls=[];
const db={prepare(sql){sqls.push(sql);return {args:[],bind(...args){this.args=args;return this;},async first(){return sqlite.prepare(sql).get(...this.args)||null;},async all(){const results=sqlite.prepare(sql).all(...this.args);return {results,meta:{rows_read:results.length+1,rows_written:0}};},async run(){if(fail&&sql.includes(fail)){fail='';throw Error('Injected failure');}const r=sqlite.prepare(sql).run(...this.args);return {meta:{changes:Number(r.changes),rows_read:Number(r.changes)+1,rows_written:Number(r.changes)}};}};},async batch(statements){sqlite.exec('SAVEPOINT delivery');try{const rows=[];for(const s of statements)rows.push(await s.run());sqlite.exec('RELEASE delivery');return rows;}catch(e){sqlite.exec('ROLLBACK TO delivery; RELEASE delivery');throw e;}}};
const objects=new Map();let generation=0,puts=0;
const r2={async put(key,input,options={}){const prior=objects.get(key);if(options.onlyIf?.etagMatches&&prior?.etag!==options.onlyIf.etagMatches)return null;if(options.onlyIf?.etagDoesNotMatch==='*'&&prior)return null;const bytes=typeof input==='string'?Buffer.from(input):input instanceof Uint8Array?input:new Uint8Array(await new Response(input).arrayBuffer());const etag='test-'+(++generation);objects.set(key,{bytes,etag});puts++;return {etag};},async get(key){const item=objects.get(key);if(!item)return null;return {etag:item.etag,arrayBuffer:async()=>item.bytes.buffer.slice(item.bytes.byteOffset,item.bytes.byteOffset+item.bytes.byteLength),body:new Response(item.bytes).body,json:async()=>JSON.parse(Buffer.from(item.bytes).toString())};},async head(key){const o=objects.get(key);return o?{etag:o.etag}:null;}};
const context={env:{ROSTER_DB:db,ROSTER_FILES:r2}},day=new Date().toISOString().slice(0,10);
sqlite.prepare('INSERT INTO roster_account_budget(utc_day,maximum_writes,maximum_reads,valid_until)VALUES(?,?,?,?)').run(day,100000,4000000,new Date(Date.now()+600000).toISOString());
function account(){beginMaintenanceAccounting(db,'test-'+(++sequence),true);}
const doctors=[{key:'SYNTHETIC DOCTOR',displayName:'Synthetic Doctor',seniority:'Junior Registrar'}];
async function fixture(id,date,end=date){
 const file={id,sourceId:'monash-adults',sourceType:'mmc',name:id+'.xlsx',active:true};
 const events=[{id:'night',source:'MMC',title:'Night',seniority:'Junior Registrar',start:date+'T23:00:00',end:end+'T08:00:00'}];
 await replaceDerivedRosterFile(db,file,doctors,{[doctors[0].key]:events},{});return file;
}
await fixture('term3','2026-11-01','2026-11-02');
await fixture('term4','2026-11-02','2026-11-03');
const env={...context.env,ROSTER_AUTOMATION_TOKEN:'synthetic',ROSTER_AUTOMATION_WRITES_ENABLED:'true',ROSTER_AUTOMATION_QUEUE_ENABLED:'true',ROSTER_AUTOMATION_SOURCE_ALLOWLIST:'monash-adults',FACILITY_AUTOMATIC_PUBLICATION_ENABLED:'true',FACILITY_SHARED_EMERGENCY_PAUSED:'false',FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST:'mmc'};
async function http(){account();const result=await refresh({env,request:new Request('https://test/api/automation/facility-refresh',{method:'POST',headers:{authorization:'Bearer synthetic'},body:JSON.stringify({sourceId:'monash-adults'})})});const body=await result.json();assert.equal(result.status,200,JSON.stringify(body));return body;}
// Begin Term 4 and pause its first batch. An arriving older-term job must not
// displace its publication owner on the next HTTP invocation.
await db.batch(facilityRefreshStatements(db,'mmc',['2026-11-02','2026-11-03'],'term4'));
const first=await http();assert.ok(first.operationRevision,JSON.stringify(first));await http();
await db.batch(facilityRefreshStatements(db,'mmc',['2026-11-01'],'term3'));
const beforeFailure=sqlite.prepare("SELECT next_month FROM facility_refresh_jobs WHERE term_start='2026-11-02'").get().next_month;
fail='UPDATE facility_refresh_jobs SET plan_json';
await assert.rejects(http(),/Injected failure/);
assert.equal(sqlite.prepare("SELECT next_month FROM facility_refresh_jobs WHERE term_start='2026-11-02'").get().next_month,beforeFailure);
assert.ok(sqlite.prepare("SELECT remaining_writes FROM roster_publication_holds WHERE term_start='2026-11-02'").get().remaining_writes>0,'failed checkpoint retains its delivery commitment');
const resumed=await http();assert.equal(resumed.mode,'build-month');assert.equal(resumed.month,'2026-11');
assert.equal(sqlite.prepare("SELECT next_batch FROM facility_refresh_jobs WHERE term_start='2026-08-03'").get().next_batch,0,'Term 3 cannot preempt an unfinished owner');
// Small publication steps can run when the import-launch minimum is not met.
const gate={maximum_writes:500,allocated_writes:0,maximum_reads:200000,allocated_reads:0,valid_until:new Date(Date.now()+60000).toISOString()};
assert.equal(rosterProcessingHeadroom(gate),false);assert.equal(rosterProcessingHeadroom(gate,new Date(),true),true);
// Protected commitments apply across sources, survive a budget refresh and
// UTC-day rollover, and convert with the actual metered reservation.
const hold=sqlite.prepare('SELECT * FROM roster_publication_holds WHERE term_start=?').get('2026-11-02');assert.ok(hold.remaining_writes>0);
sqlite.prepare('UPDATE roster_account_budget SET maximum_writes=allocated_writes+1000 WHERE utc_day=?').run(day);
account();assert.equal(await reserveRosterMaintenanceBudget(db,1000,1),false,'bulk cannot consume another operation delivery hold');
const job=sqlite.prepare("SELECT * FROM facility_refresh_jobs WHERE term_start='2026-11-02'").get();
account();configurePublicationReservation(db,job);assert.equal(await reserveRosterMaintenanceBudget(db,64,125000),true,'publication can advance even if total commitments exceed this grant');
const reduced=sqlite.prepare('SELECT remaining_writes FROM roster_publication_holds WHERE term_start=?').get(job.term_start).remaining_writes;assert.equal(reduced,hold.remaining_writes,'capacity remains protected until a durable cursor commits, including unknown responses');
const before=sqlite.prepare('SELECT allocated_writes FROM roster_account_budget').get().allocated_writes;
await finishMaintenanceAccounting(db,{rowsRead:20,rowsWritten:10,metadataComplete:true});await finishMaintenanceAccounting(db,{rowsRead:20,rowsWritten:10,metadataComplete:true});
assert.equal(sqlite.prepare('SELECT allocated_writes FROM roster_account_budget').get().allocated_writes,before-(112-34),'a settled receipt refunds exactly once');
sqlite.prepare('UPDATE roster_account_budget SET maximum_writes=100000,maximum_reads=4000000 WHERE utc_day=?').run(day);
assert.ok(sqlite.prepare(`SELECT ${publicationHeldWrites} AS n`).get().n>0,'grant replacement does not erase commitments');
sqlite.prepare("INSERT INTO roster_account_budget(utc_day,maximum_writes,maximum_reads)VALUES('2099-01-01',100000,4000000)").run();
assert.ok(sqlite.prepare(`SELECT ${publicationHeldWrites} AS n`).get().n>0,'commitments are operation-scoped, not forgotten at midnight');
// Unknown settlement keeps the full receipt; a new UTC grant still honors
// the same operation's held capacity rather than treating it as free credit.
const RealDate=globalThis.Date;const tomorrow=Date.parse(day+'T00:00:00Z')+86400000+3600000;
sqlite.prepare('INSERT OR REPLACE INTO roster_account_budget(utc_day,maximum_writes,maximum_reads,valid_until)VALUES(?,?,?,?)').run(new Date(tomorrow).toISOString().slice(0,10),100000,4000000,new Date(tomorrow+600000).toISOString());
try{
 globalThis.Date=class extends RealDate{constructor(...args){super(...(args.length?args:[tomorrow]));}static now(){return tomorrow;}};
 account();configurePublicationReservation(db,job);assert.equal(await reserveRosterMaintenanceBudget(db,64,125000),true);
 const spent=sqlite.prepare('SELECT allocated_writes FROM roster_account_budget WHERE utc_day=?').get(new Date().toISOString().slice(0,10)).allocated_writes;
 await finishMaintenanceAccounting(db,{rowsRead:1,rowsWritten:1,metadataComplete:false});
 assert.equal(sqlite.prepare('SELECT allocated_writes FROM roster_account_budget WHERE utc_day=?').get(new Date().toISOString().slice(0,10)).allocated_writes,spent,'unknown metadata never refunds a reservation');
}finally{globalThis.Date=RealDate;}
// Force another publication to move the shared manifest. The retained Term 4
// daily batch must rebase onto it without querying the roster days again.
const oldPlan=first,term3Plan=await runFacilityPublicationStep(context,'mmc',{mode:'plan',termStart:'2026-08-03',dates:['2026-11-01']});
// Finish Term 4 without subscription gating for this unrelated-base fixture.
await http(); // finalize
const rootBefore=await r2.get(facilityMetadataManifestKey('mmc'));assert.ok(rootBefore);
// Start a new Term 4 operation, retain its daily data, then publish an unrelated
// term through its own completed operation. Simulate external metadata changes
// by editing only the shared manifest, as the independent admin tools can do.
const ownPlan=await runFacilityPublicationStep(context,'mmc',{mode:'plan',termStart:'2026-11-02',dates:['2026-11-02','2026-11-03']});
await runFacilityPublicationStep(context,'mmc',{mode:'build-batch',termStart:'2026-11-02',dates:['2026-11-02','2026-11-03'],operationRevision:ownPlan.operationRevision,batchIndex:0});
const {gunzipSync,gzipSync}=await import('node:zlib');
const root=JSON.parse(gunzipSync(Buffer.from(await (await r2.get(facilityMetadataManifestKey('mmc'))).arrayBuffer())).toString());
root.revision='b'.repeat(64);root.unrelatedTermEvidence='changed';
const sharedDayKey=facilityDayObjectKey('mmc','2026-11-01','a'.repeat(64));
await r2.put(sharedDayKey,JSON.stringify({rows:[{sourceType:'mmc',doctorKey:doctors[0].key,displayName:doctors[0].displayName,event:{id:'other-term-newest',source:'MMC',title:'Night',start:'2026-11-01T23:00:00',end:'2026-11-02T08:00:00',seniority:'Junior Registrar'}}]}));
root.days['2026-11-01']={key:sharedDayKey,revision:'a'.repeat(64)};
await r2.put(facilityMetadataManifestKey('mmc'),gzipSync(JSON.stringify(root)));
sqls.length=0;
const rebased=await runFacilityPublicationStep(context,'mmc',{mode:'build-month',termStart:'2026-11-02',dates:['2026-11-02','2026-11-03'],operationRevision:ownPlan.operationRevision,month:'2026-11'});
assert.equal(rebased.reconciled,true,JSON.stringify(rebased));assert.equal(rebased.nextBatch,1);
assert.ok(!sqls.some(sql=>sql.includes('roster_events')),'unrelated base reconciliation reuses day results, not clinical D1 rows');
async function artifactRequest(input,{binary=false,metadata=false,body}={}){
 const headers={authorization:'Bearer synthetic','Content-Type':'application/json'};let payload=JSON.stringify(input);
 if(body!==undefined){payload=Buffer.from(body);headers['Content-Type']='application/octet-stream';headers['Content-Length']=String(payload.length);headers['x-doctor-artifact']=Buffer.from(JSON.stringify({...input,bytes:payload.length})).toString('base64');}
 const response=await artifact({env,request:new Request('https://test/api/automation/doctor-calendars',{method:'POST',headers,body:payload})});
 assert.equal(response.status,200,await response.clone().text());
 if(binary){const bytes=new Uint8Array(await response.arrayBuffer());return metadata?{bytes,etag:response.headers.get('x-object-etag'),key:input.key}:bytes;}
 return response.json();
}
// Build merged months and require personal indexes before any completion.
for(const month of rebased.months)await runFacilityPublicationStep(context,'mmc',{mode:'build-month',termStart:'2026-11-02',dates:['2026-11-02','2026-11-03'],operationRevision:rebased.operationRevision,month});
const gateResult=await runFacilityPublicationStep(context,'mmc',{mode:'finalize',termStart:'2026-11-02',dates:['2026-11-02','2026-11-03'],operationRevision:rebased.operationRevision,requireDoctorCalendars:true});
assert.ok(gateResult.doctorCalendars?.length);
for(const pointer of gateResult.doctorCalendars)await publishDoctorMonth({sourceType:'mmc',...pointer,request:artifactRequest});
const sourceGate=await runFacilityPublicationStep(context,'mmc',{mode:'finalize',termStart:'2026-11-02',dates:['2026-11-02','2026-11-03'],operationRevision:rebased.operationRevision,requireDoctorCalendars:true});assert.ok(sourceGate.doctorSource);
// Reproduce loss after the shared root write but before doctor promotion.
fail="SET status = 'complete'";
await assert.rejects(runFacilityPublicationStep(context,'mmc',{mode:'finalize',termStart:'2026-11-02',dates:['2026-11-02','2026-11-03'],operationRevision:rebased.operationRevision}),/Injected/);
const retry=await runFacilityPublicationStep(context,'mmc',{mode:'finalize',termStart:'2026-11-02',operationRevision:rebased.operationRevision,requireDoctorCalendars:true});assert.equal(retry.deferred,true,'already-published shortcut cannot skip the personal gate');
const {manifestKey}=retry.doctorSource;
const candidate=JSON.parse(gunzipSync(Buffer.from(await (await r2.get(manifestKey)).arrayBuffer())).toString());
await publishDoctorSource({sourceType:'mmc',request:artifactRequest,...retry.doctorSource});
const repaired=await runFacilityPublicationStep(context,'mmc',{mode:'finalize',termStart:'2026-11-02',operationRevision:rebased.operationRevision,requireDoctorCalendars:true});assert.equal(repaired.repaired,true);
assert.equal((await r2.get(doctorSourceCurrentKey('mmc'))).json instanceof Function,true);
assert.equal((await (await r2.get(doctorSourceCurrentKey('mmc'))).json()).parentEtag,(await r2.head(facilityMetadataManifestKey('mmc'))).etag);
const status=await artifactRequest({mode:'delivery-status',sourceType:'mmc'});assert.equal(status.ready,true);assert.ok(status.terms.some(t=>t.termStart==='2026-11-02'));
const actualSource=(await artifactRequest({mode:'source-current',sourceType:'mmc'})).current;
const catalogue=JSON.parse(Buffer.from(await artifactRequest({mode:'object',sourceType:'mmc',key:actualSource.catalogueKey},{binary:true})).toString());
const person=catalogue.doctors[doctors[0].key];
const personal=await (await r2.get(person.key)).json();
assert.ok(personal.events.some(e=>e.id==='other-term-newest'),'rebased November preserves the newest previous-term night');
assert.ok(personal.events.some(e=>String(e.start).startsWith('2026-11-02')),'actual personal artifact includes Term 4');
assert.ok(personal.events.every(e=>!e.seniority||e.seniority==='Junior Registrar'),'historical grades survive preparation');
// A carry-in boundary does not become a permanently pending full next term.
sqlite.exec("UPDATE facility_refresh_jobs SET status='complete';DELETE FROM roster_publication_holds");
await fixture('end-night','2027-01-31','2027-02-01');
sqlite.prepare("INSERT OR IGNORE INTO facility_term_visibility(source_type,term_start,visible_from)VALUES('mmc','2027-02-01','2027-01-01')").run();
await db.batch(facilityRefreshStatements(db,'mmc',['2027-02-01'],'boundary'));
const carry=await http();assert.equal(carry.boundaryOnly,true,JSON.stringify(carry));
assert.equal(sqlite.prepare("SELECT status FROM facility_refresh_jobs WHERE term_start='2027-02-01'").get().status,'complete');
assert.equal(sqlite.prepare("SELECT COUNT(*) n FROM roster_events WHERE file_id='end-night' AND end_date='2027-02-01'").get().n,1,'retiring the boundary job preserves the overnight event');

// A different site's pending publication protects the same account pool.
sqlite.prepare("INSERT INTO facility_refresh_jobs(source_type,term_start,dates_json,content_signature,request_revision,updated_at)VALUES('mch','2026-11-02','[\"2026-11-02\"]','mch','mch-owner','now')").run();
await publicationHoldStatement(db,'mch','2026-11-02','mch-owner',['2026-11-02']).run();
const mchHold=sqlite.prepare("SELECT remaining_writes FROM roster_publication_holds WHERE source_type='mch'").get().remaining_writes;
sqlite.prepare('UPDATE roster_account_budget SET maximum_writes=allocated_writes+? WHERE utc_day=?').run(mchHold,day);
account();assert.equal(await reserveRosterMaintenanceBudget(db,1,1),false,'MMC bulk work cannot spend MCH delivery capacity');
sqlite.prepare("UPDATE facility_refresh_jobs SET request_revision='mch-successor' WHERE source_type='mch'").run();
await publicationHoldStatement(db,'mch','2026-11-02','mch-successor',['2026-11-02']).run();
await releasePublicationHoldStatement(db,'mch','2026-11-02','mch-owner').run();
assert.equal(sqlite.prepare("SELECT owner_id FROM roster_publication_holds WHERE source_type='mch'").get().owner_id,'mch-successor','a stale owner cannot release successor capacity');
sqlite.prepare("UPDATE facility_refresh_jobs SET status='complete' WHERE source_type='mch'").run();await releasePublicationHoldStatement(db,'mch','2026-11-02','mch-successor').run();
assert.equal(sqlite.prepare("SELECT COUNT(*) n FROM roster_publication_holds WHERE source_type='mch'").get().n,0);
// The driver handles one stale-plan reconciliation, caps retries and reports
// a run-limit yield instead of mislabelling it a quota failure.
let steps=0;const budget={remaining:2};
const yielded=await drainRosterPublication({budget,step:async()=>{steps++;return {completed:true};}});assert.equal(yielded.reason,'run-limit');assert.equal(steps,2);
let once=0;assert.equal((await drainRosterPublication({step:async()=>++once===1?{recoverableConflict:true}:{idle:true}})).completed,true);
assert.equal((await drainRosterPublication({step:async()=>({recoverableConflict:true})})).reason,'publication-conflict');
console.log('Delivery: HTTP owner priority, shared capacity, exactly-once settlement, midnight protection, cached rebase, personal completion gate and bounded continuation passed. Synthetic data only; local metadata is not Cloudflare billing evidence.');
