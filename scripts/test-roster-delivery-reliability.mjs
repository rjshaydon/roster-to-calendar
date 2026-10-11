import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {readFile,readdir,mkdtemp,writeFile,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {spawnSync} from 'node:child_process';
import {markRosterDeliveryDeferred,loadRosterDeliveryWarnings,clearRosterDeliveryWarning,loadActiveRosterImport} from '../functions/_lib/roster-delivery-health.js';
import {onRequestPost as metadataCheck} from '../functions/api/automation/roster-check.js';
import {onRequestGet as pendingCheck} from '../functions/api/automation/pending.js';
import {requestQueuedRosterProcessing,recordRosterDispatchLifecycle} from '../functions/_lib/automation-dispatch.js';
const sqlite=new DatabaseSync(':memory:');
for(const name of (await readdir('migrations')).filter(n=>n.endsWith('.sql')).sort()) sqlite.exec(await readFile('migrations/'+name,'utf8'));
for(const definition of ['name TEXT','source_type TEXT','size INTEGER','last_modified INTEGER'])sqlite.exec('ALTER TABLE raw_roster_files ADD COLUMN '+definition);
const db={prepare(sql){return {args:[],bind(...args){this.args=args;return this;},async first(){return sqlite.prepare(sql).get(...this.args)||null;},async all(){return {results:sqlite.prepare(sql).all(...this.args)};},async run(){const r=sqlite.prepare(sql).run(...this.args);return {meta:{changes:Number(r.changes)}};}};}};
const objects=new Map();let puts=0;
const r2={async get(key){const value=objects.get(key);return value?{json:async()=>JSON.parse(value)}:null;},async put(key,value){puts++;objects.set(key,value);},async delete(key){objects.delete(key);}};
const queuedAt=new Date(Date.now()-20*60000).toISOString();
await markRosterDeliveryDeferred(r2,'monash-adults',queuedAt);
await markRosterDeliveryDeferred(r2,'monash-adults',queuedAt);
assert.equal(puts,1,'unchanged deferred status causes no repeat R2 writes');
assert.deepEqual(await loadRosterDeliveryWarnings(r2,['mmc','mmc','mch']),[{sourceType:'mmc',status:'delayed'}]);
assert.deepEqual(await loadRosterDeliveryWarnings(r2,['mmc'],Date.parse(queuedAt)+14*60000),[],'brief coalescing stays quiet');
objects.set('roster-delivery/v1/mch.json','invalid');
assert.deepEqual(await loadRosterDeliveryWarnings(r2,['mch']),[],'diagnostics cannot break the roster interface');
await clearRosterDeliveryWarning(r2,'monash-adults');
assert.deepEqual(await loadRosterDeliveryWarnings(r2,['mmc']),[]);
const now=new Date();
sqlite.prepare('INSERT INTO roster_account_budget(utc_day,allocated_reads,allocated_writes,maximum_reads,maximum_writes,valid_until) VALUES(?,100,100,100,100,?)').run(now.toISOString().slice(0,10),new Date(Date.now()+600000).toISOString());
sqlite.prepare("INSERT INTO raw_roster_files(file_id,name,source_type,size,last_modified,object_key) VALUES('file','Roster.xlsx','mmc',100,0,'test-object')").run();
sqlite.prepare("INSERT INTO roster_sync_runs(id,source_id,file_id,source_file_id,status,started_at) VALUES('run','monash-adults','file','file','queued',?)").run(queuedAt);
assert.equal((await loadActiveRosterImport(db,'monash-adults')).id,'run');
const env={ROSTER_DB:db,ROSTER_FILES:r2,ROSTER_AUTOMATION_WRITES_ENABLED:'true',ROSTER_AUTOMATION_QUEUE_ENABLED:'true',ROSTER_AUTOMATION_SOURCE_ALLOWLIST:'monash-adults',ROSTER_ACCOUNT_BUDGET_ENABLED:'true',GITHUB_ACTIONS_TOKEN:'test-only'};
const originalFetch=globalThis.fetch;
try{
 globalThis.fetch=async()=>{throw Error('Blocked budget must not launch GitHub or download a roster');};
 const blocked=await requestQueuedRosterProcessing(env,{sourceId:'monash-adults'});
 assert.equal(blocked.deferred,true);
 assert.equal(sqlite.prepare('SELECT COUNT(*) n FROM roster_dispatches').get().n,0,'no no-op workflow dispatch');
 assert.equal(sqlite.prepare("SELECT status FROM roster_sync_runs WHERE id='run'").get().status,'queued');
 assert.deepEqual(await loadRosterDeliveryWarnings(r2,['mmc']),[{sourceType:'mmc',status:'delayed'}]);
}finally{globalThis.fetch=originalFetch;}
// Lifecycle records an admitted-but-deferred workflow honestly and coalesces
// subsequent retry attempts; publication clears the warning only after work.
try{
 globalThis.fetch=async()=>new Response(null,{status:204});
 const started=await requestQueuedRosterProcessing({...env,ROSTER_ACCOUNT_BUDGET_ENABLED:'false'},{sourceId:'monash-adults'});
 assert.equal(started.dispatched,true);
 const outcome=await recordRosterDispatchLifecycle(env,{sourceId:'monash-adults',dispatchId:started.dispatch.id,event:'deferred'});
 assert.equal(outcome.ok,true);
 assert.equal(sqlite.prepare('SELECT status FROM roster_dispatches WHERE id=?').get(started.dispatch.id).status,'deferred');
 globalThis.fetch=async()=>{throw Error('Deferred dispatch must coalesce until retry time');};
 assert.equal((await requestQueuedRosterProcessing({...env,ROSTER_ACCOUNT_BUDGET_ENABLED:'false'},{sourceId:'monash-adults'})).dispatched,false);
 sqlite.prepare("UPDATE roster_sync_runs SET status='success' WHERE id='run'").run();
 sqlite.prepare("INSERT INTO roster_sync_runs(id,source_id,file_id,source_file_id,status,started_at) VALUES('obsolete','monash-adults','file','file','queued','2000-01-01T00:00:00.000Z')").run();
 assert.equal(await loadActiveRosterImport(db,'monash-adults'),null,'an obsolete queued autosave cannot block identities after the latest success');
 const publishingEnv={...env,FACILITY_AUTOMATIC_PUBLICATION_ENABLED:'true',ROSTER_METADATA_CHECK_ENABLED:'true',ROSTER_AUTOMATION_TOKEN:'fixture'};
 sqlite.prepare("INSERT INTO facility_term_visibility(source_type,term_start,visible_from) VALUES('mmc','2026-08-03','2026-07-01')").run();
 sqlite.prepare("INSERT INTO facility_refresh_jobs(source_type,term_start,dates_json,content_signature,request_revision,updated_at) VALUES('mmc','2026-08-03','[]','revision','revision',?)").run(queuedAt);
 await recordRosterDispatchLifecycle(publishingEnv,{sourceId:'monash-adults',dispatchId:started.dispatch.id,event:'completed'});
 assert.deepEqual(await loadRosterDeliveryWarnings(r2,['mmc']),[{sourceType:'mmc',status:'delayed'}],'unfinished publication retains the warning after import success');
 sqlite.exec("UPDATE raw_roster_files SET name='AdultTerm3.2026.xlsx';UPDATE roster_sync_runs SET provider_version='1'");
 const check=await (await metadataCheck({env:publishingEnv,request:new Request('https://test/api/automation/roster-check',{method:'POST',headers:{authorization:'Bearer fixture'},body:JSON.stringify({sourceId:'monash-adults',fileName:'AdultTerm3.2026.xlsx',providerVersion:'1'})})})).json();
 assert.equal(check.download,false,'unchanged roster publication never requests another workbook download');
 assert.equal(check.status,'deferred','unfinished publication is visible through unchanged metadata polling');
 const pending=await (await pendingCheck({env:publishingEnv,request:new Request('https://test/api/automation/pending?sourceId=monash-adults',{headers:{authorization:'Bearer fixture'}})})).json();
 assert.equal(pending.publicationPending,true);
 assert.equal(pending.runs.length,0,'publication resumes after the file import is already successful');
 sqlite.prepare("UPDATE facility_refresh_jobs SET status='complete'").run();
 sqlite.prepare("INSERT INTO facility_refresh_jobs(source_type,term_start,dates_json,content_signature,request_revision,updated_at) VALUES('mmc','2026-11-02','[\"2026-11-02\"]','boundary','boundary',?)").run(queuedAt);
 const boundary=await (await pendingCheck({env:publishingEnv,request:new Request('https://test/api/automation/pending?sourceId=monash-adults',{headers:{authorization:'Bearer fixture'}})})).json();
 assert.equal(boundary.publicationPending,false,'an overnight boundary without a next-term roster must not create no-op jobs or a delayed-sync warning');
 await recordRosterDispatchLifecycle(publishingEnv,{sourceId:'monash-adults',dispatchId:started.dispatch.id,event:'completed'});
 assert.deepEqual(await loadRosterDeliveryWarnings(r2,['mmc']),[]);
}finally{globalThis.fetch=originalFetch;}
// The priority inspection itself must stay bounded even if a legacy source
// has accumulated many obsolete pending rows. Pause optional work at the cap.
for(let i=0;i<66;i++)sqlite.prepare("INSERT INTO roster_sync_runs(id,source_id,file_id,source_file_id,status,started_at) VALUES(?,'monash-adults','file','file','queued','1999-01-01T00:00:00.000Z')").run('old-'+i);
assert.ok(await loadActiveRosterImport(db,'monash-adults'),'oversized pending inspection fails closed');
const folder=await mkdtemp(join(tmpdir(),'roster-outcomes-'));
try{
 // Run the real processor in a child so its intentional successful exit on
 // deferral cannot terminate this test. No network access or workbook fetch.
 const preload=join(folder,'preload.mjs'),output=join(folder,'output');
 await writeFile(preload,`let published=false;globalThis.fetch=async url=>{const path=new URL(url).pathname;console.log('TESTCALL '+path);if(path.endsWith('account-budget'))return Response.json({deferred:process.env.TEST_DEFER==='true'});if(path.endsWith('facility-refresh')){published=true;return Response.json(process.env.TEST_PUBLICATION==='deferred'?{deferred:true}:{completed:true});}return Response.json({runs:[],boundedImportEnabled:true,publicationPending:Boolean(process.env.TEST_PUBLICATION)&&(process.env.TEST_PUBLICATION==='deferred'||!published)});};`);
 for(const [defer,publication,expected] of [['true','','deferred'],['false','','completed'],['false','done','completed'],['false','deferred','deferred']]){
  await writeFile(output,'');
  const result=spawnSync(process.execPath,['--import',preload,'scripts/process-roster-queue.mjs'],{encoding:'utf8',env:{...process.env,TEST_DEFER:defer,TEST_PUBLICATION:publication,ROSTER_AUTOMATION_TOKEN:'test-only',ROSTER_AUTOMATION_SOURCE_ID:'monash-adults',GITHUB_OUTPUT:output}});
  assert.equal(result.status,0,result.stderr);
  assert.equal((await readFile(output,'utf8')).trim(),'result='+expected,'GitHub receives the actual processor result');
  assert.equal(result.stdout.includes('/api/automation/raw'),false);
  assert.equal(result.stdout.includes('/api/automation/parser-config'),false,'publication-only work never reparses an unchanged file');
  if(publication)assert.equal(result.stdout.split('TESTCALL /api/automation/facility-refresh').length-1,1);
 }
}finally{await rm(folder,{recursive:true,force:true});}
console.log('Delayed-only warnings, corrupt diagnostics, blocked dispatch, retained queue and honest GitHub outcomes passed.');
