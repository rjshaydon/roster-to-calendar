import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {readFile,readdir,mkdtemp,writeFile,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {spawnSync} from 'node:child_process';
import {markRosterDeliveryDeferred,loadRosterDeliveryWarnings,clearRosterDeliveryWarning} from '../functions/_lib/roster-delivery-health.js';
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
 await recordRosterDispatchLifecycle(env,{sourceId:'monash-adults',dispatchId:started.dispatch.id,event:'completed'});
 assert.deepEqual(await loadRosterDeliveryWarnings(r2,['mmc']),[]);
}finally{globalThis.fetch=originalFetch;}
const folder=await mkdtemp(join(tmpdir(),'roster-outcomes-'));
try{
 // Run the real processor in a child so its intentional successful exit on
 // deferral cannot terminate this test. No network access or workbook fetch.
 const preload=join(folder,'preload.mjs'),output=join(folder,'output');
 await writeFile(preload,`globalThis.fetch=async url=>{const path=new URL(url).pathname;return Response.json(path.endsWith('account-budget')?{deferred:process.env.TEST_DEFER==='true'}:{runs:[]});};`);
 for(const [defer,expected] of [['true','deferred'],['false','completed']]){
  await writeFile(output,'');
  const result=spawnSync(process.execPath,['--import',preload,'scripts/process-roster-queue.mjs'],{encoding:'utf8',env:{...process.env,TEST_DEFER:defer,ROSTER_AUTOMATION_TOKEN:'test-only',ROSTER_AUTOMATION_SOURCE_ID:'monash-adults',GITHUB_OUTPUT:output}});
  assert.equal(result.status,0,result.stderr);
  assert.equal((await readFile(output,'utf8')).trim(),'result='+expected,'GitHub receives the actual processor result');
 }
}finally{await rm(folder,{recursive:true,force:true});}
console.log('Delayed-only warnings, corrupt diagnostics, blocked dispatch, retained queue and honest GitHub outcomes passed.');
