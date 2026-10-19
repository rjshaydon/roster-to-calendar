import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {readFile,readdir} from 'node:fs/promises';
import {beginBoundedRosterImport,stageBoundedRosterBatch,prepareBoundedRosterPresence} from '../functions/_lib/roster-import-staging.js';
import {planRosterImportBatches} from '../functions/_lib/roster-import-batches.js';
import {attachRosterCompactPlan} from './roster-compact-plan.mjs';
import {executeRosterRebase} from './roster-staging-reuse.mjs';
import {onRequestPost as derived} from '../functions/api/automation/derived.js';
import {beginMaintenanceAccounting,rosterProcessingHeadroom} from '../functions/_lib/roster-maintenance-budget.js';
import {readRosterRebasePage} from '../functions/_lib/roster-staging-reuse.js';
import {listQueuedRosterSyncRuns,replaceDerivedRosterFile,refreshFacilityOverviewMaterializationForFile} from '../functions/_lib/d1-calendar.js';
const sqlite=new DatabaseSync(':memory:');
for(const n of(await readdir('migrations')).filter(n=>n.endsWith('.sql')).sort())sqlite.exec(await readFile('migrations/'+n,'utf8'));
for(const def of ['name TEXT','source_type TEXT','size INTEGER','last_modified INTEGER'])sqlite.exec('ALTER TABLE raw_roster_files ADD COLUMN '+def);
let fail='',calls=0;
const sqls=[];
const db={prepare(sql){sqls.push(sql);return {args:[],bind(...args){assert.ok(args.length<=100);this.args=args;return this;},async first(){return sqlite.prepare(sql).get(...this.args)||null;},async all(){return {results:sqlite.prepare(sql).all(...this.args)};},async run(){if(fail&&sql.includes(fail)){fail='';throw Error('Injected failure');}const r=sqlite.prepare(sql).run(...this.args);return {meta:{changes:Number(r.changes)}};}};},async batch(statements){sqlite.exec('SAVEPOINT reuse_test');try{const r=[];for(const s of statements)r.push(await s.run());sqlite.exec('RELEASE reuse_test');return r;}catch(e){sqlite.exec('ROLLBACK TO reuse_test; RELEASE reuse_test');throw e;}}};
const data=new Map(),bucket={async put(k,v){data.set(k,v);},async get(k){return data.has(k)?{json:async()=>JSON.parse(data.get(k))}:null;}};
const day=new Date().toISOString().slice(0,10);
sqlite.prepare('INSERT INTO roster_account_budget(utc_day,maximum_reads,maximum_writes,valid_until) VALUES(?,?,?,?)').run(day,50000000,5000000,new Date(Date.now()+3600000).toISOString());
const sourceId='monash-adults',doctors=Array.from({length:48},(_,i)=>({key:'SYNTHETIC '+i,displayName:'Synthetic '+i,seniority:'Junior Registrar'}));
const old={file:{id:'retained-stage',sourceId,sourceType:'mmc',name:'AdultTerm4.2026.xlsx',contentHash:'a'.repeat(64),lastModified:1},doctors,
 eventsByDoctor:Object.fromEntries(doctors.map(d=>[d.key,Array.from({length:90},(_,i)=>{
 const date=new Date(Date.UTC(2026,10,2+i)).toISOString().slice(0,10);
 return {id:'day-'+i,source:'MMC',title:'Day',rawValue:'AM',seniority:'Junior Registrar',start:date+'T08:00:00',end:date+'T16:00:00'};
 })])),issuesByDoctor:{}};
const previous={...structuredClone(old),file:{...old.file,id:'active-current',name:'AdultTerm3.2026.xlsx',active:true},eventsByDoctor:Object.fromEntries(doctors.map(d=>[d.key,[{id:'current',source:'MMC',title:'Night',start:'2026-10-19T23:00:00',end:'2026-10-20T08:00:00'}]]))};
await replaceDerivedRosterFile(db,previous.file,previous.doctors,previous.eventsByDoctor,{});
for(const [id,file,version,date] of [['old-run',old.file,'7','2026-10-19T00:00:00Z'],['new-run',{...old.file,id:'incoming-latest'},'19','2026-10-19T01:00:00Z'],['current-run',{...old.file,id:'incoming-current',name:'AdultTerm3.2026.xlsx'},'520','2026-10-18T23:59:00Z']]){
 sqlite.prepare('INSERT INTO raw_roster_files(file_id,name,source_type,last_modified) VALUES(?,?,?,?)').run(file.id,file.name,'mmc',Date.parse(date));
 sqlite.prepare("UPDATE raw_roster_files SET object_key='synthetic-only' WHERE file_id=?").run(file.id);
 sqlite.prepare('INSERT INTO roster_sync_runs(id,source_id,file_id,source_file_id,provider_version,content_hash,status,started_at) VALUES(?,?,?,?,?,?,?,?)').run(id,sourceId,file.id,file.id,version,file.contentHash,'queued',date);
}
const account=()=>beginMaintenanceAccounting(db,'reuse-'+(++calls),true);
const plan=await attachRosterCompactPlan(await planRosterImportBatches(old,{maximumOccurrences:32}),old);
account();await beginBoundedRosterImport(db,'old-run',old.file,plan,{allowReplacement:true});
for(const batch of plan.batches){account();await stageBoundedRosterBatch(db,'old-run',old.file,plan.revision,batch);}
for(const batch of plan.batches.slice(0,5)){account();await prepareBoundedRosterPresence(db,'old-run',old.file,plan.revision,batch);}
const oldFacts=sqlite.prepare('SELECT id,event_json FROM roster_events WHERE file_id=? ORDER BY id').all(old.file.id);
const latest=structuredClone(old);latest.file={...old.file,id:'incoming-latest',contentHash:'b'.repeat(64),lastModified:Date.parse('2026-10-19T01:00:00Z')};
sqlite.prepare('UPDATE roster_sync_runs SET content_hash=? WHERE id=?').run(latest.file.contentHash,'new-run');
latest.eventsByDoctor[doctors[1].key].splice(30,1);latest.eventsByDoctor[doctors[2].key][40].title='Changed day';
latest.eventsByDoctor[doctors[3].key].push({...latest.eventsByDoctor[doctors[3].key][50],id:'added'});
const env={ROSTER_DB:db,ROSTER_FILES:bucket,ROSTER_AUTOMATION_TOKEN:'synthetic-only',ROSTER_AUTOMATION_WRITES_ENABLED:'true',ROSTER_AUTOMATION_QUEUE_ENABLED:'true',ROSTER_AUTOMATION_SOURCE_ALLOWLIST:sourceId,ROSTER_AUTOMATION_REVIEWED_FACT_LIMIT:'1250',ROSTER_AUTOMATION_BOUNDED_IMPORT_ENABLED:'true',ROSTER_BOUNDED_REPLACEMENT_ENABLED:'true',ROSTER_STAGING_REUSE_SOURCES:sourceId};
let stopAfterBatch=false,activeRun='new-run';
async function callback(step){
 account();
 const r=await derived({env,request:new Request('https://fixture.test/api/automation/derived',{method:'POST',headers:{Authorization:'Bearer synthetic-only','Content-Type':'application/json'},body:JSON.stringify({...step,runId:activeRun,sourceId,file:latest.file})})});
 const result=await r.json();assert.equal(r.status,200,JSON.stringify(result));
 if(stopAfterBatch&&step.phase==='bounded-rebase-stage'&&!result.deferred){stopAfterBatch=false;throw Error('Lost response after saved batch');}
 return result;
}
const latestPlan=await attachRosterCompactPlan(await planRosterImportBatches(latest,{maximumOccurrences:32}),latest);
const located=await callback({phase:'bounded-begin',deferStaging:true,manifest:latestPlan.manifest,revision:latestPlan.revision});
assert.equal(located.mode,'rebase');assert.equal(located.fileId,old.file.id);
const queued=await listQueuedRosterSyncRuns(db,sourceId,2);assert.equal(queued[0].id,'new-run','unpublished term is not blocked behind an older current-term correction');
const range={startDate:latestPlan.startDate,rosterEndDate:latestPlan.rosterEndDate};
stopAfterBatch=true;
await assert.rejects(executeRosterRebase(latest,located,range,callback),/Lost response/);
assert.equal(sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(old.file.id).active,0);
assert.equal(sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(previous.file.id).active,1);
const cursor=sqlite.prepare('SELECT next_batch FROM roster_rebase_jobs').get().next_batch;assert.equal(cursor,1);
// An older worker cannot continue modifying the reused file.
account();await assert.rejects(prepareBoundedRosterPresence(db,'old-run',old.file,plan.revision,plan.batches[5]),/malformed JSON/);
// Check rollback of a batch, then resume the exact immutable saved plan.
fail='UPDATE roster_rebase_jobs SET next_batch';
await assert.rejects(executeRosterRebase(latest,located,range,callback),/Injected failure/);
assert.equal(sqlite.prepare('SELECT next_batch FROM roster_rebase_jobs').get().next_batch,1);
// A later autosave can reuse even an interrupted rebase. The old worker loses
// its inactive-file lease; its delayed completion cannot activate old facts.
sqlite.prepare("INSERT INTO raw_roster_files(file_id,name,source_type,last_modified,object_key) VALUES('incoming-third','AdultTerm4.2026.xlsx','mmc',?,'synthetic-only')").run(Date.parse('2026-10-19T02:00:00Z'));
sqlite.prepare("INSERT INTO roster_sync_runs(id,source_id,file_id,source_file_id,provider_version,content_hash,status,started_at) VALUES('third-run',?,'incoming-third','incoming-third','20',?,'queued','2026-10-19T02:00:00Z')").run(sourceId,latest.file.contentHash);
assert.equal((await callback({phase:'bounded-rebase-stage',batchIndex:1,patchRevision:sqlite.prepare("SELECT plan_revision FROM roster_rebase_jobs WHERE run_id='new-run'").get().plan_revision})).superseded,true);
activeRun='third-run';latest.file.id='incoming-third';
let inspected=false;
await executeRosterRebase(latest,located,range,async step=>{
 const result=await callback(step);
 if(step.phase==='bounded-rebase-start'&&!inspected){inspected=true;account();await assert.rejects(readRosterRebasePage(db,{id:'new-run',sourceId},{section:'events'}),/malformed JSON/);}
 return result;
});
assert.equal(sqlite.prepare('SELECT status FROM roster_sync_runs WHERE id=?').get('third-run').status,'success');
assert.equal(sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(old.file.id).active,1);
assert.equal(sqlite.prepare('SELECT active FROM roster_files WHERE id=?').get(previous.file.id).active,1);
const after=sqlite.prepare('SELECT id,event_json FROM roster_events WHERE file_id=? ORDER BY id').all(old.file.id);
assert.equal(after.length,4320);assert.equal(after.filter(r=>oldFacts.some(o=>o.id===r.id&&o.event_json===r.event_json)).length,4318,'4318 unchanged events were reused without rewriting');
const material=await refreshFacilityOverviewMaterializationForFile(db,old.file.id,{dryRun:true,contentRevision:sqlite.prepare("SELECT plan_revision FROM roster_rebase_jobs WHERE run_id='third-run'").get().plan_revision});assert.equal(material.proposedWrites,0,'compact facts agree with the normal publication builder');
assert.equal(sqlite.prepare('SELECT COUNT(*) n FROM roster_events e JOIN roster_daily_presence d ON d.event_id=e.id WHERE e.file_id=?').get(old.file.id).n,4320);
assert.equal((await callback({phase:'bounded-rebase-prepare',patchRevision:sqlite.prepare("SELECT plan_revision FROM roster_rebase_jobs WHERE run_id='third-run'").get().plan_revision})).completed,true,'completed-run replay cannot activate twice');
assert.ok(sqls.some(s=>s.includes('idx_roster_daily_presence_event_id')));
const gate={valid_until:new Date(Date.now()+60000).toISOString(),stop_reason:'',maximum_reads:100000,allocated_reads:0,maximum_writes:2100,allocated_writes:100};
assert.equal(rosterProcessingHeadroom(gate),false,'a nonzero but unusable grant must not launch/parse');
gate.maximum_writes=3000;assert.equal(rosterProcessingHeadroom(gate),true);gate.stop_reason='cost-overrun:request';assert.equal(rosterProcessingHeadroom(gate),false);
console.log('Staging reuse passed actual derived HTTP routing, inactive baseline locking, partial presence completion, 4318 unchanged rows, interrupted/rolled-back/superseded batch recovery, current-term preservation, compact parity, queue fairness and affordability. Synthetic data only.');
