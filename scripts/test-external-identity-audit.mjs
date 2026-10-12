import assert from 'node:assert/strict';import {DatabaseSync} from 'node:sqlite';import {createHash} from 'node:crypto';import {readFile,readdir} from 'node:fs/promises';
import {planIdentityAudit} from './identity-audit-plan.mjs';
import {pinExternalIdentityAudit,applyExternalIdentityAuditBatch,finishExternalIdentityAudit} from '../functions/_lib/identity-audit-checkpoints.js';
import {beginMaintenanceAccounting,finishMaintenanceAccounting} from '../functions/_lib/roster-maintenance-budget.js';
import {onRequest} from '../functions/_middleware.js';
import {onRequestPost} from '../functions/api/automation/identity-audit.js';
import {continueExternalIdentityAudit} from './process-external-identity-audit.mjs';
const sqlite=new DatabaseSync(':memory:');for(const f of(await readdir('migrations')).filter(f=>f.endsWith('.sql')).sort())sqlite.exec(await readFile('migrations/'+f,'utf8'));
// Raw provider metadata existed before numbered migrations.
for(const d of ['name TEXT','last_modified INTEGER'])sqlite.exec('ALTER TABLE raw_roster_files ADD COLUMN '+d);
let writes=0,fail='',race=null,requestCount=0;const db={prepare(sql){return{args:[],bind(...args){this.args=args;return this;},async first(){return sqlite.prepare(sql).get(...this.args)||null;},async all(){const results=sqlite.prepare(sql).all(...this.args);return{results,meta:{rows_read:results.length,rows_written:0}};},async run(){if(fail&&sql.includes(fail)){fail='';throw Error('Injected interruption');}if(/^\s*(SELECT|WITH)\b/i.test(sql))return this.all();const r=sqlite.prepare(sql).run(...this.args);writes+=Number(r.changes);return{meta:{changes:Number(r.changes),rows_read:Number(r.changes),rows_written:Number(r.changes)}};}};},async batch(items){if(race){const f=race;race=null;f();}sqlite.exec('SAVEPOINT audit_test');try{const r=[];for(const i of items)r.push(await i.run());sqlite.exec('RELEASE audit_test');return r;}catch(e){sqlite.exec('ROLLBACK TO audit_test');sqlite.exec('RELEASE audit_test');throw e;}}};
const day=new Date().toISOString().slice(0,10);sqlite.prepare('INSERT INTO roster_account_budget(utc_day,maximum_reads,maximum_writes,valid_until)VALUES(?,?,?,?)').run(day,100000000,10000000,new Date(Date.now()+3600000).toISOString());
// Simulated table-change metadata exercises checkpoint behavior, not D1 billing.
const call=async fn=>{const before=writes;beginMaintenanceAccounting(db,'checkpoint-test-'+(++requestCount),true,{purpose:'identity'});try{return await fn();}finally{await finishMaintenanceAccounting(db,{rowsRead:0,rowsWritten:writes-before,metadataComplete:true},db);}};
const hash=x=>createHash('sha256').update(JSON.stringify(x)).digest('hex');
const rows=Array.from({length:6},(_,i)=>{const r={source_type:'mmc',doctor_key:'ANAME'+i+' DENSE',display_name:'Aname'+i+' Dense',person_id:'person:test-'+i,person_status:'active',audited_fingerprint:''};r.fingerprint=hash([r.source_type,r.doctor_key,r.display_name,r.person_id]);return r;});
for(const r of rows){sqlite.prepare("INSERT INTO roster_people(person_id,preferred_display_name,status)VALUES(?,?,'active')").run(r.person_id,r.display_name);sqlite.prepare('INSERT INTO roster_person_aliases(source_type,doctor_key,display_name,person_id)VALUES(?,?,?,?)').run(r.source_type,r.doctor_key,r.display_name,r.person_id);sqlite.prepare("INSERT INTO roster_identity_features(source_type,doctor_key,display_name,normalized_key,given_block,surname_block,fingerprint,updated_at)VALUES(?,?,?,'fixture','fixture','fixture',?,'fixture')").run(r.source_type,r.doctor_key,r.display_name,r.fingerprint);}

const objects=new Map();let etag=0,failWake=false,loseCAS=false;
const r2={async get(key){const o=objects.get(key);return o?{etag:o.etag,json:async()=>JSON.parse(o.text)}:null;},async put(key,text,options={}){if(key==='identity/revision.json'&&failWake){failWake=false;throw Error('Wake interrupted');}const old=objects.get(key);if(loseCAS){loseCAS=false;return null;}if(options.onlyIf?.etagMatches&&old?.etag!==options.onlyIf.etagMatches||options.onlyIf?.etagDoesNotMatch==='*'&&old)return null;const value={text,etag:String(++etag)};objects.set(key,value);return value;}};
const env={ROSTER_DB:db,ROSTER_FILES:r2,ROSTER_AUTOMATION_TOKEN:'test-runner',ROSTER_WATCHDOG_TOKEN:'test-watchdog',IDENTITY_EXTERNAL_AUDIT_ENABLED:'true',IDENTITY_REVIEW_ENABLED:'true',ROSTER_ACCOUNT_BUDGET_ENABLED:'true'};
const http=async(body,token='test-runner')=>{const context={env,data:{},request:new Request('https://example.test/api/automation/identity-audit',{method:'POST',headers:{authorization:'Bearer '+token,'Content-Type':'application/json'},body:JSON.stringify(body)})};context.next=()=>onRequestPost(context);return onRequest(context);};
const request=async body=>{const response=await http(body),data=await response.json();if(response.status===409&&data.status==='stale')throw Object.assign(Error('Stale'),{staleEvidence:true});if(!response.ok&&data.status!=='lease-expired')throw Error('HTTP '+response.status);return data;};
assert.equal((await http({mode:'start'},'bad')).status,401);
assert.equal((await http({mode:'start'},'test-watchdog')).status,401);
env.IDENTITY_EXTERNAL_AUDIT_ENABLED='false';assert.equal((await request({mode:'start'})).status,'external-disabled');env.IDENTITY_EXTERNAL_AUDIT_ENABLED='true';
loseCAS=true;assert.equal((await request({mode:'start'})).status,'busy');
let dropped=false;
await assert.rejects(continueExternalIdentityAudit({request:async input=>{const data=await request(input);if(input.kind==='features'&&!dropped){dropped=true;throw Error('Lost successful response');}return data;}}),/Lost successful response/);
const saved=JSON.parse(objects.get('identity/external-audit-progress.json').text);
assert.equal(saved.status,'running');assert.equal(sqlite.prepare('SELECT next_feature_batch FROM external_identity_audit_jobs WHERE run_id=?').get(saved.runId).next_feature_batch,1);
assert.equal((await request({mode:'start'})).status,'busy','another runner cannot steal an active lease');
const first=sqlite.prepare('SELECT pair_key FROM roster_identity_candidates LIMIT 1').get().pair_key;sqlite.prepare("UPDATE roster_identity_candidates SET status='rejected',reviewer='human' WHERE pair_key=?").run(first);
failWake=true;await assert.rejects(continueExternalIdentityAudit({request,dispatchId:saved.leaseToken}),/HTTP 500/);
assert.equal(sqlite.prepare('SELECT status FROM external_identity_audit_jobs WHERE run_id=?').get(saved.runId).status,'complete','wake failure retains a completed durable job');
assert.equal((await continueExternalIdentityAudit({request,dispatchId:saved.leaseToken})).status,'complete');
assert.equal(sqlite.prepare('SELECT status FROM roster_identity_candidates WHERE pair_key=?').get(first).status,'rejected');
assert.equal(sqlite.prepare('SELECT COUNT(*) n FROM roster_identity_features WHERE audited_fingerprint<>fingerprint').get().n,0);
assert.equal((await request({mode:'start'})).status,'unchanged');
// Explicit bounds release the lease and resume the same immutable plan.
sqlite.exec("UPDATE roster_identity_features SET audited_fingerprint=''");const paused=await continueExternalIdentityAudit({request,maxSteps:4});assert.equal(paused.status,'deferred');const pending=JSON.parse(objects.get('identity/external-audit-progress.json').text);assert.equal(pending.status,'pending');assert(pending.runId);assert.equal((await continueExternalIdentityAudit({request})).status,'complete');
// Revision changes during an exported snapshot cannot apply an old plan.
sqlite.exec("UPDATE roster_identity_features SET audited_fingerprint=''");const stale=await continueExternalIdentityAudit({request:async input=>{const data=await request(input);if(input.mode==='snapshot')sqlite.exec('UPDATE roster_identity_revision SET revision=revision+1');return data;}});assert.equal(stale.status,'deferred');assert.equal(JSON.parse(objects.get('identity/external-audit-progress.json').text).runId,'');
const originalFetch=globalThis.fetch;let dispatches=0;
env.GITHUB_ACTIONS_TOKEN='test-dispatch';
try{globalThis.fetch=async(url,options)=>{assert.match(String(url),/identity-maintenance.yml\/dispatches$/);const body=JSON.parse(options.body);assert.equal(body.inputs.mode,'audit');assert(body.inputs.audit_dispatch_id);dispatches++;return new Response(null,{status:204});};assert.equal((await request({mode:'check'})).status,'queued');assert.equal((await request({mode:'check'})).status,'busy');assert.equal(dispatches,1,'coalesce concurrent UI/watchdog checks');const dispatched=JSON.parse(objects.get('identity/external-audit-progress.json').text);assert.equal((await continueExternalIdentityAudit({request,dispatchId:dispatched.leaseToken})).status,'complete');assert.equal((await request({mode:'check'})).status,'unchanged');assert.equal(dispatches,1,'unchanged cache starts no workflow');}finally{globalThis.fetch=originalFetch;}
assert.equal(sqlite.prepare('SELECT COUNT(*) n FROM roster_events').get().n,0);
console.log('External API/Node integration passed authentication, CAS collision, lost-response/feature replay, wake failure, rejection preservation, bounded yield/resume and stale snapshot rejection.');
