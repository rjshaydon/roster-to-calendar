// Connected synthetic acceptance: requests, ownership, publication, the actual
// account calendar reader and both subscription cache/fallback paths. No network.
import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {readFile,readdir} from 'node:fs/promises';
import {seedSupportFixture,expandSupportFixture,boundSupportFixture} from './support-request-fixture.mjs';
import {onRequestPost} from '../functions/api/requests.js';
import {onRequestGet as feed} from '../functions/api/feed.js';
import {subscriptionEvidence,SUBSCRIPTION_ARTIFACT_VERSION} from '../functions/_lib/subscription-artifacts.js';
import {publishIdentityOperation} from '../functions/_lib/doctor-identity.js';
import {loadFastAccountSnapshotPayload} from '../functions/api/state.js';
import {customEventForAccount} from '../functions/_lib/d1-calendar.js';
const sql=new DatabaseSync(':memory:');sql.exec('PRAGMA foreign_keys=ON');
for(const n of(await readdir('migrations')).filter(n=>n.endsWith('.sql')).sort())sql.exec(await readFile('migrations/'+n,'utf8'));
// The historical bootstrap index exists in production; fixture mirrors it.
sql.exec('CREATE INDEX IF NOT EXISTS idx_roster_events_file_doctor ON roster_events(file_id,doctor_key)');
const db={prepare(query){return{query,args:[],bind(...args){this.args=args;return this;},async all(){return{results:sql.prepare(query).all(...this.args)};},async first(){return sql.prepare(query).get(...this.args)||null;},async run(){if(/^\s*(SELECT|WITH)/i.test(query))return this.all();return{meta:{changes:Number(sql.prepare(query).run(...this.args).changes)}};}};},async batch(items){sql.exec('SAVEPOINT acceptance');try{const out=[];for(const s of items)out.push(await s.run());sql.exec('RELEASE acceptance');return out;}catch(e){sql.exec('ROLLBACK TO acceptance;RELEASE acceptance');throw e;}}};
const actors=await seedSupportFixture(db);await expandSupportFixture(db);
class Bucket{
 objects=new Map();fail=false;
 async put(key,value,options={}){if(this.fail)throw Error('Synthetic interrupted publication');const bytes=typeof value==='string'?new TextEncoder().encode(value):new Uint8Array(value);this.objects.set(key,{bytes,customMetadata:options.customMetadata||{},etag:crypto.randomUUID()});return{etag:this.objects.get(key).etag};}
 async get(key){const entry=this.objects.get(key);if(!entry)return null;return{...entry,body:entry.bytes,arrayBuffer:async()=>entry.bytes.buffer.slice(entry.bytes.byteOffset,entry.bytes.byteOffset+entry.bytes.byteLength),json:async()=>JSON.parse(new TextDecoder().decode(entry.bytes))};}
 async head(key){return this.get(key);}
}
const r2=new Bucket(),cache=new Bucket(),env={ROSTER_DB:db,ROSTER_FILES:r2,ROSTER_CACHE:cache,IDENTITY_REVIEW_ENABLED:'true',SUBSCRIPTION_ARTIFACT_MODE:'ready',BOUNDED_MANUAL_ROSTER_ENABLED:'true',ROSTER_AUTH_CONFIG:JSON.stringify({enabled:true,sessionsEnabled:true,requestsEnabled:true,requestTransfersEnabled:true}),AUTH_LOGIN_LIMITER:{async limit(){return{success:true};}}};
const aliases=sql.prepare("SELECT source_type,doctor_key FROM roster_person_aliases ORDER BY source_type").all(),events=[];
for(const a of aliases){
 const event={id:'synthetic-'+a.source_type,source:a.source_type.toUpperCase(),doctorKey:a.doctor_key,doctor:'Synthetic Doctor',title:'Synthetic '+a.source_type+' shift',start:'2026-10-18T08:00:00+11:00',end:'2026-10-18T16:00:00+11:00',seniority:'Junior Registrar',rawValue:'AM'};events.push(event);
 sql.prepare('INSERT INTO roster_files(id,name,source_type,parsed_at) VALUES(?,?,?,?)').run(a.source_type,'Synthetic roster',a.source_type,'fixture');
 sql.prepare('INSERT INTO roster_events(id,file_id,source_type,doctor_key,display_name,start_date,end_date,start_ts,end_ts,title,raw_value,seniority,location,all_day,time_label,event_json) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)').run(event.id,a.source_type,a.source_type,a.doctor_key,'Synthetic Doctor','2026-10-18','2026-10-18',event.start,event.end,event.title,'AM',event.seniority,'',0,'AM',JSON.stringify(event));
}
const personal={id:'personal',source:'CUSTOM',title:'Holder private appointment',start:'2026-10-19T10:00:00+11:00',end:'2026-10-19T11:00:00+11:00'};
sql.prepare('INSERT INTO custom_events(owner_email,id,title,start_date,end_date,start_time,end_time) VALUES(?,?,?,?,?,?,?)').run(actors.holder.email,'personal',personal.title,'2026-10-19','2026-10-19','10:00','11:00');
const aliasInput=aliases.map(a=>({sourceType:a.source_type,key:a.doctor_key})),session={doctorKey:'SYNTHETIC DOCTOR',settings:{},customEvents:[customEventForAccount(sql.prepare('SELECT * FROM custom_events WHERE owner_email=?').get(actors.holder.email),actors.holder.email)]};
sql.prepare('INSERT INTO account_states(email,session_json,updated_at) VALUES(?,?,?)').run(actors.holder.email,JSON.stringify(session),'fixture');
sql.prepare('INSERT INTO canonical_doctors(canonical_key,display_name,aliases_json) VALUES(?,?,?)').run(session.doctorKey,'Synthetic Doctor',JSON.stringify(aliasInput));
sql.prepare("INSERT INTO snapshot_registry(owner_type,owner_id,doctor_key,range_key,status,artifact_key,built_revision) VALUES('user-account',?,?,'2026-10-01..2026-10-31','ready','synthetic-cache','fixture')").run(actors.holder.email,session.doctorKey);
await cache.put('synthetic-cache',JSON.stringify({preview:{events:[...events,personal]},session,doctorOptions:[{key:session.doctorKey,aliases:aliasInput}]}),{customMetadata:{revision:'fixture'}});
async function request(actor,input){const r=await onRequestPost({env,request:new Request('https://fixture.test/api/requests',{method:'POST',headers:{Origin:'https://fixture.test',Cookie:actors[actor].cookie,'Content-Type':'application/json'},body:JSON.stringify({...input,expectedAccount:actors[actor].email})})});const d=await r.json();assert.equal(r.status,200,JSON.stringify(d));return d;}
const mutate=(actor,input)=>request(actor,{...input,mutationId:crypto.randomUUID()});
async function getFeed(actor,etag=''){const response=await feed({env,request:new Request('https://fixture.test/api/feed?token=synthetic-feed-'+actor,{headers:etag?{'If-None-Match':etag}:{}})});const body=await response.text();assert.equal(response.status,200,body);return{body,etag:response.headers.get('etag')};}
async function accountCalendar(){const claims=sql.prepare('SELECT source_type AS sourceType,doctor_key AS key FROM account_claims WHERE email=?').all(actors.holder.email);return loadFastAccountSnapshotPayload({env},{targetRecord:{email:actors.holder.email,realName:'Synthetic holder',role:'user',claims,state:{session}},prepared:{role:'user',claims,state:{session}},startDate:'2026-10-01',endDate:'2026-10-31'});}
const before=await getFeed('holder');for(const a of aliases)assert.match(before.body,new RegExp('Synthetic '+a.source_type+' shift'));assert.match(before.body,/Holder private appointment/);
const evidence=await subscriptionEvidence(db,'synthetic-feed-holder');await r2.put(evidence.key,before.body,{customMetadata:{version:SUBSCRIPTION_ARTIFACT_VERSION,revision:evidence.revision}});
const cached=await getFeed('holder');assert.ok(cached.etag);assert.equal(cached.body,before.body);
assert.equal((await accountCalendar()).snapshot.preview.events.filter(e=>e.id.startsWith('synthetic-')).length,4);
const rosterBefore=JSON.stringify(sql.prepare('SELECT * FROM roster_events ORDER BY id').all()),privateBefore=JSON.stringify(sql.prepare('SELECT * FROM custom_events').all());
const created=await mutate('user',{action:'create',kind:'identity',subject:'Synthetic access review',body:'Private requester explanation',alias:aliasInput[0]});
let detail=await request('admin',{action:'detail',caseId:created.caseId});const preview=await request('admin',{action:'previewResolution',caseId:created.caseId,version:detail.case.version,kind:'transfer'});
const changed=await mutate('admin',{action:'commitResolution',caseId:created.caseId,version:detail.case.version,kind:'transfer',previewToken:preview.preview.previewToken,reason:'Synthetic reviewed transfer'});
// Simulate interrupted publication. Access must change at the D1 transaction,
// even when a stale legacy artifact is subsequently marked ready again.
r2.fail=true;assert.equal((await publishIdentityOperation(db,r2,changed.operationId)).status,'failed');r2.fail=false;
sql.prepare("UPDATE snapshot_registry SET status='ready' WHERE artifact_key='synthetic-cache'").run();
const oldCalendar=await accountCalendar();assert.equal(oldCalendar.snapshot.preview.events.some(e=>e.id.startsWith('synthetic-')),false,'cached doctor selection cannot restore revoked access');assert.ok(oldCalendar.snapshot.preview.events.some(e=>e.title===personal.title),'private custom events remain with their owner');
const removed=await getFeed('holder',cached.etag);assert.doesNotMatch(removed.body,/Synthetic (mmc|mch|ddh|vhh) shift/);assert.match(removed.body,/Holder private appointment/);
const received=await getFeed('user');for(const a of aliases)assert.match(received.body,new RegExp('Synthetic '+a.source_type+' shift'));assert.doesNotMatch(received.body,/Holder private appointment/);
assert.equal((await publishIdentityOperation(db,r2,changed.operationId)).status,'complete');assert.equal((await publishIdentityOperation(db,r2,changed.operationId)).status,'complete','publication retry is idempotent');
assert.ok(r2.objects.has('identity/revision.json'));
detail=await request('admin',{action:'detail',caseId:created.caseId});const reverse=await request('admin',{action:'previewResolution',caseId:created.caseId,version:detail.case.version,kind:'reverse'});
const restored=await mutate('admin',{action:'commitResolution',caseId:created.caseId,version:detail.case.version,kind:'reverse',previewToken:reverse.preview.previewToken,reason:'Synthetic reviewed reversal'});
assert.equal((await publishIdentityOperation(db,r2,restored.operationId)).status,'complete');
const restoredFeed=await getFeed('holder');for(const a of aliases)assert.match(restoredFeed.body,new RegExp('Synthetic '+a.source_type+' shift'));assert.match(restoredFeed.body,/Holder private appointment/);assert.doesNotMatch((await getFeed('user')).body,/Synthetic (mmc|mch|ddh|vhh) shift/);
assert.equal(JSON.stringify(sql.prepare('SELECT * FROM roster_events ORDER BY id').all()),rosterBefore,'historical roster shifts and grades never change');assert.equal(JSON.stringify(sql.prepare('SELECT * FROM custom_events').all()),privateBefore,'private events never move');
for(const count of [8,16]){
 await boundSupportFixture(db,count);detail=await request('admin',{action:'detail',caseId:created.caseId});
 const p=await request('admin',{action:'previewResolution',caseId:created.caseId,version:detail.case.version,kind:'transfer'});
 assert.equal(p.preview.before.claimOwners,undefined,'raw claim records stay in the private prepared snapshot');assert.equal(p.preview.before.aliases.length,count);
 const stored=JSON.parse(sql.prepare('SELECT payload_json FROM support_ownership_previews WHERE preview_id=?').get(p.preview.previewToken.slice(9)).payload_json);assert.equal(stored.before.claimOwners.length,64);
 await mutate('admin',{action:'commitResolution',caseId:created.caseId,version:detail.case.version,kind:'transfer',previewToken:p.preview.previewToken,reason:'Synthetic full-bound transfer'});
 detail=await request('admin',{action:'detail',caseId:created.caseId});const undo=await request('admin',{action:'previewResolution',caseId:created.caseId,version:detail.case.version,kind:'reverse'});
 await mutate('admin',{action:'commitResolution',caseId:created.caseId,version:detail.case.version,kind:'reverse',previewToken:undo.preview.previewToken,reason:'Synthetic full-bound reversal'});
}
console.log('Connected four-site ownership acceptance passed transfer/reversal, failed/retried publication, actual cached account reader, subscription cache/fallback revocation and preserved personal events/historical grades. Synthetic data only.');
