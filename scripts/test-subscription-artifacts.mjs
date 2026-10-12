import {subscriptionBuildWire,readSubscriptionBuildInput,readBoundedSubscriptionBytes} from '../functions/_lib/subscription-wire.js';
import {onRequest as middleware} from '../functions/_middleware.js';
import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {readFile,readdir} from 'node:fs/promises';
import {subscriptionEvidence,queueSubscriptionArtifact,prepareSubscriptionExport,subscriptionEventPage,subscriptionDigest} from '../functions/_lib/subscription-artifacts.js';
import {onRequestGet as feed,buildD1SubscriptionFeed} from '../functions/api/feed.js';
import {onRequestPost as build} from '../functions/api/automation/subscription-build.js';
import {processSubscriptionBuild} from './process-subscription-build.mjs';
const sql=new DatabaseSync(':memory:');
for(const name of (await readdir(new URL('../migrations',import.meta.url))).filter(n=>n.endsWith('.sql')).sort())sql.exec(await readFile(new URL('../migrations/'+name,import.meta.url),'utf8'));
for(const definition of ["name TEXT NOT NULL DEFAULT ''","source_type TEXT NOT NULL DEFAULT ''","size INTEGER NOT NULL DEFAULT 0","last_modified INTEGER NOT NULL DEFAULT 0"])sql.exec('ALTER TABLE raw_roster_files ADD COLUMN '+definition);
sql.exec('CREATE INDEX IF NOT EXISTS idx_roster_events_file_doctor ON roster_events(file_id,doctor_key)');
const db={prepare(query){let args=[];const st={bind(...a){args=a;return st},async all(){if(db.beforeAll)db.beforeAll(query);const results=sql.prepare(query).all(...args);return {results,meta:{changes:0,rows_read:results.length,rows_written:0}}},async first(){return (await st.all()).results[0]||null},async run(){const r=sql.prepare(query).run(...args);return {meta:{changes:Number(r.changes),rows_written:Number(r.changes),rows_read:0}}}};return st},async batch(st){sql.exec('BEGIN');try{const r=[];for(const s of st)r.push(await (/^\s*(SELECT|WITH)/i.test(s.query||'')?s.all():s.run()));sql.exec('COMMIT');return r;}catch(e){sql.exec('ROLLBACK');throw e;}}};
// Native SELECT batches are needed by the existing fallback renderer.
db.prepare=((original)=>query=>{const st=original(query);st.query=query;return st;})(db.prepare);
class Bucket{
 map=new Map();fail='';
 async put(key,value,options={}){if(this.fail===key)throw Error('Injected publication failure');const old=this.map.get(key),condition=options.onlyIf;if(condition?.etagDoesNotMatch==='*'&&old||condition?.etagMatches&&old?.etag!==condition.etagMatches)return null;const bytes=typeof value==='string'?new TextEncoder().encode(value):value instanceof ReadableStream?new Uint8Array(await new Response(value).arrayBuffer()):new Uint8Array(value);if(options.sha256){const expected=[...new Uint8Array(options.sha256)].map(v=>v.toString(16).padStart(2,'0')).join('');if(await subscriptionDigest(bytes)!==expected)throw Error('SHA-256 checksum mismatch');}const entry={key,etag:crypto.randomUUID(),value:new TextDecoder().decode(bytes),customMetadata:options.customMetadata||{}};this.map.set(key,entry);return entry;}
 async get(key){const v=this.map.get(key);return v?{...v,body:v.value,json:async()=>JSON.parse(v.value),text:async()=>v.value}:null;}
 async head(key){return this.get(key);}
 async delete(key){this.map.delete(key);}
 async list({prefix,limit=1000}){return {objects:[...this.map.keys()].filter(k=>k.startsWith(prefix)).sort().slice(0,limit).map(key=>({key}))};}
}
const r2=new Bucket(),session={settings:{includeLocations:true},exportRange:{startDate:'2026-10-01',endDate:'2026-10-31',allFuture:false},customEvents:[]};
sql.exec(`INSERT INTO account_profiles(email,real_name,subscription_token) VALUES('a@test','Doctor A','legacy');
 INSERT INTO subscription_tokens(token,email,created_at) VALUES('token','a@test','fixture');
 INSERT INTO account_claims(email,source_type,doctor_key,display_name) VALUES('a@test','mmc','DOC A','Doctor A');
 INSERT INTO roster_people(person_id,preferred_display_name) VALUES('person:a','Doctor A');
 INSERT INTO account_people(email,person_id) VALUES('a@test','person:a');
 INSERT INTO roster_person_aliases(source_type,doctor_key,person_id,display_name) VALUES('mmc','DOC A','person:a','Doctor A'),('vhh','ALT A','person:a','Alternative A');
 INSERT INTO roster_files(id,name,source_type,parsed_at) VALUES('f','roster','mmc','one'),('g','roster','vhh','one');`);
sql.prepare('INSERT INTO account_states(email,session_json,updated_at) VALUES(?,?,?)').run('a@test',JSON.stringify(session),'one');
for(let i=0;i<130;i++){
 const date=i===0?'2022-01-01':'2026-10-'+String(i%28+1).padStart(2,'0'),source=i%2?'vhh':'mmc';
 const event={id:'e'+i,source:source.toUpperCase(),title:'Shift '+i,start:date+'T08:00:00+11:00',end:date+'T16:00:00+11:00',allDay:false,location:'Original',seniority:'Junior Registrar',rawValue:'AM'};
 sql.prepare('INSERT INTO roster_events(id,file_id,source_type,doctor_key,display_name,start_date,end_date,start_ts,end_ts,title,raw_value,seniority,location,all_day,time_label,event_json) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)').run(event.id,source==='mmc'?'f':'g',source,source==='mmc'?'DOC A':'ALT A','Doctor A',date,date,event.start,event.end,event.title,'AM',event.seniority,event.location,0,'08–16',JSON.stringify(event));
}
const env={ROSTER_DB:db,ROSTER_FILES:r2,SUBSCRIPTION_ARTIFACT_MODE:'ready',IDENTITY_REVIEW_ENABLED:'true',ROSTER_AUTOMATION_WRITES_ENABLED:'true',ROSTER_AUTOMATION_TOKEN:'secret',ROSTER_ACCOUNT_BUDGET_ENABLED:'true'};
sql.prepare("INSERT INTO roster_account_budget(utc_day,maximum_reads,maximum_writes,valid_until) VALUES(?,4000000,80000,?)").run(new Date().toISOString().slice(0,10),new Date(Date.now()+3600000).toISOString());
let actions=[];
async function call(input){
 actions.push(input.mode);
 const wire=subscriptionBuildWire(input);
 const context={env:{...env},data:{},request:new Request('https://test/api/automation/subscription-build',{method:'POST',headers:{authorization:'Bearer secret',...wire.headers},body:wire.body})};
 context.next=()=>build(context);
 const savedLog=console.log;console.log=()=>{};
 let response;try{response=await middleware(context);}finally{console.log=savedLog;}
 const result=await response.json();if(!response.ok)throw Object.assign(Error(result.error||result.status),{status:response.status,result});return result;
}
assert.equal((await build({env,request:new Request('https://test',{method:'POST',body:'{}'})})).status,401);
assert.equal(await subscriptionEvidence(db,'unknown'),null);
const initial=await subscriptionEvidence(db,'token');
let response=await feed({env,request:new Request('https://test/api/feed?token=token')});
const baseline=await response.text();assert.equal(response.status,200,baseline);assert.equal((baseline.match(/BEGIN:VEVENT/g)||[]).length,130);
assert.equal((await r2.list({prefix:'subscriptions/requests/'})).objects.length,1);
await queueSubscriptionArtifact(r2,'token','full',initial);assert.equal((await r2.list({prefix:'subscriptions/requests/'})).objects.length,1,'same inputs coalesce');
let result=await processSubscriptionBuild({call,admit:async()=>true});assert.equal(result.status,'complete');assert.equal(result.built,1);assert.ok(actions.filter(a=>a==='events').length===3,'all 130 shifts cross paged exports');
response=await feed({env,request:new Request('https://test/api/feed?token=token')});const cached=await response.text();
const withoutStamp=text=>text.replace(/^DTSTAMP:.*\r\n/gm,'');assert.equal(withoutStamp(cached),withoutStamp(baseline),'complete historical content, UID and ordering parity');
assert.equal((await r2.list({prefix:'subscriptions/requests/'})).objects.length,0);
const currentTag=response.headers.get('etag');assert.ok(currentTag);
const conditional=await feed({env,request:new Request('https://test/api/feed?token=token',{headers:{'If-None-Match':currentTag}})});
assert.equal(conditional.status,304);assert.equal(await conditional.text(),'');

// Range exports preserve the same full renderer, including personal overrides.
response=await feed({env,request:new Request('https://test/api/feed?token=token&view=range')});
const rangeBaseline=await response.text();
result=await processSubscriptionBuild({call,admit:async()=>true});assert.equal(result.status,'complete');
response=await feed({env,request:new Request('https://test/api/feed?token=token&view=range')});
assert.equal(withoutStamp(await response.text()),withoutStamp(rangeBaseline));
assert.equal((rangeBaseline.match(/BEGIN:VEVENT/g)||[]).length,129,'range excludes only the old historical shift');
// Lost publication responses leave a complete immutable object; replay consumes
// it without re-exporting events or changing existing UIDs.
const queueKey=await queueSubscriptionArtifact(r2,'token','full',initial);
const replayBegin=await call({mode:'start'});
assert.equal((await call({mode:'snapshot',leaseId:replayBegin.leaseId,queueKey})).status,'unchanged');
await call({mode:'finish',leaseId:replayBegin.leaseId});
// A source import defers before any reservation/export writes.
sql.exec("INSERT INTO roster_sync_runs(id,source_id,status,started_at) VALUES('pending','monash-adults','queued','fixture');");
await queueSubscriptionArtifact(r2,'token','full',initial);
const pendingBegin=await call({mode:'start'}),pendingNext=await call({mode:'next',leaseId:pendingBegin.leaseId});
const beforeReservation=sql.prepare('SELECT COUNT(*) n FROM roster_maintenance_receipts').get().n;
assert.equal((await call({mode:'snapshot',leaseId:pendingBegin.leaseId,queueKey:pendingNext.queueKey})).reason,'import-active');
assert.equal(sql.prepare('SELECT COUNT(*) n FROM roster_maintenance_receipts').get().n,beforeReservation);
await call({mode:'finish',leaseId:pendingBegin.leaseId});sql.exec("DELETE FROM roster_sync_runs WHERE id='pending'");
await processSubscriptionBuild({call,admit:async()=>true});
// An interrupted R2 publication keeps previous artifacts and the retry request.
sql.exec("UPDATE account_profiles SET real_name='Reviewed Doctor A' WHERE email='a@test'");
const changedEvidence=await subscriptionEvidence(db,'token');
await feed({env,request:new Request('https://test/api/feed?token=token')});
r2.fail=changedEvidence.key;
await assert.rejects(()=>processSubscriptionBuild({call,admit:async()=>true}),/Injected publication failure/);
assert.ok(await r2.head(initial.key),'failed publication never removes the previous complete object');
assert.equal(await r2.head(changedEvidence.key),null);
assert.ok((await r2.list({prefix:'subscriptions/requests/'})).objects.length);
r2.fail='';const interruptedLease=await r2.get('subscriptions/build-lease.json');
await r2.put('subscriptions/build-lease.json',JSON.stringify({...await interruptedLease.json(),until:0}),{onlyIf:{etagMatches:interruptedLease.etag}});
assert.equal((await processSubscriptionBuild({call,admit:async()=>true})).status,'complete');
// Oversized evidence explicitly keeps the original full subscription path.
for(let i=0;i<201;i++)sql.prepare('INSERT INTO custom_events(owner_email,id,title,start_date,end_date,all_day,include) VALUES(?,?,?,?,?,1,0)').run('a@test','disabled'+i,'Disabled custom','2026-10-01','2026-10-01');
response=await feed({env,request:new Request('https://test/api/feed?token=token')});assert.equal(response.status,200);
assert.equal((await response.text()).match(/BEGIN:VEVENT/g).length,130);
sql.exec("DELETE FROM custom_events WHERE id LIKE 'disabled%'");
// A concurrent identity edit during provisional preparation cannot authorize
// an event page or publication against the earlier evidence.
const provisionalEvidence=await subscriptionEvidence(db,'token');
await queueSubscriptionArtifact(r2,'token','range',await subscriptionEvidence(db,'token','range'));
// Force a new range snapshot instead of its previously completed immutable object.
const oldRange=await subscriptionEvidence(db,'token','range');await r2.delete(oldRange.key);
const raceBegin=await call({mode:'start'}),raceNext=await call({mode:'next',leaseId:raceBegin.leaseId});
db.beforeAll=query=>{if(query.includes('FROM account_people a JOIN roster_people p')){db.beforeAll=null;sql.exec("UPDATE roster_people SET preferred_display_name='Concurrent reviewed name' WHERE person_id='person:a'");}};
const racePin=await call({mode:'snapshot',leaseId:raceBegin.leaseId,queueKey:raceNext.queueKey});
const raceSnapshot=await call({mode:'prepare',leaseId:raceBegin.leaseId,queueKey:raceNext.queueKey,revision:racePin.revision});
assert.equal(raceSnapshot.status,'running');
await assert.rejects(()=>call({mode:'events',leaseId:raceBegin.leaseId,queueKey:raceNext.queueKey,revision:raceSnapshot.revision}),e=>e.status===409&&e.result.status==='stale');
await call({mode:'finish',leaseId:raceBegin.leaseId});
await processSubscriptionBuild({call,admit:async()=>true});
// Every material input invalidates. Unrelated account edits do not.
for(const change of [
 "UPDATE account_states SET session_json=json_set(session_json,'$.settings.includeLocations',0) WHERE email='a@test'",
 "UPDATE roster_files SET parsed_at='two' WHERE id='f'",
 "UPDATE roster_person_aliases SET display_name='New spelling' WHERE doctor_key='DOC A'",
 "INSERT INTO account_hospital_locations(email,source_type,location) VALUES('a@test','mmc','New place')",
 "INSERT INTO custom_events(owner_email,id,title,start_date,end_date,all_day) VALUES('a@test','custom','Leave','2026-10-01','2026-10-02',1)",
 "UPDATE roster_identity_revision SET revision=revision+1",
 "UPDATE account_profiles SET real_name='New name' WHERE email='a@test'"
]){const before=(await subscriptionEvidence(db,'token')).revision;sql.exec(change);assert.notEqual((await subscriptionEvidence(db,'token')).revision,before,change);}
const stable=(await subscriptionEvidence(db,'token')).revision;
sql.exec("INSERT INTO account_profiles(email,real_name) VALUES('b@test','Other');");assert.equal((await subscriptionEvidence(db,'token')).revision,stable);
response=await feed({env,request:new Request('https://test/api/feed?token=token')});assert.equal(response.status,200,'stale input uses complete current fallback');
const begin=await call({mode:'start'}),next=await call({mode:'next',leaseId:begin.leaseId}),base={leaseId:begin.leaseId,queueKey:next.queueKey};
const pinned=await call({...base,mode:'snapshot'});
const snapshot=await call({...base,mode:'prepare',revision:pinned.revision});
sql.exec("UPDATE subscription_tokens SET email='b@test' WHERE token='token'");
await assert.rejects(()=>call({...base,mode:'events',revision:snapshot.revision}),e=>e.status===409&&e.result.status==='stale');
response=await feed({env,request:new Request('https://test/api/feed?token=token',{headers:{'If-None-Match':currentTag}})});assert.notEqual(response.status,304);assert.ok(!(await response.text()).includes('Shift 0'),'reassigned token cannot serve prior owner artifact');
await call({mode:'finish',leaseId:begin.leaseId});
sql.exec("DELETE FROM subscription_tokens WHERE token='token'; UPDATE account_profiles SET subscription_token='' WHERE email='a@test'");
assert.equal((await feed({env,request:new Request('https://test/api/feed?token=token')})).status,404,'revocation blocks existing artifacts');
console.log('Subscription artifacts passed complete paged history/ICS parity, ownership, revocation, input invalidation, coalescing, stale exports and authentication.');
// Build-only rollout primes both stable URL views before enabling request reads.
sql.exec("INSERT INTO subscription_tokens(token,email,created_at) VALUES('new-token','a@test','fixture')");
const primeResult=await processSubscriptionBuild({call,admit:async()=>true,prime:true});assert.equal(primeResult.status,'complete');
for(const view of ['full','range']){
 const evidence=await subscriptionEvidence(db,'new-token',view);
 assert.ok(await r2.head(evidence.key),'priming publishes '+view+' before read activation');
}
assert.equal(sql.prepare("SELECT stop_reason FROM roster_account_budget").get().stop_reason,'');
assert.equal(sql.prepare("SELECT COUNT(*) n FROM roster_maintenance_receipts WHERE finished_at='' OR metadata_complete<>1").get().n,0,'actual middleware settles every known callback');
console.log('Build-only priming and shared middleware reservation/settlement integration passed.');
// A saved range without a start date retains the original full-calendar behavior.
sql.exec("UPDATE account_states SET session_json=json_set(session_json,'$.exportRange',json('{\"allFuture\":false,\"endDate\":\"2026-10-31\"}')) WHERE email='a@test'");
const noStartResponse=await feed({env:{...env,SUBSCRIPTION_ARTIFACT_MODE:'off'},request:new Request('https://test/api/feed?token=new-token&view=range')});
const noStartBaseline=await noStartResponse.text();
await feed({env,request:new Request('https://test/api/feed?token=new-token&view=range')});
assert.equal((await processSubscriptionBuild({call,admit:async()=>true})).status,'complete');
const noStartCached=await feed({env,request:new Request('https://test/api/feed?token=new-token&view=range')});
assert.equal(withoutStamp(await noStartCached.text()),withoutStamp(noStartBaseline),'missing range start never truncates history');
console.log('Provisional preparation fences and missing-start range parity passed.');
// Creator subscriptions follow the selected doctor rather than the Creator's
// own linked account. Changing selection invalidates the previous artifact.
sql.exec("UPDATE account_profiles SET role='creator' WHERE email='a@test'; UPDATE account_states SET session_json=json_set(session_json,'$.doctorKey','DOC A') WHERE email='a@test'");
const creatorResponse=await feed({env:{...env,SUBSCRIPTION_ARTIFACT_MODE:'off'},request:new Request('https://test/api/feed?token=new-token')});
const creatorBaseline=await creatorResponse.text();
await feed({env,request:new Request('https://test/api/feed?token=new-token')});
assert.equal((await processSubscriptionBuild({call,admit:async()=>true})).status,'complete');
const creatorCached=await feed({env,request:new Request('https://test/api/feed?token=new-token')});
assert.equal(withoutStamp(await creatorCached.text()),withoutStamp(creatorBaseline));
sql.exec("UPDATE account_states SET session_json=json_set(session_json,'$.doctorKey','UNRELATED') WHERE email='a@test'");
const selectedOther=await feed({env,request:new Request('https://test/api/feed?token=new-token')});
assert.ok(!(await selectedOther.text()).includes('Shift 0'),'changing Creator selection cannot serve the previous doctor');
console.log('Creator selection, approved cross-site aliases and artifact isolation passed.');
const unicodeWire=subscriptionBuildWire({mode:'publish',leaseId:'fixture',ics:'BEGIN:VCALENDAR\r\nSUMMARY:Zoë 医生\r\nEND:VCALENDAR\r\n',displayName:'Zoë 医生',digest:'fixture'});
const wireRequest=new Request('https://test',{method:'POST',headers:unicodeWire.headers,body:unicodeWire.body});
const wireInput=await readSubscriptionBuildInput(wireRequest);assert.equal(wireInput.rawCalendar,true);assert.equal(wireInput.displayName,'Zoë 医生');
assert.equal(new TextDecoder().decode(await readBoundedSubscriptionBytes(wireRequest)),'BEGIN:VCALENDAR\r\nSUMMARY:Zoë 医生\r\nEND:VCALENDAR\r\n');
assert.equal(await readSubscriptionBuildInput(new Request('https://test',{method:'POST',headers:{'Content-Type':'text/calendar','x-subscription-metadata':'bad!'},body:'invalid'})),null);
assert.throws(()=>subscriptionBuildWire({mode:'publish',ics:'x'.repeat(2*1024*1024+1)}),/safe byte bound/);
await assert.rejects(()=>readBoundedSubscriptionBytes(new Request('https://test',{method:'POST',body:'12345'}),4),/safe byte bound/);
assert.equal((await readSubscriptionBuildInput(new Request('https://test',{method:'POST',body:JSON.stringify({mode:'publish',rawCalendar:true})}))).rawCalendar,false,'JSON cannot forge the raw-body transport');
console.log('Raw calendar wire passed UTF-8 metadata/content parity, byte bounds and invalid-header handling.');
// A corrupt raw publication cannot create a current calendar artifact.
const otherEvidence=await subscriptionEvidence(db,'new-token');
await queueSubscriptionArtifact(r2,'new-token','full',otherEvidence);
const badBegin=await call({mode:'start'}),badNext=await call({mode:'next',leaseId:badBegin.leaseId}),badBase={leaseId:badBegin.leaseId,queueKey:badNext.queueKey};
const badPin=await call({...badBase,mode:'snapshot'});
const badPrepared=await call({...badBase,mode:'prepare',revision:badPin.revision});
await assert.rejects(()=>call({...badBase,mode:'publish',revision:badPrepared.revision,ics:'BEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n',digest:'0'.repeat(64)}),e=>e.status===400&&/digest mismatch/.test(e.result.error));
assert.equal(await r2.head(otherEvidence.key),null);
await call({mode:'finish',leaseId:badBegin.leaseId});
console.log('Corrupt raw publication preserves the existing artifact and pending request.');

sql.exec("UPDATE roster_people SET preferred_display_name='Toby O’Brien 医生' WHERE person_id='person:a'; UPDATE account_states SET session_json=json_set(session_json,'$.doctorKey','DOC A') WHERE email='a@test'");
const unicodeFeed=await feed({env:{...env,SUBSCRIPTION_ARTIFACT_MODE:'off'},request:new Request('https://test/api/feed?token=new-token')});
assert.equal(unicodeFeed.status,200);
assert.ok(unicodeFeed.headers.get('content-disposition').includes('filename*=UTF-8'));
assert.ok((await unicodeFeed.text()).includes('Toby O’Brien 医生'));
console.log('Conditional calendar responses retain ownership fences; Unicode calendar names remain valid in bodies and HTTP filenames.');
// The watchdog initiates priming even before the first subscription GET, then
// known admission deferral releases its lease while retaining the checkpoint.
await r2.delete('subscriptions/prime-progress.json');
for(const key of [...r2.map.keys()])if(key.startsWith('subscriptions/requests/')||key.startsWith('subscriptions/exports/'))await r2.delete(key);
await r2.put('subscriptions/build-lease.json',JSON.stringify({id:'idle',until:0}));
const savedFetch=globalThis.fetch;let dispatched;
env.GITHUB_ACTIONS_TOKEN='fixture';
try{
 globalThis.fetch=async(_url,options)=>{dispatched=JSON.parse(options.body);return new Response(null,{status:204});};
 assert.equal((await call({mode:'check'})).status,'dispatched');
 assert.equal(dispatched.inputs.prime,'true');
 const defer=await processSubscriptionBuild({call,admit:async()=>false,prime:true,dispatchId:dispatched.inputs.dispatch_id});
 assert.equal(defer.status,'deferred');
 assert.equal((await (await r2.get('subscriptions/build-lease.json')).json()).until,0);
 // A completed prime is an R2-only no-op and never reads/requeues all accounts.
 await r2.put('subscriptions/prime-progress.json',JSON.stringify({version:'subscription-v1',complete:true,cursor:''}));
 const receiptsBefore=sql.prepare('SELECT COUNT(*) n FROM roster_maintenance_receipts').get().n;
 assert.equal((await call({mode:'prime'})).status,'complete');
 assert.equal(sql.prepare('SELECT COUNT(*) n FROM roster_maintenance_receipts').get().n,receiptsBefore);
}finally{globalThis.fetch=savedFetch;delete env.GITHUB_ACTIONS_TOKEN;}
console.log('Automatic resumable priming passes startup dispatch, known deferral lease release and completed zero-D1 no-op.');

// Readiness is promoted only after complete priming and an empty work queue.
await r2.delete('subscriptions/readiness.json');
await r2.put('subscriptions/build-lease.json',JSON.stringify({id:'readiness-test',until:Date.now()+60000}));
await r2.put('subscriptions/prime-progress.json',JSON.stringify({version:'subscription-v1',complete:false,cursor:''}));
await call({mode:'finish',leaseId:'readiness-test'});
assert.equal(await r2.get('subscriptions/readiness.json'),null);
await r2.put('subscriptions/build-lease.json',JSON.stringify({id:'readiness-test',until:Date.now()+60000}));
await r2.put('subscriptions/prime-progress.json',JSON.stringify({version:'subscription-v1',complete:true,cursor:''}));
await r2.put('subscriptions/requests/pending.json','{}');
await call({mode:'finish',leaseId:'readiness-test'});
assert.equal(await r2.get('subscriptions/readiness.json'),null);
await r2.delete('subscriptions/requests/pending.json');
await r2.put('subscriptions/build-lease.json',JSON.stringify({id:'readiness-test',until:Date.now()+60000}));
await call({mode:'finish',leaseId:'readiness-test'});
assert.equal((await (await r2.get('subscriptions/readiness.json')).json()).accepted,true);
console.log('Readiness requires complete priming and an empty queue; partial/deferred work cannot activate cached delivery.');
