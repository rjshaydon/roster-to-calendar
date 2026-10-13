import {prepareLegacyCalendarCoverage} from './process-legacy-calendar-coverage.mjs';
import {legacyCoverageKey,summarizeLegacyCoverage,legacyHistoryIsCovered} from '../functions/_lib/legacy-calendar-coverage.js';
import {DatabaseSync} from 'node:sqlite';
import {readdir,readFile} from 'node:fs/promises';
import {onRequestPost as statePost} from '../functions/api/state.js';
import {snapshotRegistryRangeKey,loadSnapshotRegistryEntry,loadCachedSnapshot} from '../functions/_lib/d1-calendar.js';
import {loadCalendarPresentationInputs} from '../functions/_lib/d1-calendar.js';
import {doctorRangeGaps,doctorSourceKey,doctorSourceCurrentKey,doctorCatalogueShard} from '../functions/_lib/doctor-source-artifacts.js';
import {publishDoctorSource} from './process-doctor-source.mjs';
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {gzipSync} from 'node:zlib';
import {publishDoctorMonth} from './process-doctor-calendars.mjs';
import {onRequestPost} from '../functions/api/automation/doctor-calendars.js';
import {doctorMonthIndexKey,loadDoctorMonthSelection} from '../functions/_lib/doctor-calendar-artifacts.js';
import {facilityMetadataManifestKey,facilityMonthObjectKey} from '../functions/_lib/facility-overview-cache.js';
import {loadPublishedDoctorCalendar} from '../functions/_lib/published-doctor-calendar.js';
const hash=v=>createHash('sha256').update(v).digest('hex');
class Bucket{
 objects=new Map();metadata=new Map();reads=[];writes=[];failRows=false;failCatalogue=false;
 async list({prefix='',limit=8,cursor=''}){const keys=[...this.objects.keys()].filter(k=>k.startsWith(prefix)&&k>cursor).sort();const selected=keys.slice(0,limit);return {objects:selected.map(key=>({key})),truncated:keys.length>limit,cursor:selected.at(-1)||''};}
 async head(key){return this.objects.has(key)?{key,etag:hash(this.objects.get(key)),customMetadata:this.metadata.get(key)||{}}:null;}
 async get(key){this.reads.push(key);const entry=this.objects.get(key);return entry?{body:entry,customMetadata:this.metadata.get(key)||{},etag:hash(entry),async arrayBuffer(){return entry.buffer.slice(entry.byteOffset,entry.byteOffset+entry.length)}}:null;}
 async put(key,value,options={}){if(this.failRows&&key.includes('/rows/')||this.failCatalogue&&key.includes('/catalogues/'))throw Error('Injected failure');if(options.onlyIf?.etagDoesNotMatch==='*'&&this.objects.has(key)||options.onlyIf?.etagMatches&&(!this.objects.has(key)||hash(this.objects.get(key))!==options.onlyIf.etagMatches))return null;const bytes=typeof value==='string'?Buffer.from(value):value instanceof ReadableStream?Buffer.from(await new Response(value).arrayBuffer()):Buffer.from(value);if(options.sha256)assert.equal(hash(bytes),Buffer.from(options.sha256).toString('hex'));this.objects.set(key,bytes);this.metadata.set(key,options.customMetadata||{});this.writes.push(key);return {key};}
}
const r2=new Bucket(),sourceType='mmc',month='2026-10',doctor='DOCTOR A';
const event=(id,doctorKey,seniority='Junior Registrar')=>({sourceType,doctorKey,seniority,event:{id,doctorKey,source:'MMC',title:'AM shift',start:'2026-10-07T08:00:00+11:00',end:'2026-10-07T16:00:00+11:00',seniority,rawValue:'AM'}});
let rows=[event('a',doctor),event('b','DOCTOR B')],revision;
// Month publications may contain a previous month's night continuing on day 1.
const overnight=event('overnight','NIGHT DOCTOR');overnight.event.start='2026-09-30T23:00:00+10:00';overnight.event.end='2026-10-01T08:00:00+10:00';rows.push(overnight);
async function generation(){const snapshot={sourceType,month,rows};revision=hash(JSON.stringify(snapshot));await r2.put(facilityMonthObjectKey(sourceType,month,revision),gzipSync(JSON.stringify(snapshot)));return revision;}
async function request(input,{binary=false,metadata=false,body}={}){
 const headers={authorization:'Bearer fixture','Content-Type':'application/json'};let payload=JSON.stringify(input);
 if(body!==undefined){payload=Buffer.from(body);headers['Content-Type']='application/octet-stream';headers['Content-Length']=String(payload.length);headers['x-doctor-artifact']=Buffer.from(JSON.stringify({...input,bytes:payload.length})).toString('base64');}
 const response=await onRequestPost({env:{ROSTER_FILES:r2,ROSTER_AUTOMATION_TOKEN:'fixture',ROSTER_AUTOMATION_WRITES_ENABLED:'true'},request:new Request('https://fixture',{method:'POST',headers,body:payload})});
 if(!response.ok)throw Error('Callback failed '+response.status);if(binary){const bytes=new Uint8Array(await response.arrayBuffer());return metadata?{bytes,etag:response.headers.get('x-object-etag')||'',key:input.key}:bytes;}return response.json();
}
await generation();let result=await publishDoctorMonth({sourceType,month,revision,request});assert.equal(result.changed,3);
let selection=await loadDoctorMonthSelection(r2,sourceType,month,revision,[doctor]);assert.deepEqual(selection.rows,[rows[0]]);const firstDigest=selection.selected[0][1];
const writeCount=r2.writes.length,readCount=r2.reads.length;
assert.equal((await publishDoctorMonth({sourceType,month,revision,request})).status,'unchanged');assert.equal(r2.writes.length,writeCount);assert.equal(r2.reads.length,readCount,'unchanged generation performs no full-month download');
let uploadCalls=0;const originalRequest=request;const countedRequest=async(i,o)=>{if(i.kind==='rows')uploadCalls++;return originalRequest(i,o);};
rows=[rows[0],event('b-changed','DOCTOR B','Senior Registrar')];await generation();result=await publishDoctorMonth({sourceType,month,revision,request:countedRequest});assert.equal(uploadCalls,1,'unchanged doctor performs no upload request or object head check');assert.equal(result.changed,1,'only changed doctor receives new content');
selection=await loadDoctorMonthSelection(r2,sourceType,month,revision,[doctor]);assert.equal(selection.selected[0][1],firstDigest);
const staffKey='facility-overview/v1/mmc/staff/2026-08-03/'+hash('staff')+'.json.gz';
async function manifest(){await r2.put(staffKey,JSON.stringify({coverage:[{startDate:'2026-08-03',endDate:'2026-11-01'}],seniorityOverrides:[]}));await r2.put(facilityMetadataManifestKey(sourceType),JSON.stringify({revision:hash(revision),terms:[{termStart:'2026-10-01',termEnd:'2026-10-31',staffKey}],months:{[month]:{key:facilityMonthObjectKey(sourceType,month,revision),revision}}}));}
await manifest();await publishDoctorSource({sourceType,request});const profile={sourceTypes:[sourceType],aliases:[{sourceType,key:doctor}],doctorKey:doctor,displayName:'Doctor A',state:{session:{}},identityAliasesEnforced:true,doctorArtifactsEnabled:true};
const options={range:{startDate:'2026-10-01',endDate:'2026-10-31'},today:'2026-10-07'};
let calendar=await loadPublishedDoctorCalendar(r2,profile,options);assert.equal(calendar.snapshot.preview.events.length,1);assert.equal(calendar.snapshot.preview.events[0].seniority,'Junior Registrar');const personalRevision=calendar.calendarRevision;
let historicalReads=0;
const historicalSnapshot={preview:{events:[
 {...rows[0].event,id:'history',start:'2026-07-07T08:00:00+10:00',end:'2026-07-07T16:00:00+10:00'},
 {...rows[0].event,id:'former-person',doctorKey:'SOMEONE ELSE',start:'2026-07-07T08:00:00+10:00',end:'2026-07-07T16:00:00+10:00'},
 {...rows[0].event,id:'removed-current-shift'},
]}};
const lazyOptions={...options,range:{startDate:'2026-07-01',endDate:'2026-10-31'},historicalRevision:'immutable-history-v1',previousSnapshotLoader:async()=>{historicalReads++;return historicalSnapshot;}};
const lazyCalendar=await loadPublishedDoctorCalendar(r2,profile,lazyOptions);
const eagerCalendar=await loadPublishedDoctorCalendar(r2,profile,{...lazyOptions,previousSnapshotLoader:null,previousSnapshot:historicalSnapshot});
const stablePreview=({lastParsed,...preview})=>preview;
assert.deepEqual(stablePreview(lazyCalendar.snapshot.preview),stablePreview(eagerCalendar.snapshot.preview),'lazy history preserves exact events and review details');
assert.equal(historicalReads,1);
assert.deepEqual(lazyCalendar.snapshot.preview.events.map(e=>e.id).sort(),['a','history'],'retains history but excludes former identities and removed current-term shifts');
const lazyCurrent=await loadPublishedDoctorCalendar(r2,profile,{...lazyOptions,cachedRevision:lazyCalendar.calendarRevision});
assert.equal(lazyCurrent.snapshotCurrent,true);assert.equal(historicalReads,1,'unchanged calendar never downloads or processes legacy history');
const historyChanged=await loadPublishedDoctorCalendar(r2,profile,{...lazyOptions,historicalRevision:'immutable-history-v2',cachedRevision:lazyCalendar.calendarRevision});
assert.notEqual(historyChanged.snapshotCurrent,true);assert.equal(historicalReads,2,'new historical artifact invalidates the browser calendar');
assert.equal((await loadPublishedDoctorCalendar(r2,profile,{...lazyOptions,previousSnapshotLoader:async()=>null})).snapshotAvailable,false,'missing pinned history defers rather than silently dropping shifts');
r2.reads=[];calendar=await loadPublishedDoctorCalendar(r2,profile,{...options,cachedRevision:personalRevision});assert.equal(calendar.snapshotCurrent,true);assert.equal(r2.reads.some(k=>k.includes('/rows/')||k.includes('/months/')||k.includes('/people/')),false,'unchanged check loads neither doctor shifts nor site roster');
const shardCurrent=JSON.parse(r2.objects.get(doctorSourceCurrentKey(sourceType))),ownShard=await doctorCatalogueShard(doctor);
assert.equal(r2.reads.includes(shardCurrent.catalogueKey),false,'calendar never loads the full site identity catalogue when small groups are available');
assert.deepEqual(r2.reads.filter(key=>key.includes('/catalogues/')),[shardCurrent.shards[ownShard].key],'single identity reads exactly its small index group');
const ownShardBytes=Buffer.from(r2.objects.get(shardCurrent.shards[ownShard].key));
r2.objects.delete(shardCurrent.shards[ownShard].key);
assert.equal((await loadPublishedDoctorCalendar(r2,profile,options)).snapshotAvailable,false,'missing index group is pending, never an empty calendar');
r2.objects.set(shardCurrent.shards[ownShard].key,ownShardBytes);
const ungrouped={...shardCurrent};delete ungrouped.layout;delete ungrouped.metadata;delete ungrouped.shards;
r2.objects.set(doctorSourceCurrentKey(sourceType),Buffer.from(JSON.stringify(ungrouped)));
assert.equal((await loadPublishedDoctorCalendar(r2,profile,{...options,cachedRevision:personalRevision})).snapshotCurrent,true,'existing complete catalogues remain readable during migration');
r2.objects.set(doctorSourceCurrentKey(sourceType),Buffer.from(JSON.stringify(shardCurrent)));
rows=[rows[0],event('b-again','DOCTOR B')];await generation();await publishDoctorMonth({sourceType,month,revision,request});await manifest();await publishDoctorSource({sourceType,request});
calendar=await loadPublishedDoctorCalendar(r2,profile,{...options,cachedRevision:personalRevision});assert.equal(calendar.snapshotCurrent,true,'colleague edit does not invalidate this doctor');
calendar=await loadPublishedDoctorCalendar(r2,profile,{...options,cachedRevision:personalRevision,range:{startDate:'2026-10-08',endDate:'2026-10-31'}});assert.notEqual(calendar.snapshotCurrent,true,'range changes invalidate');assert.equal(calendar.snapshot.preview.events.length,0);
rows=[event('b-only','DOCTOR B')];await generation();await publishDoctorMonth({sourceType,month,revision,request});await manifest();await publishDoctorSource({sourceType,request});
calendar=await loadPublishedDoctorCalendar(r2,profile,{...options,cachedRevision:personalRevision});assert.equal(calendar.snapshot.preview.events.length,0,'removed shift is removed, not retained');
rows=[event('new',doctor)];await generation();r2.failRows=true;
await assert.rejects(publishDoctorMonth({sourceType,month,revision,request}));assert.equal(await r2.head(doctorMonthIndexKey(sourceType,month,revision)),null,'failed upload cannot publish completion index');r2.failRows=false;
await manifest();calendar=await loadPublishedDoctorCalendar(r2,profile,options);assert.equal(calendar.snapshotAvailable,false,'missing complete generation cannot publish partial calendar');
// A completed person upload survives interruption before catalogue publication.
await publishDoctorMonth({sourceType,month,revision,request});
const oldCurrent=Buffer.from(r2.objects.get(doctorSourceCurrentKey(sourceType)));
r2.failCatalogue=true;
await assert.rejects(publishDoctorSource({sourceType,request}));
assert.deepEqual(r2.objects.get(doctorSourceCurrentKey(sourceType)),oldCurrent,'failed catalogue never promotes a partially built source');
r2.failCatalogue=false;r2.reads=[];r2.writes=[];
await publishDoctorSource({sourceType,request});
assert.equal(r2.reads.some(key=>key.includes('/rows/')),false,'resume reuses completed person receipts without reassembling shifts');
assert.equal(r2.writes.some(key=>key.includes('/people/')),false,'resume never uploads completed person artifacts again');
const currentKey=doctorSourceCurrentKey(sourceType),current=JSON.parse(r2.objects.get(currentKey)),operationRevision=hash('staged-operation');
const prepared={mode:'prepare-source',sourceType,operationRevision,manifestRevision:hash(revision),catalogueKey:current.catalogueKey,digest:current.digest,baseEtag:'initial-base'};
await request(prepared);await request({...prepared,baseEtag:'rebased-current'});
assert.equal(JSON.parse(r2.objects.get(doctorSourceKey(sourceType,'generations',operationRevision))).baseEtag,'rebased-current','same source generation can rebase after concurrent bootstrap');
await assert.rejects(request({...prepared,manifestRevision:hash('different-generation')}),'different source revision cannot overwrite the staged generation');
const beforeRace=Buffer.from(r2.objects.get(currentKey));
await assert.rejects(request({...prepared,mode:'activate-source',manifestEtag:'stale-parent',baseEtag:(await r2.head(currentKey)).etag}));
assert.deepEqual(r2.objects.get(currentKey),beforeRace,'changed source parent rejects activation');
await assert.rejects(request({...prepared,mode:'activate-source',manifestEtag:(await r2.head(facilityMetadataManifestKey(sourceType))).etag,baseEtag:'stale-current'}));
assert.deepEqual(r2.objects.get(currentKey),beforeRace,'concurrent doctor publication rejects stale activation');
await r2.put('doctor-calendars/v2/build-lease.json',JSON.stringify({id:'running-build',until:Date.now()+60000}));
const watchdogEnv={ROSTER_FILES:r2,ROSTER_WATCHDOG_TOKEN:'watchdog-fixture',ROSTER_AUTOMATION_WRITES_ENABLED:'true',SUBSCRIPTION_ARTIFACT_MODE:'build'};
const watchdogPost=input=>onRequestPost({env:watchdogEnv,request:new Request('https://fixture',{method:'POST',headers:{authorization:'Bearer watchdog-fixture'},body:JSON.stringify(input)})});
assert.equal((await (await watchdogPost({mode:'check-source'})).json()).status,'busy','watchdog coalesces an active build');
assert.equal((await watchdogPost({mode:'object',sourceType,key:currentKey})).status,401,'watchdog cannot read private calendar artifacts');
// Exercise the actual authenticated calendar entry with a legacy cache, rather
// than proving the helper alone skips its history callback.
const accounts=new DatabaseSync(':memory:');
for(const name of (await readdir(new URL('../migrations',import.meta.url))).filter(n=>n.endsWith('.sql')).sort())accounts.exec(await readFile(new URL('../migrations/'+name,import.meta.url),'utf8'));
const email='calendar-fixture@example.test',salt='fixture-salt',password='fixture-password';
accounts.prepare('INSERT INTO account_profiles(email,real_name,role,password_salt,password_hash) VALUES(?,?,?,?,?)').run(email,'Doctor A','user',salt,hash(salt+':'+password));
accounts.prepare('INSERT INTO account_claims(email,source_type,doctor_key,display_name) VALUES(?,?,?,?)').run(email,sourceType,doctor,'Doctor A');
accounts.prepare('INSERT INTO account_states(email,session_json) VALUES(?,?)').run(email,JSON.stringify({doctorKey:doctor}));
const legacyBucket=new Bucket(),legacyKey='legacy-fixture.json';
await legacyBucket.put(legacyKey,JSON.stringify({...historicalSnapshot,session:{doctorKey:doctor}}),{customMetadata:{revision:'legacy-v1'}});
accounts.prepare('INSERT INTO snapshot_registry(owner_type,owner_id,doctor_key,range_key,built_revision,status,artifact_key) VALUES(?,?,?,?,?,?,?)').run('user-account',email,doctor,snapshotRegistryRangeKey(lazyOptions.range),'legacy-v1','ready',legacyKey);
const accountDb={prepare(sql){let args=[];return {bind(...values){args=values;return this},async all(){return {results:accounts.prepare(sql).all(...args)}},async first(){return accounts.prepare(sql).get(...args)||null},async run(){throw Error('Calendar reads must never write D1');}}}};
const calendarPost=cachedRevision=>statePost({data:{stateAction:'loadCalendarEvents'},env:{ROSTER_DB:accountDb,ROSTER_FILES:r2,ROSTER_CACHE:legacyBucket,IDENTITY_REVIEW_ENABLED:'true',BOUNDED_MANUAL_ROSTER_ENABLED:'true',SUBSCRIPTION_ARTIFACT_MODE:'ready'},request:new Request('https://fixture/api/state',{method:'POST',body:JSON.stringify({action:'loadCalendarEvents',email,password,doctorKey:doctor,startDate:lazyOptions.range.startDate,endDate:lazyOptions.range.endDate,cachedRevision,responseMode:'fast',allowInlineBuild:false})})});
const firstResponse=await calendarPost('');assert.equal(firstResponse.status,200);const firstPayload=await firstResponse.json();assert.equal(firstPayload.snapshotAvailable,true);
assert.deepEqual(firstPayload.snapshot.preview.events.map(e=>e.id).sort(),['history','new'],'authenticated route preserves historical shifts and the latest published replacement');
assert.equal(legacyBucket.reads.filter(k=>k===legacyKey).length,1);
legacyBucket.reads=[];
const currentResponse=await calendarPost(firstPayload.calendarRevision);assert.equal(currentResponse.status,200);assert.equal((await currentResponse.json()).snapshotCurrent,true);
assert.equal(legacyBucket.reads.length,0,'authenticated unchanged browser check never downloads legacy snapshot');
const originalManifest=Buffer.from(r2.objects.get(facilityMetadataManifestKey(sourceType)));
await r2.put(facilityMetadataManifestKey(sourceType),JSON.stringify({...JSON.parse(originalManifest),revision:hash('new-parent')}));
const pendingPayload=await (await calendarPost(firstPayload.calendarRevision)).json();
assert.equal(pendingPayload.snapshotAvailable,true,'pending publication retains a usable legacy calendar');
assert.equal(legacyBucket.reads.filter(k=>k===legacyKey).length,1,'pending publication loads legacy fallback only when required');
assert.equal(JSON.stringify(pendingPayload).includes('historicalRevision'),false,'internal cache descriptors are never included in responses');
r2.objects.set(facilityMetadataManifestKey(sourceType),originalManifest);
accounts.prepare("UPDATE account_profiles SET role='creator' WHERE email=?").run(email);
accounts.prepare('DELETE FROM account_claims WHERE email=?').run(email);
accounts.prepare('INSERT INTO canonical_doctors(canonical_key,display_name,aliases_json) VALUES(?,?,?)').run(doctor,'Doctor A',JSON.stringify([{sourceType,key:doctor,displayName:'Doctor A'}]));
accounts.prepare('INSERT INTO snapshot_registry(owner_type,owner_id,doctor_key,range_key,built_revision,status,artifact_key) VALUES(?,?,?,?,?,?,?)').run('creator-account',email,doctor,snapshotRegistryRangeKey(lazyOptions.range),'legacy-v1','ready',legacyKey);
const creatorInitial=await (await calendarPost('')).json();assert.equal(creatorInitial.snapshotAvailable,true);
assert.deepEqual(creatorInitial.snapshot.doctorOptions[0].sourceTypes,[sourceType],'indexed Creator selection preserves exact alias sites');
assert.deepEqual(creatorInitial.snapshot.preview.events,firstPayload.snapshot.preview.events);
legacyBucket.reads=[];
assert.equal((await (await calendarPost(creatorInitial.calendarRevision)).json()).snapshotCurrent,true);
assert.equal(legacyBucket.reads.length,0,'Creator without claims uses indexed selection instead of full legacy cache');
legacyBucket.metadata.set(legacyKey,{revision:'replacement-in-progress'});
assert.equal((await (await calendarPost('')).json()).snapshotAvailable,false,'cache overwritten before registry update cannot publish under the previous revision');
legacyBucket.metadata.set(legacyKey,{revision:'legacy-v1'});
assert.equal((await onRequestPost({env:{ROSTER_FILES:r2,ROSTER_AUTOMATION_TOKEN:'fixture',ROSTER_AUTOMATION_WRITES_ENABLED:'true'},request:new Request('https://fixture',{method:'POST',body:'{}'})})).status,401);
console.log('Doctor calendars pass external assembly, unchanged zero-download/zero-write, colleague isolation, metadata-only checks, grades, ranges, removals and publication failure.');

const historyKey='snapshots/user-account/fixture@test/DOCTOR/2026.json.gz',historyRevision='history-v1:'+ 'x'.repeat(5000);
const historySummary=summarizeLegacyCoverage(historicalSnapshot,historyKey,historyRevision);
const currentTerms=[{sourceType,termStart:'2026-10-01',termEnd:'2026-10-31'}];
assert.equal(legacyHistoryIsCovered(historySummary,historyKey,historyRevision,options.range,currentTerms),true,'history outside requested dates requires no decoding');
assert.equal(legacyHistoryIsCovered(historySummary,historyKey,historyRevision,lazyOptions.range,currentTerms),false,'published October must not discard retained July shifts');
assert.equal(legacyHistoryIsCovered(historySummary,historyKey,'different',options.range,currentTerms),false,'another historical generation cannot authorize skipping');
const allTerms=[{sourceType,termStart:'2026-07-01',termEnd:'2026-08-31'},{sourceType,termStart:'2026-09-01',termEnd:'2026-10-31'}];
assert.equal(legacyHistoryIsCovered(historySummary,historyKey,historyRevision,lazyOptions.range,allTerms),true,'adjacent terms cover the whole historical interval');
assert.equal(legacyHistoryIsCovered(historySummary,historyKey,historyRevision,lazyOptions.range,[allTerms[0],{...allTerms[1],termStart:'2026-09-02'}]),false,'a one-day gap forces the full historical read');
const historyStore=new Bucket();await historyStore.put(historyKey,gzipSync(JSON.stringify(historicalSnapshot)),{customMetadata:{revision:historyRevision}});
const legacyRequest=async(input,{binary=false}={})=>{
 const response=await onRequestPost({env:{ROSTER_FILES:r2,ROSTER_CACHE:historyStore,ROSTER_AUTOMATION_TOKEN:'fixture',ROSTER_AUTOMATION_WRITES_ENABLED:'true'},request:new Request('https://fixture',{method:'POST',headers:{authorization:'Bearer fixture'},body:JSON.stringify(input)})});
 if(!response.ok)throw Error('Legacy coverage callback '+response.status);
 return binary?new Uint8Array(await response.arrayBuffer()):response.json();
};
assert.equal((await prepareLegacyCalendarCoverage({request:legacyRequest})).prepared,1);
const storedSummary=JSON.parse(historyStore.objects.get(await legacyCoverageKey(historyKey,historyRevision)));const expectedSummary={...historySummary,scope:await legacyCoverageKey(historyKey,historyRevision)};delete expectedSummary.revision;assert.deepEqual(storedSummary,expectedSummary);
assert.equal(legacyHistoryIsCovered(storedSummary,historyKey,historyRevision,options.range,currentTerms,expectedSummary.scope),true);
historyStore.reads=[];assert.equal((await prepareLegacyCalendarCoverage({request:legacyRequest})).prepared,0);
assert.equal(historyStore.reads.includes(historyKey),false,'completed backfill never downloads an unchanged legacy snapshot');
await assert.rejects(legacyRequest({mode:'legacy-publish',key:historyKey,revision:historyRevision,etag:'stale',summary:historySummary}),'stale raw history cannot publish coverage');
// Proven coverage is used by the actual account route, not only the predicate.
await legacyBucket.put(await legacyCoverageKey(legacyKey,'legacy-v1'),JSON.stringify(summarizeLegacyCoverage(historicalSnapshot,legacyKey,'legacy-v1')));
legacyBucket.reads=[];
const octoberOnly={...lazyOptions.range,startDate:'2026-10-01'};
accounts.prepare('INSERT INTO snapshot_registry(owner_type,owner_id,doctor_key,range_key,built_revision,status,artifact_key) VALUES(?,?,?,?,?,?,?)').run('creator-account',email,doctor,snapshotRegistryRangeKey(octoberOnly),'legacy-v1','ready',legacyKey);
const coveredResponse=await statePost({data:{stateAction:'loadCalendarEvents'},env:{ROSTER_DB:accountDb,ROSTER_FILES:r2,ROSTER_CACHE:legacyBucket,IDENTITY_REVIEW_ENABLED:'true',BOUNDED_MANUAL_ROSTER_ENABLED:'true',SUBSCRIPTION_ARTIFACT_MODE:'ready'},request:new Request('https://fixture/api/state',{method:'POST',body:JSON.stringify({action:'loadCalendarEvents',email,password,doctorKey:doctor,startDate:octoberOnly.startDate,endDate:octoberOnly.endDate,responseMode:'fast',allowInlineBuild:false})})});
assert.equal((await coveredResponse.json()).snapshot.preview.events[0].id,'new');
assert.equal(legacyBucket.reads.includes(legacyKey),false,'account route skips compressed historical body only with complete coverage proof');

// The normal annual browser window includes the following January and keeps
// historical shifts from an existing cache keyed to the former 370-day bound.
const annualMonths={...JSON.parse(r2.objects.get(facilityMetadataManifestKey(sourceType))).months};
for(const annualMonth of ['2026-11','2026-12','2027-01']){
 const annualRows=annualMonth==='2027-01'?[{sourceType,doctorKey:doctor,event:{...rows[0].event,id:'late-january',start:'2027-01-25T08:00:00+11:00',end:'2027-01-25T16:00:00+11:00'}}]:[];
 const annualRevision=hash(JSON.stringify(annualRows)),key=facilityMonthObjectKey(sourceType,annualMonth,annualRevision);
 await r2.put(key,JSON.stringify({version:1,sourceType,month:annualMonth,revision:annualRevision,rows:annualRows}));
 await publishDoctorMonth({sourceType,month:annualMonth,revision:annualRevision,request});annualMonths[annualMonth]={key,revision:annualRevision};
}
const annualStaffKey='facility-overview/v1/mmc/staff/2026-11-02/'+hash('annual-staff')+'.json.gz';
await r2.put(annualStaffKey,JSON.stringify({coverage:[{startDate:'2026-11-02',endDate:'2027-01-31'}],seniorityOverrides:[]}));
const annualManifest=JSON.parse(r2.objects.get(facilityMetadataManifestKey(sourceType)));
await r2.put(facilityMetadataManifestKey(sourceType),JSON.stringify({...annualManifest,revision:hash('annual'),months:annualMonths,terms:[...annualManifest.terms,{termStart:'2026-11-02',termEnd:'2027-01-31',staffKey:annualStaffKey}]}));
await publishDoctorSource({sourceType,request});
const oldAnnual={startDate:'2026-01-01',endDate:'2027-01-06'},annual={...oldAnnual,endDate:'2027-01-31'};
accounts.prepare('INSERT INTO snapshot_registry(owner_type,owner_id,doctor_key,range_key,built_revision,status,artifact_key) VALUES(?,?,?,?,?,?,?)').run('creator-account',email,doctor,snapshotRegistryRangeKey(oldAnnual),'legacy-v1','ready',legacyKey);
const annualPost=()=>statePost({data:{stateAction:'loadCalendarEvents'},env:{ROSTER_DB:accountDb,ROSTER_FILES:r2,ROSTER_CACHE:legacyBucket,IDENTITY_REVIEW_ENABLED:'true',BOUNDED_MANUAL_ROSTER_ENABLED:'true',SUBSCRIPTION_ARTIFACT_MODE:'ready'},request:new Request('https://fixture/api/state',{method:'POST',body:JSON.stringify({action:'loadCalendarEvents',email,password,doctorKey:doctor,...annual,responseMode:'fast',allowInlineBuild:false})})});
assert.deepEqual((await (await annualPost()).json()).snapshot.preview.events.map(e=>e.id).sort(),['history','late-january','new'],'annual window delivers late January without discarding the old historical cache');
accounts.prepare('INSERT INTO snapshot_registry(owner_type,owner_id,doctor_key,range_key,built_revision,status,artifact_key) VALUES(?,?,?,?,?,?,?)').run('creator-account',email,doctor,snapshotRegistryRangeKey(annual),'legacy-v1','ready',legacyKey);
const exactEntry=await loadSnapshotRegistryEntry(accountDb,{ownerType:'creator-account',ownerId:email,doctorKey:doctor,rangeKey:snapshotRegistryRangeKey(annual)},{fallbackRangeKey:snapshotRegistryRangeKey(oldAnnual)});
assert.equal(exactEntry.rangeKey,snapshotRegistryRangeKey(annual),'an exact annual cache takes priority over migration fallback');

for(const compressed of [false,true]){
 const value={events:[{id:'stream-fixture',title:'Résumé'}]},bytes=compressed?gzipSync(JSON.stringify(value)):Buffer.from(JSON.stringify(value));
 const streamBucket={async get(){return {customMetadata:{revision:'stream-v1'},httpMetadata:{contentType:'application/json',...(compressed?{contentEncoding:'gzip'}:{})},body:new ReadableStream({start(controller){controller.enqueue(bytes);controller.close()}}),async arrayBuffer(){return bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength);}}}};
 assert.deepEqual(await loadCachedSnapshot(streamBucket,'stream.json',{revision:'stream-v1'}),value,'JSON/gzip retains exact Unicode content');
 assert.equal(await loadCachedSnapshot(streamBucket,'stream.json',{revision:'stream-v2'}),null,'cache checks the expected revision before decoding its body');
}

// Interval gaps must exactly reproduce the old daily coverage semantics.
for(const terms of [[],[{termStart:'2026-08-03',termEnd:'2026-11-01',empty:true,coverage:[]}],[{termStart:'2026-08-03',termEnd:'2026-11-01',empty:false,coverage:[{startDate:'2026-08-05',endDate:'2026-08-10'},{startDate:'2026-09-01',endDate:'2026-10-30'}]}]]){
 const expected=[];let gap=null;
 for(let date='2026-08-01';date<='2026-11-03';){const term=terms.find(t=>t.termStart<=date&&t.termEnd>=date),covered=term&&(term.empty||!Array.isArray(term.coverage)||term.coverage.some(c=>c.startDate<=date&&c.endDate>=date));const reason=!term?'history-unavailable':!covered?'coverage-unavailable':'';if(reason&&gap?.reason===reason)gap.endDate=date;else{if(gap)expected.push(gap);gap=reason?{sourceType:'mmc',startDate:date,endDate:date,reason}:null;}const next=new Date(date+'T12:00:00Z');next.setUTCDate(next.getUTCDate()+1);date=next.toISOString().slice(0,10);}if(gap)expected.push(gap);
 assert.deepEqual(doctorRangeGaps('mmc',terms,'2026-08-01','2026-11-03'),expected);
}

const presentation=new DatabaseSync(':memory:');presentation.exec(`CREATE TABLE account_hospital_locations(email TEXT,source_type TEXT,location TEXT);CREATE TABLE custom_events(owner_email TEXT,id TEXT,title TEXT,start_date TEXT,end_date TEXT,all_day INTEGER,start_time TEXT,end_time TEXT,location TEXT,include INTEGER);INSERT INTO account_hospital_locations VALUES('a@test','mmc','Personal location');INSERT INTO custom_events VALUES('a@test','custom','Annual Leave','2026-10-01','2026-10-02',1,'','','',1);`);
let presentationQueries=0;const pDb={prepare(query){presentationQueries++;let args=[];return {bind(...values){args=values;return this},async first(){return presentation.prepare(query).get(...args)},async all(){return {results:presentation.prepare(query).all(...args)}}}}};
const inputs=await loadCalendarPresentationInputs(pDb,'a@test',{});assert.equal(presentationQueries,1);assert.equal(inputs.locations.mmc,'Personal location');assert.equal(inputs.customEvents[0].ownerEmail,'a@test');assert.equal(inputs.customEvents[0].allDay,true);
for(let i=0;i<201;i++)presentation.prepare('INSERT INTO custom_events VALUES(?,?,?,?,?,1,?,?,?,?)').run('a@test','extra'+i,'Leave','2026-10-01','2026-10-02','','','',1);
const oversized=await loadCalendarPresentationInputs(pDb,'a@test',{});assert.equal(oversized.customEvents.length,202,'existing large custom calendar is never truncated');
