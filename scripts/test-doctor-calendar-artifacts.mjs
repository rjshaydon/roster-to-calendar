import {DatabaseSync} from 'node:sqlite';
import {loadCalendarPresentationInputs} from '../functions/_lib/d1-calendar.js';
import {doctorRangeGaps} from '../functions/_lib/doctor-source-artifacts.js';
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
 objects=new Map();reads=[];writes=[];failRows=false;
 async head(key){return this.objects.has(key)?{key,etag:hash(this.objects.get(key))}:null;}
 async get(key){this.reads.push(key);const entry=this.objects.get(key);return entry?{body:entry,etag:hash(entry),async arrayBuffer(){return entry.buffer.slice(entry.byteOffset,entry.byteOffset+entry.length)}}:null;}
 async put(key,value,options={}){if(this.failRows&&key.includes('/rows/'))throw Error('Injected failure');if(options.onlyIf?.etagDoesNotMatch==='*'&&this.objects.has(key)||options.onlyIf?.etagMatches&&(!this.objects.has(key)||hash(this.objects.get(key))!==options.onlyIf.etagMatches))return null;const bytes=typeof value==='string'?Buffer.from(value):value instanceof ReadableStream?Buffer.from(await new Response(value).arrayBuffer()):Buffer.from(value);if(options.sha256)assert.equal(hash(bytes),Buffer.from(options.sha256).toString('hex'));this.objects.set(key,bytes);this.writes.push(key);return {key};}
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
async function manifest(){await r2.put(staffKey,JSON.stringify({coverage:[{startDate:'2026-08-03',endDate:'2026-11-01'}],seniorityOverrides:[]}));await r2.put(facilityMetadataManifestKey(sourceType),JSON.stringify({revision:hash(revision),terms:[{termStart:'2026-08-03',termEnd:'2026-11-01',staffKey}],months:{[month]:{key:facilityMonthObjectKey(sourceType,month,revision),revision}}}));}
await manifest();await publishDoctorSource({sourceType,request});const profile={sourceTypes:[sourceType],aliases:[{sourceType,key:doctor}],doctorKey:doctor,displayName:'Doctor A',state:{session:{}},identityAliasesEnforced:true,doctorArtifactsEnabled:true};
const options={range:{startDate:'2026-10-01',endDate:'2026-10-31'},today:'2026-10-07'};
let calendar=await loadPublishedDoctorCalendar(r2,profile,options);assert.equal(calendar.snapshot.preview.events.length,1);assert.equal(calendar.snapshot.preview.events[0].seniority,'Junior Registrar');const personalRevision=calendar.calendarRevision;
r2.reads=[];calendar=await loadPublishedDoctorCalendar(r2,profile,{...options,cachedRevision:personalRevision});assert.equal(calendar.snapshotCurrent,true);assert.equal(r2.reads.some(k=>k.includes('/rows/')||k.includes('/months/')||k.includes('/people/')),false,'unchanged check loads neither doctor shifts nor site roster');
rows=[rows[0],event('b-again','DOCTOR B')];await generation();await publishDoctorMonth({sourceType,month,revision,request});await manifest();await publishDoctorSource({sourceType,request});
calendar=await loadPublishedDoctorCalendar(r2,profile,{...options,cachedRevision:personalRevision});assert.equal(calendar.snapshotCurrent,true,'colleague edit does not invalidate this doctor');
calendar=await loadPublishedDoctorCalendar(r2,profile,{...options,cachedRevision:personalRevision,range:{startDate:'2026-10-08',endDate:'2026-10-31'}});assert.notEqual(calendar.snapshotCurrent,true,'range changes invalidate');assert.equal(calendar.snapshot.preview.events.length,0);
rows=[event('b-only','DOCTOR B')];await generation();await publishDoctorMonth({sourceType,month,revision,request});await manifest();await publishDoctorSource({sourceType,request});
calendar=await loadPublishedDoctorCalendar(r2,profile,{...options,cachedRevision:personalRevision});assert.equal(calendar.snapshot.preview.events.length,0,'removed shift is removed, not retained');
rows=[event('new',doctor)];await generation();r2.failRows=true;
await assert.rejects(publishDoctorMonth({sourceType,month,revision,request}));assert.equal(await r2.head(doctorMonthIndexKey(sourceType,month,revision)),null,'failed upload cannot publish completion index');r2.failRows=false;
await manifest();calendar=await loadPublishedDoctorCalendar(r2,profile,options);assert.equal(calendar.snapshotAvailable,false,'missing complete generation cannot publish partial calendar');
assert.equal((await onRequestPost({env:{ROSTER_FILES:r2,ROSTER_AUTOMATION_TOKEN:'fixture',ROSTER_AUTOMATION_WRITES_ENABLED:'true'},request:new Request('https://fixture',{method:'POST',body:'{}'})})).status,401);
console.log('Doctor calendars pass external assembly, unchanged zero-download/zero-write, colleague isolation, metadata-only checks, grades, ranges, removals and publication failure.');

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
