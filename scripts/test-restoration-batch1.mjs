import assert from 'node:assert/strict';
import { DatabaseSync } from 'node:sqlite';
import { readFile,readdir } from 'node:fs/promises';
import { rosterTermAvailableFrom, melbourneDateKey } from '../public/static/roster-term-policy.js';
import { createCalendarRevisionPoller } from '../public/static/calendar-revision-poller.js';
import { sharepointRosterWindows,onRequestPost as checkMetadata } from '../functions/api/automation/roster-check.js';
import { onRequestGet as revision } from '../functions/api/roster-revision.js';
import { loadPublishedFacilityRange,facilityMetadataManifestKey } from '../functions/_lib/facility-overview-cache.js';
import { queryFacilityStaffSeniorityOverrides } from '../functions/_lib/d1-calendar.js';
assert.equal(rosterTermAvailableFrom('2026-11-02'),'2026-10-01');
assert.equal(rosterTermAvailableFrom('2027-01-04'),'2026-12-01');
assert.equal(melbourneDateKey(new Date('2026-09-30T14:00:00Z')),'2026-10-01');
assert.equal(sharepointRosterWindows(new Date('2026-09-30T13:59:00Z')).length,3);
assert.equal(sharepointRosterWindows(new Date('2026-09-30T14:00:00Z')).length,5);
assert.equal(sharepointRosterWindows(new Date('2026-10-04T01:00:00Z')).find(w=>w.sourceId==='monash-paeds'&&w.termStart==='2026-11-02').fileName,'Paeds - Term 4 2026.xlsx');

let view={key:'account-A',sources:['mmc']},allowed=true,reads=0,refreshes=0,fail=false,version='one';
const timers=new Map();let nextTimer=0;
const poller=createCalendarRevisionPoller({context:()=>view,eligible:()=>allowed,read:async()=>{reads++;return version;},refresh:async()=>{refreshes++;return !fail;},schedule:f=>{timers.set(++nextTimer,f);return nextTimer;},cancel:id=>timers.delete(id)});
await poller.tick(); await poller.tick();
assert.equal(reads,2);assert.equal(refreshes,1,'unchanged ticks must not load account calendars');
allowed=false; await poller.tick(); assert.equal(reads,2,'pending edits block polling');
allowed=true; version='two';fail=true;await poller.tick();fail=false;await poller.tick();assert.equal(refreshes,3,'failed refresh retries unchanged fingerprint');
view=null;poller.update();assert.equal(timers.size,0,'hidden view stops timers');
view={key:'account-B',sources:['mmc']};await poller.tick();assert.equal(refreshes,4,'account switch has independent baseline');
let release;
const switching=createCalendarRevisionPoller({context:()=>view,eligible:()=>true,read:()=>new Promise(r=>release=r),refresh:async()=>{throw Error('must not apply prior account');},schedule:()=>0,cancel:()=>{}});
const inFlight=switching.tick();view={key:'account-C',sources:['mmc']};release('three');await inFlight;

const sqlite=new DatabaseSync(':memory:');
for(const name of (await readdir(new URL('../migrations',import.meta.url))).filter(n=>n.endsWith('.sql')).sort())sqlite.exec(await readFile(new URL(`../migrations/${name}`,import.meta.url),'utf8'));
for(const sql of ["ALTER TABLE raw_roster_files ADD COLUMN name TEXT NOT NULL DEFAULT ''","ALTER TABLE raw_roster_files ADD COLUMN source_type TEXT NOT NULL DEFAULT ''","ALTER TABLE raw_roster_files ADD COLUMN size INTEGER NOT NULL DEFAULT 0","ALTER TABLE raw_roster_files ADD COLUMN last_modified INTEGER NOT NULL DEFAULT 0"])sqlite.exec(sql);
let writes=0;const sqlLog=[];
const db = { prepare(sql) {
 sqlLog.push(sql);
 return {
  args: [], bind(...args) { this.args = args; return this; },
  async first() { return sqlite.prepare(sql).get(...this.args) || null; },
  async all() { return { results: sqlite.prepare(sql).all(...this.args) }; },
  async run() { const result = sqlite.prepare(sql).run(...this.args); writes += Number(result.changes); return { meta: { changes: Number(result.changes) } }; },
 };
}};
sqlite.exec(`INSERT INTO raw_roster_files(file_id,name,object_key) VALUES('raw','AdultTerm3.2026.xlsx','retained'); INSERT INTO roster_sync_runs(id,source_id,provider_version,file_id,source_file_id,status,started_at) VALUES('run','monash-adults','1.0','raw','raw','success','2026-10-01');`);
const env={ROSTER_DB:db,ROSTER_AUTOMATION_TOKEN:'fixture',ROSTER_METADATA_CHECK_ENABLED:'true',ROSTER_AUTOMATION_WRITES_ENABLED:'true',ROSTER_AUTOMATION_SOURCE_ALLOWLIST:'monash-adults'};
const call=body=>checkMetadata({env,request:new Request('https://example.com/api/automation/roster-check',{method:'POST',headers:{authorization:'Bearer fixture','content-type':'application/json'},body:JSON.stringify(body)})});
const metadata={sourceId:'monash-adults',fileName:'AdultTerm3.2026.xlsx',providerVersion:'1.0'};
assert.equal((await (await call(metadata)).json()).download,false);assert.equal(writes,0);
assert.equal(sqlLog.some(sql=>/roster_events|CREATE |UPDATE |INSERT /i.test(sql)),false,'unchanged metadata must only read indexed compact state');
assert.equal((await (await call({...metadata,providerVersion:'1.1'})).json()).download,true,'different version at same Modified time must download');
sqlite.exec("UPDATE roster_sync_runs SET status='failed' WHERE id='run'");assert.equal((await (await call(metadata)).json()).status,'repair-required');assert.equal(writes,0,'invalid unchanged candidates cannot cause retry storms');
const before=sqlLog.length;assert.equal((await call({...metadata,sourceId:'casey-manual'})).status,403);assert.equal(sqlLog.length,before);

sqlite.exec(`INSERT INTO facility_staff_seniority_overrides(id,source_type,doctor_key,display_name,seniority,term_start,active,created_at,updated_at) VALUES('old','mmc','ALIAS','Example','Junior Registrar','2026-05-04',1,'now','now');`);
assert.equal((await queryFacilityStaffSeniorityOverrides(db,{sourceType:'mmc',termStart:'2026-05-04'})).length,1);
assert.equal((await queryFacilityStaffSeniorityOverrides(db,{sourceType:'mmc',termStart:'2026-08-03'})).length,0,'old correction never follows promotion');
const objects=new Map();let r2Reads=0;
const r2={async get(key){r2Reads++;const value=objects.get(key);return value?{async arrayBuffer(){return new TextEncoder().encode(JSON.stringify(value)).buffer}}:null}};
const terms=[['2026-05-04','2026-08-02','Junior Registrar'],['2026-08-03','2026-11-01','Senior Registrar'],['2026-11-02','2027-01-31','SMS']];
const manifest={revision:'published-one',terms:terms.map(([termStart,termEnd],i)=>({termStart,termEnd,visibleFrom:'2099-01-01',staffKey:`staff-${i}`,staffRevision:`s${i}`})),months:{}};
objects.set(facilityMetadataManifestKey('mmc'),manifest);
for(const [i,[start,end,grade]] of terms.entries()){
 objects.set(`staff-${i}`,{seniorityOverrides:[{sourceType:'mmc',doctorKey:'ALIAS',termStart:'2026-05-04',seniority:'Junior Registrar'}]});
 const month=start.slice(0,7);manifest.months[month]={key:month,revision:month};objects.set(month,{rows:[{sourceType:'mmc',doctorKey:'ALIAS',seniority:grade,event:{id:month,start:`${start}T08:00:00`,end:`${start}T17:00:00`,seniority:grade}}]});
}
const history=await loadPublishedFacilityRange(r2,['mmc'],'2026-05-04','2026-11-02','2026-10-01');
assert.deepEqual(history.events.map(row=>row.seniority),['Junior Registrar','Senior Registrar','SMS'],'alias history keeps each roster term grade, next term visible from preceding month');
const request=new Request('https://example.com/api/roster-revision?sites=mmc');
const revisionEnv={ROSTER_VISIBLE_REFRESH_ENABLED:'true',ROSTER_FILES:r2,ROSTER_DB:{prepare(){throw Error('Revision must never query D1')}}};
const first=await (await revision({request,env:revisionEnv})).json();assert.deepEqual(Object.keys(first).sort(),['ok','revision']);assert.match(first.revision,/^[a-f0-9]{64}$/);
const same=await (await revision({request,env:revisionEnv})).json();assert.equal(same.revision,first.revision);
manifest.revision='published-two';const changed=await (await revision({request,env:revisionEnv})).json();assert.notEqual(changed.revision,first.revision);
assert.equal((await revision({request:new Request('https://example.com/api/roster-revision?sites=unknown'),env:revisionEnv})).status,400);
console.log('Batch 1 passed Melbourne eligibility, metadata no-op/version/failure isolation, historical grades, zero-D1 fingerprints, unchanged/hidden/editing/switched-account refresh.');
