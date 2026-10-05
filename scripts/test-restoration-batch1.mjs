import assert from 'node:assert/strict';
import { DatabaseSync } from 'node:sqlite';
import { runInNewContext } from 'node:vm';
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
const call=body=>checkMetadata({env,request:new Request('https://example.com/api/automation/roster-check',{method:'POST',headers:{authorization:'Bearer fixture','content-type':'application/json'},body:JSON.stringify(body)})},new Date('2026-10-04T01:00:00Z'));
const metadata={sourceId:'monash-adults',fileName:'AdultTerm3.2026.xlsx',providerVersion:'1.0'};
assert.equal((await (await call(metadata)).json()).download,false);assert.equal(writes,0);
assert.equal(sqlLog.some(sql=>/roster_events|CREATE |UPDATE |INSERT /i.test(sql)),false,'unchanged metadata must only read indexed compact state');
assert.equal((await (await call({...metadata,providerVersion:'1.1'})).json()).download,true,'different version at same Modified time must download');
sqlite.exec("UPDATE roster_sync_runs SET status='failed' WHERE id='run'");assert.equal((await (await call(metadata)).json()).status,'repair-required');assert.equal(writes,0,'invalid unchanged candidates cannot cause retry storms');
const before=sqlLog.length;assert.equal((await call({...metadata,sourceId:'casey-manual'})).status,403);assert.equal(sqlLog.length,before);

sqlite.exec("UPDATE roster_sync_runs SET status='success' WHERE id='run'");
const libraries=[{site:'monash',files:[{FileRef:'/Shared Documents/Medical Roster/AdultTerm3.2026.xlsx',Modified:'2026-10-04T01:00:00Z',OData__UIVersionString:'1.0',File:{ETag:'current-etag'}}]},{site:'vhh',files:[]}];
const bulk=await (await call({mode:'reconcile',libraries})).json();
assert.deepEqual(bulk.downloads,[],'unchanged library reconciliation does not download');
assert.equal(bulk.checks.find(check=>check.fileName==='AdultTerm4.2026.xlsx').status,'waiting-for-file','newly eligible file is checked independently');
assert.equal(writes,0);
const newLibraries=structuredClone(libraries);newLibraries[0].files.push({FileRef:'/Shared Documents/Medical Roster/AdultTerm4.2026.xlsx',Modified:'2026-10-04T01:00:00Z',OData__UIVersionString:'1.0',File:{ETag:'next-etag'}});
const delivery=await (await call({mode:'reconcile',libraries:newLibraries})).json();
assert.equal(delivery.downloads.length,1);assert.equal(delivery.downloads[0].termStart,'2026-11-02','same provider version in a different eligible file still imports');
const incomplete=await (await call({mode:'reconcile',libraries:[{...libraries[0],nextLink:'incomplete'},libraries[1]]})).json();
assert.deepEqual(incomplete.downloads,[]);assert.equal(incomplete.checks[0].status,'incomplete-inventory','truncated provider inventory fails closed for its site');
const outage=await (await call({mode:'reconcile',libraries:[newLibraries[0],{...libraries[1],unavailable:true}]})).json();
assert.equal(outage.downloads.length,1,'one unavailable library does not suppress another site’s changed next-term file');
assert.equal((await call({mode:'reconcile',libraries:[libraries[0],libraries[0]]})).status,400,'duplicate library cannot substitute for a missing site');
const misplaced=structuredClone(libraries);misplaced[0].files[0].FileRef='/Shared Documents/Elsewhere/AdultTerm3.2026.xlsx';
assert.deepEqual((await (await call({mode:'reconcile',libraries:misplaced})).json()).downloads,[],'file name alone cannot select a workbook from the wrong folder');

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
sqlite.exec("INSERT INTO roster_account_budget(utc_day,stop_reason) VALUES(date('now'),'cost-overrun:request')");
assert.equal((await (await checkMetadata({env:{...env,ROSTER_ACCOUNT_BUDGET_ENABLED:'true'},request:new Request('https://example.com/api/automation/roster-check',{method:'POST',headers:{authorization:'Bearer fixture'},body:JSON.stringify({...metadata,providerVersion:'new'})})},new Date('2026-10-04T01:00:00Z'))).json()).download,false,'budget stop prevents provider download');
console.log('Batch 1 passed Melbourne eligibility, metadata no-op/version/failure isolation, historical grades, zero-D1 fingerprints, unchanged/hidden/editing/switched-account refresh.');

// Exercise the actual browser apply path against responses arriving after edits
// and account switches, rather than only testing the scheduling abstraction.
const app = await readFile(new URL('../public/static/app.js',import.meta.url),'utf8');
const start = app.indexOf('async function refreshVisibleRosterView(');
const code = app.slice(start,app.indexOf('async function loadCloudCalendarEvents(',start));
let responseResolver, rendered = 0, session = { settings: { dateFrom:'2026-05-01',dateTo:'2026-12-01',hospitalFilter:'MMC' },customEvents:[] };
const globals = {
 calendarRevisionRefreshAllowed:()=>true,calendarRevisionView:()=>({key:globals.key}),key:'profile-A',
 activeDoctorProfile:{id:'A',doctorKey:'EXAMPLE',sourceTypes:['mmc']},
 currentUserEmail:'fixture',currentUserPassword:'fixture',authUserEmail:'fixture',authUserPassword:'fixture',
 buildActiveSessionState:()=>session,fetch:()=>new Promise(resolve=>{responseResolver=resolve}),
 readJsonResponse:async response=>response,sanitizeWorkspaceSnapshot:s=>structuredClone(s),
 applyLoadedCalendarFileRefs:()=>{},renderWorkspaceFromSnapshot:(_snapshot,_session,options)=>{assert.equal(options.preserveScroll,true);rendered++;},
};
runInNewContext(code+';this.refresh=refreshVisibleRosterView',globals);
const payload={snapshot:{session:{settings:{dateFrom:'wrong'},customEvents:[]},preview:{events:[]}},calendarRevision:'fresh'};
let applying=globals.refresh({key:'profile-A'});
session.customEvents.push({id:'unsaved'});responseResolver(payload);assert.equal(await applying,false);assert.equal(rendered,0,'in-flight edit prevents applying old server state');
applying=globals.refresh({key:'profile-A'});globals.key='profile-B';responseResolver(payload);assert.equal(await applying,false);assert.equal(rendered,0,'account switch discards old response');
globals.key='profile-A';applying=globals.refresh({key:'profile-A'});responseResolver({...payload,snapshotStale:true});assert.equal(await applying,false);
applying=globals.refresh({key:'profile-A'});responseResolver(payload);assert.equal(await applying,true);assert.equal(rendered,1);assert.equal(globals.currentSnapshot.session.settings.dateFrom,'2026-05-01');
console.log('Actual browser refresh passed in-flight edit/account guards, stale-response rejection, filter and scroll preservation.');

const saveStart = app.indexOf('async function saveCloudState(snapshot = null)');
const saveCode = app.slice(saveStart,app.indexOf('function savePayloadMatchesActiveCalendar',saveStart));
const saves = { cloudStateSaveActive:0,cloudStateSaveQueue:Promise.resolve(),unsavedCalendarContexts:new Set(),
 sessionSaveContext:payload=>payload.context,snapshotCloudSavePayload:()=>({context:'account-A'}),
 saveCloudStateNow:async()=>{throw Error('offline');} };
runInNewContext(saveCode+';this.save=saveCloudState',saves);
await assert.rejects(saves.save());assert.equal(saves.unsavedCalendarContexts.has('account-A'),true,'failed save retains protection after pending timer clears');assert.equal(saves.cloudStateSaveActive,0);
saves.saveCloudStateNow=async()=>{};await saves.save();assert.equal(saves.unsavedCalendarContexts.has('account-A'),false,'successful save releases protection');
console.log('Failed saves remain protected from background refresh until saved successfully.');
