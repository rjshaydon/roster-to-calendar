import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {runInNewContext} from 'node:vm';
import {createCalendarRevisionPoller} from '../public/static/calendar-revision-poller.js';
const source=await readFile(new URL('../public/static/app.js',import.meta.url),'utf8');
const section=(start,end)=>source.slice(source.indexOf(start),source.indexOf(end,source.indexOf(start)));
const refresh=section('async function refreshVisibleRosterView(', 'async function loadCloudCalendarEvents(');
const load=section('async function loadCloudCalendarEvents(', 'async function calendarLoadResponseIsRetryable(');
let requests=[],rendered=0,sanitized=0,persisted=0,response={snapshotCurrent:true,calendarRevision:'existing'};
let view={key:'account',overview:false},resolveResponse;
const g={
 currentSnapshot:{calendarRevision:'existing',preview:{events:[]},session:{doctorKey:'DOCTOR'}},currentSnapshotStale:false,currentCalendarRevision:'existing',restoredSessionState:{},
 activeDoctorProfile:null,authUserEmail:'owner@test',authUserPassword:'fixture',currentUserEmail:'user@test',currentUserPassword:'fixture',currentDefaultDoctorKey:'DOCTOR',
 adminViewingEmail:'',cloudAvailable:true,currentRosterClaims:[],OWNER_DOCTOR_KEY:'OWNER',
 calendarRevisionView:()=>view,calendarRevisionRefreshAllowed:()=>true,buildActiveSessionState:()=>({settings:{dateFrom:'2026-10-01'}}),
 calendarTransitionStillCurrent:()=>true,activeCalendarTransitionKey:()=>view.key,normalizeEmail:v=>v,normalizeRosterName:v=>v,
 selectedDoctor:()=>({key:'DOCTOR'}),isCreatorAuthenticated:()=>false,viewedAccountEmail:()=>g.adminViewingEmail||g.currentUserEmail,
 preferredDoctorKeyForAccountEmail:()=> 'DOCTOR',cloudCalendarEventRange:()=>({startDate:'2026-01-01',endDate:'2027-01-06'}),
 fetch:async(_url,options)=>{requests.push(JSON.parse(options.body));return {status:200};},readJsonResponse:async()=>response,
 renderWorkspaceFromSnapshot:()=>rendered++,sanitizeWorkspaceSnapshot:v=>{sanitized++;return v;},applyLoadedCalendarFileRefs:()=>{},
 clearCloudLoadedSnapshotFilters:v=>v,clearCloudLoadedSessionFilters:v=>v,currentCalendarSnapshotCacheKey:()=> 'account:DOCTOR',
 saveCalendarSnapshotCacheForContext:()=>persisted++,saveWorkspaceSnapshotForEmail:()=>persisted++,activeWorkspaceOwnerKey:()=> 'user@test',selectedFiles:[],importRefForWorkspace:v=>v,
};
g.accountFetch=(...args)=>g.fetch(...args);g.hasAccountAuthentication=value=>Boolean(value);
runInNewContext(refresh+load+';this.refresh=refreshVisibleRosterView;this.load=loadCloudCalendarEvents;',g);
for(const target of ['', 'entered-user@test']){
 g.adminViewingEmail=target;
 assert.equal(await g.refresh(view),true);
 assert.equal(requests.at(-1).cachedRevision,'existing');
 assert.equal(requests.at(-1).targetEmail,target);
}
assert.equal(rendered,0);assert.equal(sanitized,0);assert.equal(persisted,0,'unchanged personal/admin calendars do not rewrite browser caches');
g.activeDoctorProfile={id:'profile',doctorKey:'DOCTOR',displayName:'Doctor',sourceTypes:['mmc']};
assert.equal(await g.refresh(view),true);assert.equal(requests.at(-1).cachedRevision,'existing');
assert.equal(rendered,0);assert.equal(sanitized,0);assert.equal(persisted,0,'unchanged profiles are acknowledged without DOM/storage work');
response={snapshot:{preview:{events:[{id:'changed'}]},session:{settings:{}}},calendarRevision:'changed'};
assert.equal(await g.refresh(view),true);assert.equal(rendered,1);assert.equal(g.currentSnapshot.preview.events[0].id,'changed');
assert.equal(g.currentSnapshot.session.settings.dateFrom,'2026-10-01','changed roster retains the current view filters');
g.currentSnapshotStale=true;response={snapshotCurrent:true};await g.refresh(view);assert.equal(requests.at(-1).cachedRevision,'','stale content is never offered as a current revision');
g.currentSnapshotStale=false;
g.fetch=async(_url,options)=>{requests.push(JSON.parse(options.body));return new Promise(resolve=>{resolveResponse=resolve;});};
const stale=g.refresh(view);view={key:'another-account',overview:false};resolveResponse({status:200});
assert.equal(await stale,false);assert.equal(rendered,1,'a late response cannot repaint a different account');
let reads=0,refreshes=0;
const poller=createCalendarRevisionPoller({context:()=>view,eligible:()=>true,read:async()=>{reads++;return 'shared-change';},refresh:async()=>{refreshes++;return true;},schedule:()=>1,cancel:()=>{}});
await poller.tick();await poller.tick();assert.equal(reads,2);assert.equal(refreshes,1,'an unchanged acknowledged result advances the shared fingerprint and avoids repeated refreshes');
console.log('Calendar refresh passes real personal/admin/profile conditional requests, no-op DOM/storage, changed data, view filters and stale-context fencing.');
