import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { runInNewContext } from 'node:vm';
const app = await readFile(new URL('../public/static/app.js', import.meta.url), 'utf8');
const section=(start,end)=>app.slice(app.indexOf(start),app.indexOf(end,app.indexOf(start)));
const opening=section('function openFacilityOverview(options = {})','async function openFacilityOverviewByStream()');
let loaded=0;
const noop=()=>{};
const globals={facilityOverviewOpeningPromise:null,facilityOverviewOpeningRunId:0,facilityOverviewNavigationLocked:false,facilityOverviewIgnoreToggleUntil:0,isFacilityOverviewOpen:()=>true,currentFacilityOverviewShiftWindow:()=>null,applyFacilityOverviewSiteScope:noop,canUseFacilityOverview:()=>true,activeCalendarTransitionKey:()=> 'subject',calendarTransitionRunId:1,
 currentFacilityOverviewAccessReady:true,currentFacilityOverviewMaintenance:false,currentNonClinical:false,facilityOverviewSessionNeedsInitialization:true,resetFacilityOverviewSessionState:()=>{globals.facilityOverviewState.tab='together';globals.facilityOverviewSessionNeedsInitialization=false;},
 facilityOverviewState:{tab:'together',preferredFacilityKey:'DDH',facilityKey:'MMC'},
 refreshFacilityOverviewPreferredFacility:noop,loadFacilityOverviewAvailableTerms:async()=>{},
 currentDirectorViewEnabled:false,contactOperationalDate:()=> '2026-10-04',formatDateKey:()=> '2026-10-04',
 currentFacilityOverviewAccess:{mode:'all'},loadFacilityOverviewMetadata:async()=>{},facilityOverviewFacilityOptions:()=>['MMC','DDH'],
 form:{classList:{add:noop}},previewSection:{classList:{add:noop}},facilityOverviewSection:{classList:{remove:noop}},
 resetFacilityOverviewScroll:noop,syncFacilityOverviewNavigationState:noop,renderFacilityOverview:noop,
 loadFacilityOverviewOnShift:async()=>{loaded++;},loadFacilityOverviewTogether:()=>{throw Error('normal opening must not load Together');},
 loadFacilityOverviewStaff:async()=>{},};
runInNewContext(`${opening}; this.open = openFacilityOverview`,globals);
await globals.open();
assert.equal(globals.facilityOverviewState.tab,'on-shift','normal opening overrides the previous Together tab');
assert.equal(globals.facilityOverviewState.facilityKey,'DDH','opening uses the calendar-derived working-week hospital');
assert.equal(globals.facilityOverviewState.date,'2026-10-04');
assert.equal(loaded,1,'opening requests the selected date and hospital');
globals.facilityOverviewState.tab='staff';
await globals.open({preserveStaffTerm:true});
assert.equal(globals.facilityOverviewState.tab,'staff','explicit staff links keep their destination');
const tabHandler=section('  const tab = event.target.closest("[data-facility-overview-tab]");','  if (event.target.closest("[data-facility-overview-by-stream-add]"))');
const calls=[];
const tabGlobals={event:{target:{closest:()=>({dataset:{facilityOverviewTab:'on-shift'}})}},
 facilityOverviewState:{tab:'together',requestId:0,byStreamRequestId:0,facilityKey:'DDH',date:'2026-10-04'},
 cancelFacilityOverviewDataRequest:noop,collapseFacilityOverviewContactReview:noop,clearFacilityOverviewStaffMultiSelect:noop,
 applyFacilityOverviewSiteScope:()=>calls.push('scope'),resetFacilityOverviewScroll:noop,
 loadFacilityOverviewOnShift:()=>calls.push('load'),};
runInNewContext(`function changeTab(){${tabHandler}}; changeTab()`,tabGlobals);
assert.deepEqual(calls,['scope','load'],'returning to On shift applies access scope and requests the roster immediately');
assert.equal(tabGlobals.facilityOverviewState.date,'2026-10-04');
assert.equal(tabGlobals.facilityOverviewState.facilityKey,'DDH');
console.log('Overview opening and tab navigation passed working-week hospital selection, automatic On shift load and explicit staff-link preservation.');

const eventHelpers=app.match(/function eventRosterDateKey[\s\S]*?(?=\nfunction filterWhenInsightEvents)/)[0];
const dateHelpers=app.match(/function parseDateOnly[\s\S]*?(?=\nfunction formatLongDate)/)[0];
const preferredHelpers=section('function facilityOverviewMelbourneClock','function refreshFacilityOverviewPreferredFacility');
const preferred=new Function(`${eventHelpers}\n${dateHelpers}\n${preferredHelpers}\nreturn facilityOverviewPreferredFacilityFromEvents;`)();
assert.equal(preferred([
 {source:'mmc',title:'MMC: Shift',start:'2026-08-03T08:00:00',end:'2026-08-03T17:00:00'},
 {source:'ddh',title:'DDH: Shift',start:'2026-08-11T08:00:00',end:'2026-08-11T17:00:00'},
],{today:'2026-08-10',now:new Date('2026-08-10T00:00:00Z'),linkedSourceTypes:['mmc','ddh']}).facilityKey,'DDH','this week wins over another linked hospital or last week');
// A slow server must not leave the calendar visible or let a repeated tap
// cancel opening. Closing explicitly invalidates the delayed continuation.
let finishRoster;
let visible=false, dataLoads=0;
const delayed={...globals,facilityOverviewOpeningPromise:null,facilityOverviewOpeningRunId:0,facilityOverviewIgnoreToggleUntil:0,facilityOverviewNavigationLocked:false,
 facilityOverviewSessionNeedsInitialization:false,facilityOverviewState:{tab:'on-shift',preferredFacilityKey:'DDH',facilityKey:'DDH'},
 facilityOverviewSection:{classList:{remove:()=>{visible=true;}}},isFacilityOverviewOpen:()=>visible,
 loadFacilityOverviewMetadata:()=>{throw Error('On shift must not wait for stream metadata');},loadFacilityOverviewOnShift:()=>{dataLoads++;return new Promise(resolve=>{finishRoster=resolve;});},
 clinicalOnShiftStartupPending:true,closeFacilityOverview:()=>{throw Error('second tap must not close a pending opening');},setStatus:noop};
const toggle=section('function toggleFacilityOverview()','facilityOverviewButton?.addEventListener');
runInNewContext(`${opening}; ${toggle}; this.open = openFacilityOverview;this.toggle=toggleFacilityOverview`,delayed);
delayed.toggle();
assert.equal(visible,true,'the overview loading shell is shown before metadata returns');
const pending=delayed.facilityOverviewOpeningPromise;
delayed.toggle();
assert.equal(delayed.facilityOverviewOpeningPromise,pending,'repeated taps reuse the opening request');
finishRoster();await pending;
assert.equal(dataLoads,1);
// Immediately repeated taps after a fast response are ignored too.
delayed.toggle();
assert.equal(visible,true);
delayed.facilityOverviewIgnoreToggleUntil=0;visible=false;dataLoads=0;
const cancelled=delayed.open();
delayed.facilityOverviewOpeningRunId++;visible=false;
finishRoster();await cancelled;
assert.equal(visible,false,'an explicit return to calendar must not be reversed by delayed roster completion');
console.log('Slow opening passed immediate feedback, duplicate-tap protection and cancelled-opening guards.');

// Cold startup paints the cached calendar before the authenticated capability
// envelope arrives. Opening must wait for that existing request, without polls.
const accessHelpers=section('function cancelFacilityOverviewAccessWait()', 'function renderFacilityOverviewMaintenance(');
let coldLoads=0;
const messages=[];
let coldVisible=false;
const cold={...globals,facilityOverviewNavigationLocked:false,facilityOverviewIgnoreToggleUntil:0,currentFacilityOverviewAccessReady:false,currentFacilityOverviewMaintenance:true,
 facilityOverviewAccessWaiters:new Set(),facilityOverviewOpeningPromise:null,facilityOverviewOpeningRunId:0,
 facilityOverviewSessionNeedsInitialization:false,facilityOverviewState:{tab:'on-shift',preferredFacilityKey:'DDH',facilityKey:'DDH'},
 facilityOverviewSection:{classList:{remove:()=>{coldVisible=true;}}},isFacilityOverviewOpen:()=>coldVisible,
 renderFacilityOverviewMaintenance:({loading=false}={})=>{coldVisible=true;messages.push(loading?'loading':'maintenance');},
 loadFacilityOverviewOnShift:async()=>{coldLoads++;},
 clinicalOnShiftStartupPending:true,closeFacilityOverview:()=>{throw Error('startup second tap must not close');},setStatus:noop};
runInNewContext(`${accessHelpers}; ${opening}; ${toggle}; this.open=openFacilityOverview;this.toggle=toggleFacilityOverview;this.ready=markFacilityOverviewAccessReady;this.cancelWait=cancelFacilityOverviewAccessWait`,cold);
cold.toggle();
const coldOpening=cold.facilityOverviewOpeningPromise;
assert.deepEqual(messages,['loading']);
assert.equal(coldLoads,0,'no roster request before authenticated access is ready');
cold.toggle();assert.equal(cold.facilityOverviewOpeningPromise,coldOpening);
cold.currentFacilityOverviewMaintenance=false;cold.ready();await coldOpening;
assert.equal(coldLoads,1,'startup capability arrival automatically loads the selected roster');
assert.deepEqual(messages,['loading'],'a temporary default never displays maintenance');
// A real maintenance response still blocks all roster queries.
cold.currentFacilityOverviewAccessReady=false;cold.currentFacilityOverviewMaintenance=true;
const maintenanceOpening=cold.open();cold.ready();await maintenanceOpening;
assert.equal(messages.at(-1),'maintenance');assert.equal(coldLoads,1);
// Explicit close and account transitions cancel the waiter. Late responses
// cannot reopen a closed view or render the previous account's roster.
for(const cancel of ['close','switch']) {
 cold.currentFacilityOverviewAccessReady=false;const pending=cold.open();
 cold.facilityOverviewOpeningRunId++;cold.cancelWait();coldVisible=false;
 cold.currentFacilityOverviewMaintenance=false;cold.ready();await pending;
 assert.equal(coldVisible,false);assert.equal(coldLoads,1,cancel+' must cancel the pending opening');
 assert.equal(cold.facilityOverviewAccessWaiters.size,0);
}
const maintenanceRenderer=section('function renderFacilityOverviewMaintenance(', 'function facilityOverviewMelbourneClock');
assert.match(maintenanceRenderer,/loading \? "Loading At a glance…"/);
for(const fn of ['applyCloudStateIdentity','applyCloudStateContext','loadDoctorProfileFacilityOverviewAccess']) {
 const body=section((fn.startsWith('load')?'async ':'')+'function '+fn+'(', '\n'+(fn==='loadDoctorProfileFacilityOverviewAccess'?'async ':'')+'function ');
 assert.match(body,/syncFacilityOverviewAccess\(\);[\s\S]*markFacilityOverviewAccessReady\(\)/,fn+' must wake the opener after applying permissions');
}
console.log('Cold startup passed loading feedback, automatic refresh, deduplication, confirmed maintenance and close/account cancellation with zero additional startup requests.');
