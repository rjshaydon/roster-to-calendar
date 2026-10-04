import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { runInNewContext } from 'node:vm';
const app = await readFile(new URL('../public/static/app.js', import.meta.url), 'utf8');
const section=(start,end)=>app.slice(app.indexOf(start),app.indexOf(end,app.indexOf(start)));
const opening=section('async function openFacilityOverview(options = {})','async function openFacilityOverviewByStream()');
let loaded=0;
const noop=()=>{};
const globals={canUseFacilityOverview:()=>true,activeCalendarTransitionKey:()=> 'subject',calendarTransitionRunId:1,
 currentFacilityOverviewMaintenance:false,currentNonClinical:false,facilityOverviewSessionNeedsInitialization:true,resetFacilityOverviewSessionState:()=>{globals.facilityOverviewState.tab='together';},
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
