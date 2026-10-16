import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { runInNewContext } from 'node:vm';
const app = await readFile(new URL('../public/static/app.js', import.meta.url), 'utf8');
function section(start, end) { return app.slice(app.indexOf(start), app.indexOf(end, app.indexOf(start))); }
const validate = section('async function validateClaimedAccountCalendarInBackground(', 'async function validateDoctorProfileCalendarInBackground(');
let requests = [], rendered = 0, startupChecks = [];
const globals = {
  currentSnapshot: {preview:{}, cacheKey:'calendar', calendarRevision:'current'}, currentSnapshotStale:false,
  visibleSnapshotIsCurrent:()=>true, normalizeEmail:v=>v, OWNER_EMAIL:'creator@example.com',
  calendarTransitionStillCurrent:()=>true,
  restoreCloudState:async options=> { requests.push(options); globals.enabled = options.adminTargetEmail === 'titus@example.com' || !options.adminTargetEmail; },
  hydrateAuthenticatedWorkspace:async()=> { throw new Error('current calendar must not require full hydration'); },
  renderLoginState:()=>rendered++, enabled:false,
  launchClinicalOnShiftWorkspace:options=>startupChecks.push({enabled:globals.enabled,transition:options.transition}),
};
runInNewContext(`${validate}; this.validate = validateClaimedAccountCalendarInBackground`,globals);
for (const [email, enabled] of [['tinnie@example.com',false], ['titus@example.com',true], ['creator@example.com',true], ['titus@example.com',true]]) {
  await globals.validate({ownerEmail:email},{preserveRenderedSnapshot:true});
  assert.equal(globals.enabled,enabled,'cached switching must refresh the entered account permission');
}
assert.equal(requests.length,4);
assert.ok(requests.every(request=>request.responseMode==='fast' && request.allowInlineBuild===false));
assert.equal(rendered,4);
assert.deepEqual(startupChecks.map(check=>check.enabled),[false,true,true,true], 'cached entry checks shift eligibility after restoring each account permission');

// Exercise cached God mode entry with the actual startup and eligibility
// functions. The viewed account remains restricted until the server confirms.
let eligibilityReply, authReply, eligibilityCount=0, opened=0, transition=1;
const noop=()=>{};
const cachedEntry={
 rosterComparisonCache:{clear:()=>{}},
 currentSnapshot:{preview:{events:[]},cacheKey:'tony',calendarRevision:'current'},currentSnapshotStale:false,
 visibleSnapshotIsCurrent:()=>true,calendarSnapshotMatchesActiveContext:()=>true,normalizeEmail:v=>v,OWNER_EMAIL:'creator@example.com',
 calendarTransitionStillCurrent:t=>t===transition,activeCalendarTransitionKey:()=>`account:${transition}`,calendarTransitionRunId:1,
 restoreCloudState:()=>new Promise(resolve=>{authReply=()=>{cachedEntry.currentOnShiftForAllEnabled=true;cachedEntry.currentFacilityOverviewMaintenance=false;resolve();};}),
 hydrateAuthenticatedWorkspace:async()=>{throw Error('cached entry must not rebuild the calendar');},renderLoginState:noop,
 clinicalOnShiftStartupPending:true,clinicalOnShiftWindowPromise:null,currentFacilityOverviewAutomaticLaunchEnabled:true,
 currentFacilityOverviewMaintenance:false,currentNonClinical:false,currentOnShiftForAllEnabled:false,
 canUseFacilityOverview:()=>false,canUseFullFacilityOverview:()=>false,activeCalendarMode:()=> 'claimed-account',
 authUserEmail:'creator@example.com',authUserPassword:'fixture',facilityOverviewTargetEmail:()=> 'tony@example.com',
 fetch:async(url,request)=>{eligibilityCount++;const body=JSON.parse(request.body);assert.equal(body.action,'queryFacilityOverviewLaunchWindow');assert.equal(body.targetEmail,'tony@example.com');return new Promise(resolve=>{eligibilityReply=resolve;});},
 readJsonResponse:async response=>response,applyOnShiftAccessWindow:window=>{cachedEntry.accessWindow=window;},syncFacilityOverviewAccess:noop,
 facilityOverviewSessionNeedsInitialization:false,facilityOverviewState:{requestId:0,byStreamRequestId:0},applyFacilityOverviewSiteScope:noop,
 clearOnShiftAccess:noop,cancelFacilityOverviewAccessWait:noop,syncFacilityOverviewAccess:noop,formatDateKey:()=> '2026-10-13',
 facilityOverviewOpeningRunId:0,cancelFacilityOverviewDataRequest:noop,collapseFacilityOverviewContactReview:noop,
 stopFacilityOverviewContactRefresh:noop,clearFacilityOverviewStaffMultiSelect:noop,resetFacilityOverviewScroll:noop,
 syncFacilityOverviewNavigationState:noop,form:null,facilityOverviewSection:null,previewSection:null,latestPreview:null,
 openFacilityOverview:async()=>{opened++;},markLoginPhase:noop,
};
cachedEntry.accountFetch=(...args)=>cachedEntry.fetch(...args);cachedEntry.hasAccountAuthentication=value=>Boolean(value);
runInNewContext(`${validate};${section('function requestClinicalStartupShiftWindow(', 'async function loginWithEmail(')};this.validate=validateClaimedAccountCalendarInBackground`,cachedEntry);
runInNewContext(`${section('function resetFacilityOverviewAccessForEnteredUser(', 'function facilityOverviewTargetEmail(')};${section('function closeFacilityOverview(', 'function cancelFacilityOverviewDataRequest(')};this.reset=resetFacilityOverviewAccessForEnteredUser;this.close=closeFacilityOverview`,cachedEntry);
cachedEntry.reset();
assert.equal(cachedEntry.clinicalOnShiftStartupPending,true,'account cleanup must preserve the new startup intent');
const enter=()=>cachedEntry.validate({ownerEmail:'tony@example.com'},{preserveRenderedSnapshot:true,transition});
const pending=enter();assert.equal(eligibilityCount,0,'no shift check before the access envelope is authenticated');
authReply();await pending;assert.equal(eligibilityCount,1,'a current calendar still checks the entered staff shift');
assert.equal(opened,0,'personal cache cannot grant overview access');
const confirmed={facilityKey:'DDH',rosterDate:'2026-10-13',start:1791889200000,end:1791923400000};
eligibilityReply({shiftWindow:confirmed});await cachedEntry.clinicalOnShiftWindowPromise;
assert.equal(opened,1);assert.equal(cachedEntry.facilityOverviewState.date,confirmed.rosterDate);assert.equal(cachedEntry.facilityOverviewState.facilityKey,'DDH');
assert.equal(cachedEntry.clinicalOnShiftStartupPending,false);

cachedEntry.clinicalOnShiftStartupPending=true;cachedEntry.clinicalOnShiftWindowPromise=null;
const off=enter();authReply();await off;eligibilityReply({shiftWindow:null});await cachedEntry.clinicalOnShiftWindowPromise;
assert.equal(opened,1,'off-shift cached accounts stay on the calendar');

cachedEntry.clinicalOnShiftWindowPromise=null;const staleEntry=enter();transition=2;cachedEntry.calendarTransitionRunId=2;authReply();await staleEntry;
assert.equal(eligibilityCount,2,'a superseded account entry cannot launch a shift check');
cachedEntry.close();assert.equal(cachedEntry.clinicalOnShiftStartupPending,false,'manual close still cancels pending automatic opening');
console.log('Cached God mode startup passed authenticated eligibility, restricted automatic opening, off-shift behaviour and stale account cancellation.');

const access = section('async function loadDoctorProfileFacilityOverviewAccess(', 'async function loadUnclaimedSourceImports(');
const profile = {id:'titus',doctorKey:'TITUS'};
let resolveResponse;
const profileGlobals = {
  activeDoctorProfile:profile, isCreatorAuthenticated:()=>true, activeCalendarMode:()=> 'doctor-profile',
  authUserEmail:'creator@example.com',authUserPassword:'test',normalizeEmail:v=>v,
  fetch:async()=>new Promise(resolve=> { resolveResponse=resolve; }),
  readJsonResponse:async response=>response, calendarTransitionStillCurrent:t=>t===profileGlobals.transition,
  transition:2,currentFacilityOverviewEnabled:false,sanitizeFacilityOverviewAccess:v=>v,
  applyFacilityOverviewSiteScope:()=>{},markFacilityOverviewAccessReady:()=>{},syncFacilityOverviewAccess:()=>{},console,
};
profileGlobals.accountFetch=(...args)=>profileGlobals.fetch(...args);profileGlobals.hasAccountAuthentication=value=>Boolean(value);
runInNewContext(`${access}; this.loadAccess = loadDoctorProfileFacilityOverviewAccess`,profileGlobals);
const stale = profileGlobals.loadAccess(profile,{transition:1});
resolveResponse({facilityOverviewEnabled:true,facilityOverviewAccess:{mode:'all'}});
await stale;
assert.equal(profileGlobals.currentFacilityOverviewEnabled,false,'a previous visit to the same profile cannot overwrite the new transition');
const active = profileGlobals.loadAccess(profile,{transition:2});
resolveResponse({facilityOverviewEnabled:true,facilityOverviewMaintenance:false,facilityOverviewAccess:{mode:'sites',facilityKeys:['DDH']},facilityOverviewAccountEmail:'titus@example.com'});
await active;
assert.equal(profileGlobals.currentFacilityOverviewEnabled,true);
assert.equal(profileGlobals.activeDoctorProfile.facilityOverviewAccountEmail,'titus@example.com');
assert.equal(profileGlobals.currentFacilityOverviewAccess.mode,'sites','entered profile must keep its restricted scope');

// Exercise the real return function while its calendar cache is still loading.
const back = section('async function returnToCreatorAccount(', 'async function clearLocalWorkspace(');
let closed = false;
const backGlobals = {
  captureCalendarViewState:()=>({}),beginFacilityOverviewAccountSession:()=>{},
  resetFacilityOverviewAccessForEnteredUser:()=> { closed=true; },performance:{now:()=>1},
  authUserEmail:'creator@example.com',authUserPassword:'test',OWNER_EMAIL:'creator@example.com',OWNER_DOCTOR_KEY:'CREATOR',
  normalizeEmail:v=>v,cancelScheduledCloudStateSave:()=>{},setActiveCalendarContext:()=>{},
  primeInsightsAccessForCurrentView:()=>{},beginCalendarTransition:()=>1,localStorage:{setItem(){}},sessionStorage:{setItem(){}},
  CURRENT_EMAIL_KEY:'email',CURRENT_PASSWORD_KEY:'password',forceConsoleSkin:()=>{},setStatus:()=>{},
  forceCreatorDoctorSession:()=>{},restoreCreatorImportFilesIfNeeded:()=>{},accountCalendarContextForEmail:()=>({}),
  renderCachedCalendarSnapshotForContextAsync:()=>new Promise(()=>{}),
};
backGlobals.accountTransport={expected:()=>false};backGlobals.hasAccountAuthentication=value=>Boolean(value);
runInNewContext(`${back}; this.returnToCreator = returnToCreatorAccount`,backGlobals);
void backGlobals.returnToCreator({skipOutgoingSave:true});
assert.equal(closed,true,'Back to creator must immediately close the old overview before waiting for the calendar');
assert.equal(backGlobals.currentUserRole,'creator');
assert.equal(backGlobals.activeDoctorProfile,null);
console.log('Clinician switching passed cached disabled/enabled/Creator transitions, profile access refresh, stale-response denial and immediate overview exit.');
