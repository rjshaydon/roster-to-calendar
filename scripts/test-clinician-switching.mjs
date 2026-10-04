import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { runInNewContext } from 'node:vm';
const app = await readFile(new URL('../public/static/app.js', import.meta.url), 'utf8');
function section(start, end) { return app.slice(app.indexOf(start), app.indexOf(end, app.indexOf(start))); }
const validate = section('async function validateClaimedAccountCalendarInBackground(', 'async function validateDoctorProfileCalendarInBackground(');
let requests = [], rendered = 0;
const globals = {
  currentSnapshot: {preview:{}, cacheKey:'calendar', calendarRevision:'current'}, currentSnapshotStale:false,
  visibleSnapshotIsCurrent:()=>true, normalizeEmail:v=>v, OWNER_EMAIL:'creator@example.com',
  calendarTransitionStillCurrent:()=>true,
  restoreCloudState:async options=> { requests.push(options); globals.enabled = options.adminTargetEmail === 'titus@example.com' || !options.adminTargetEmail; },
  hydrateAuthenticatedWorkspace:async()=> { throw new Error('current calendar must not require full hydration'); },
  renderLoginState:()=>rendered++, enabled:false,
};
runInNewContext(`${validate}; this.validate = validateClaimedAccountCalendarInBackground`,globals);
for (const [email, enabled] of [['tinnie@example.com',false], ['titus@example.com',true], ['creator@example.com',true], ['titus@example.com',true]]) {
  await globals.validate({ownerEmail:email},{preserveRenderedSnapshot:true});
  assert.equal(globals.enabled,enabled,'cached switching must refresh the entered account permission');
}
assert.equal(requests.length,4);
assert.ok(requests.every(request=>request.responseMode==='fast' && request.allowInlineBuild===false));
assert.equal(rendered,4);

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
runInNewContext(`${back}; this.returnToCreator = returnToCreatorAccount`,backGlobals);
void backGlobals.returnToCreator({skipOutgoingSave:true});
assert.equal(closed,true,'Back to creator must immediately close the old overview before waiting for the calendar');
assert.equal(backGlobals.currentUserRole,'creator');
assert.equal(backGlobals.activeDoctorProfile,null);
console.log('Clinician switching passed cached disabled/enabled/Creator transitions, profile access refresh, stale-response denial and immediate overview exit.');
