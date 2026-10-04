import assert from 'node:assert/strict';
import {onShiftLaunchWindow} from '../public/static/shift-launch-policy.js';
const shift={source:'DDH',title:'DDH: PM',start:'2026-10-05T15:00:00',end:'2026-10-06T00:00:00'};
const at=iso=>onShiftLaunchWindow([shift],new Date(iso));
assert.equal(at('2026-10-05T13:59:59+11:00'),null);
assert.equal(at('2026-10-05T14:00:00+11:00').facilityKey,'DDH');
assert.equal(at('2026-10-06T00:59:59+11:00').rosterDate,'2026-10-05');
assert.equal(at('2026-10-06T01:00:00+11:00').rosterDate,'2026-10-05');
assert.equal(at('2026-10-06T01:00:01+11:00'),null);
const night={source:'MMC',title:'MMC: Night',start:'2026-10-03T23:00:00',end:'2026-10-04T08:00:00'};
assert.equal(onShiftLaunchWindow([night],new Date('2026-10-04T08:59:00+11:00')).rosterDate,'2026-10-03','night shifts survive the Melbourne DST change');
assert.equal(onShiftLaunchWindow([night],new Date('2026-10-04T09:00:01+11:00')),null);
assert.equal(onShiftLaunchWindow([{...shift,allDay:true}],new Date('2026-10-05T16:00:00+11:00')),null);
assert.equal(onShiftLaunchWindow([{...shift,title:'Annual Leave'}],new Date('2026-10-05T16:00:00+11:00')),null);
assert.equal(onShiftLaunchWindow([{...shift,start:'2026-10-05T04:00:00Z',end:'2026-10-05T13:00:00Z'}],new Date('2026-10-06T00:30:00+11:00')).rosterDate,'2026-10-05');
console.log('Shift launch windows passed before/after boundaries, midnight, overnight DST, explicit offsets and non-working events.');
// Exercise the actual startup integration, including the contained Creator path.
const {readFile}=await import('node:fs/promises');
const {runInNewContext}=await import('node:vm');
const app=await readFile(new URL('../public/static/app.js',import.meta.url),'utf8');
const section=(start,end)=>app.slice(app.indexOf(start),app.indexOf(end,app.indexOf(start)));
const launch=section('function launchClinicalOnShiftWorkspace(','async function loginWithEmail');
let opened=0;
const noop=()=>{};
const g={clinicalOnShiftStartupPending:true,currentFacilityOverviewAutomaticLaunchEnabled:true,currentFacilityOverviewMaintenance:false,currentNonClinical:false,
 requestClinicalStartupShiftWindow:()=>{},canUseFacilityOverview:()=>true,calendarTransitionStillCurrent:()=>true,currentSnapshot:{preview:{events:[shift]}},currentSnapshotStale:false,
 calendarSnapshotMatchesActiveContext:()=>true,onShiftLaunchWindow,currentFacilityOverviewAccess:{mode:'all'},facilityAccessKeys:a=>a.facilityKeys||[],
 facilityOverviewSessionNeedsInitialization:false,facilityOverviewState:{},applyFacilityOverviewSiteScope:noop,
 openFacilityOverview:async options=>{opened++;assert.equal(options.preserveDate,true);assert.equal(options.preserveFacility,true);},markLoginPhase:noop,setStatus:noop,normalizeAuthMessage:v=>v};
runInNewContext(`${launch};this.launch=launchClinicalOnShiftWorkspace`,g);
assert.equal(g.launch({now:new Date('2026-10-06T00:30:00+11:00')}),true);
assert.equal(g.facilityOverviewState.date,'2026-10-05');
assert.equal(g.facilityOverviewState.facilityKey,'DDH');
assert.equal(g.facilityOverviewState.followOperationalDate,false);
assert.equal(g.launch({now:new Date('2026-10-06T00:30:00+11:00')}),false,'startup cannot repeatedly override navigation');
g.clinicalOnShiftStartupPending=true;g.currentFacilityOverviewAccess={mode:'sites',facilityKeys:['MMC']};
assert.equal(g.launch({now:new Date('2026-10-06T00:30:00+11:00')}),false,'startup cannot bypass roster site permissions');
g.currentFacilityOverviewAccess={mode:'all'};g.currentFacilityOverviewAutomaticLaunchEnabled=false;
assert.equal(g.launch({now:new Date('2026-10-06T00:30:00+11:00')}),false);
g.currentFacilityOverviewAutomaticLaunchEnabled=true;g.currentSnapshotStale=true;
assert.equal(g.launch({now:new Date('2026-10-06T00:30:00+11:00')}),false,'stale calendar cannot trigger an automatic launch');
g.currentSnapshotStale=false;
Object.assign(g,{cancelDeferredAccountContextLoad:noop,cancelDeferredBootstrapImports:noop,renderWorkspaceFromSnapshot:noop,
 restoredSessionState:{},markLoginPaintCommitted:noop,queuePostLoginSnapshotRefresh:noop,queueStoredCalendarSnapshotMaintenance:noop});
const creator=section('function finishContainedCreatorStartup(','function launchNonClinicalDirectorWorkspace');
runInNewContext(`${creator};this.creator=finishContainedCreatorStartup`,g);
g.creator({now:new Date('2026-10-06T00:30:00+11:00')});
assert.equal(opened,2,'contained Creator startup uses the same shift window');
console.log('Startup launch passed Creator, restricted permissions, midnight roster date, rollout disablement and stale-calendar protection.');
const requestHelper=section('function requestClinicalStartupShiftWindow(','function launchClinicalOnShiftWorkspace(');
let reply,requests=0,launches=0;
const startup={clinicalOnShiftWindowPromise:null,clinicalOnShiftStartupPending:true,activeCalendarTransitionKey:()=> 'account',calendarTransitionRunId:1,
 authUserEmail:'test@example.com',authUserPassword:'test',facilityOverviewTargetEmail:()=>'',
 fetch:()=>{requests++;return new Promise(resolve=>{reply=resolve;});},readJsonResponse:async data=>data,
 launchClinicalOnShiftWorkspace:()=>{launches++;}};
runInNewContext(`${requestHelper};this.request=requestClinicalStartupShiftWindow`,startup);
startup.request();startup.request();
assert.equal(requests,1,'the startup fallback adds only one bounded request');
startup.clinicalOnShiftStartupPending=false;reply({shiftWindow:{facilityKey:'DDH'}});
await startup.clinicalOnShiftWindowPromise;
assert.equal(launches,0,'a manual navigation cancels delayed automatic launch');
startup.clinicalOnShiftWindowPromise=null;startup.clinicalOnShiftStartupPending=true;startup.request();
startup.calendarTransitionRunId=2;reply({shiftWindow:{facilityKey:'DDH'}});
await startup.clinicalOnShiftWindowPromise;
assert.equal(launches,0,'another account transition cancels delayed automatic launch');
console.log('Startup fallback passed request deduplication and manual/account navigation cancellation.');
const windowHelpers=section('function currentFacilityOverviewShiftWindow()','function facilityOverviewIsSiteScoped()');
const scopeHelper=section('function applyFacilityOverviewSiteScope()','function resetFacilityOverviewAccessForEnteredUser()');
let boundaryNow=RealDateForTest('2026-11-02T00:30:00+11:00');
function RealDateForTest(value){return Date.parse(value);}
const boundary={Date:{now:()=>boundaryNow},currentFacilityOverviewAccess:{mode:'sites',facilityKeys:['MCH']},
 facilityAccessKeys:a=>a.facilityKeys||[],facilityOverviewState:{tab:'on-shift',date:'2026-11-01',facilityKey:'DDH',
 startupShiftWindow:{facilityKey:'DDH',rosterDate:'2026-11-01',start:Date.parse('2026-11-01T15:00:00+11:00'),end:Date.parse('2026-11-02T00:00:00+11:00')},
 byStreamRows:[{facilityKey:'DDH'}],byStreamCatalog:[]}};
runInNewContext(`${windowHelpers};${scopeHelper};this.apply=applyFacilityOverviewSiteScope`,boundary);
boundary.apply();
assert.equal(boundary.facilityOverviewState.facilityKey,'DDH','the previous shift hospital survives the next-term site scope during the grace hour');
assert.equal(boundary.facilityOverviewState.byStreamRows[0].facilityKey,'MCH','the timed exception does not widen stream access');
boundaryNow=Date.parse('2026-11-02T01:01:00+11:00');boundary.apply();
assert.equal(boundary.facilityOverviewState.facilityKey,'MCH','the previous hospital allowance ends with the grace hour');
console.log('Term-boundary browser scope passed the expiring On shift exception without widening other views.');
