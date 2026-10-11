import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import vm from 'node:vm';

const source = await readFile(new URL('../public/static/app.js', import.meta.url), 'utf8');
function extract(name) {
 const start = source.indexOf(`function ${name}(`);
 const end = source.indexOf('\nfunction ', start + 1);
 const asyncEnd = source.indexOf('\nasync function ', start + 1);
 return source.slice(source.slice(start - 6, start) === 'async ' ? start - 6 : start,
  Math.min(...[end, asyncEnd].filter(n => n >= 0)));
}
const state = {tab:'on-shift',date:'2026-10-11',facilityKey:'MMC',requestId:1,
 contactAccessToken:'token-0',onShiftRevision:'roster',onShiftData:[],
 contactList:{status:'available',revision:'contacts',sourceDate:'2026-10-11',contacts:[{phone:'12010'}]},content:'12010'};
let renders = 0, renewals = 0, expired = false, mode = 'normal', stale = false;
const calls = [];
const noop = () => {};
const context = vm.createContext({
 facilityOverviewState:state,currentFacilityOverviewMaintenance:false,document:{hidden:false},
 facilityOverviewBody:{scrollTop:0},scrollFacilityOverviewToInitialShift:()=>{throw Error('Background renewal must not jump to a shift');},
 facilityOverviewContactRefreshInFlight:false,authUserEmail:'fixture',authUserPassword:'fixture',
 currentUserEmail:'',currentUserPassword:'',console:{warn:noop},
 canUseFacilityOverview:()=>true,canUseFullFacilityOverview:()=>true,isFacilityOverviewOpen:()=>true,
 applyFacilityOverviewSiteScope:noop,stopFacilityOverviewContactRefresh:noop,
 collapseFacilityOverviewContactReview:noop,scheduleFacilityOverviewContactRefresh:noop,
 contactOperationalDate:()=>state.date,contactExtractHasExpired:()=>false,
 renderFacilityOverview:()=>renders++,renderFacilityOverviewOnShiftPreservingViewport:()=>renders++,
 renderFacilityOverviewOnShiftResults:()=>state.contactList?.contacts?.[0]?.phone || '',
 loadFacilityOverviewSnapshot:async()=>null,storeFacilityOverviewSnapshot:noop,
 beginFacilityOverviewDataRequest:()=>({signal:undefined}),finishFacilityOverviewDataRequest:noop,
 facilityOverviewTargetEmail:()=>'',refreshFacilityOverviewSnapshotAccess:noop,
 mergeContactResolutionRefresh:(_old,next)=>next,renderFacilityOverviewCoverageNotice:()=>'',
 facilityOverviewRequestWasCancelled:()=>false,facilityOverviewCachedReadDenied:e=>e.status===403 || e.status===401,
 escapeHtml:String,
 async readJsonResponse(response) {
  if(response.status!==200)throw Object.assign(new Error('Access denied'),{status:response.status});
  return response.data;
 },
 async fetch(_url, options) {
  const body=JSON.parse(options.body); calls.push(body);
  if(body.action==='queryFacilityOverviewContactList') {
   assert.equal(body.password,undefined,'routine contact polls never authenticate against D1');
   if(stale) {state.requestId++;state.facilityKey='MCH';return {status:403};}
   if(expired)return {status:403};
   if(mode==='changed')return {status:200,data:{contactList:{...state.contactList,revision:'new-contacts',contacts:[{phone:'12020'}]}}};
   if(mode==='heartbeat')return {status:200,data:{contactList:{...state.contactList,syncHealth:{status:'healthy',lastSuccessAt:'2026-10-11T02:00:00Z'}}}};
   return {status:200,data:{unchanged:true}};
  }
  assert.equal(body.action,'queryFacilityOverviewOnShift');
  assert.equal(body.password,'fixture','renewal rechecks authenticated site access');
  assert.equal(state.content,'12010','renewal must not blank the existing view');
  assert.equal(state.contactList.contacts[0].phone,'12010');
  renewals++;
  if(mode==='network')throw new Error('Temporary network failure');
  if(mode==='denied')return {status:403};
  expired=false;
  return {status:200,data:{rosterUnchanged:true,revision:'roster',
   contactAccessToken:`token-${renewals}`,contactList:state.contactList}};
 }
});
vm.runInContext(extract('facilityOverviewContactRefreshIsActive'),context);
vm.runInContext(extract('loadFacilityOverviewOnShift'),context);
vm.runInContext(extract('refreshFacilityOverviewContactList'),context);

// Simulate a ten-hour shift: unchanged minute polls, with authentication only
// when the server rejects the fifteen-minute token, no page refresh or flicker.
for(let minute=1;minute<=600;minute++) {
 expired=minute%15===0;
 await context.refreshFacilityOverviewContactList();
}
assert.equal(renewals,40); assert.equal(renders,0);
assert.equal(calls.filter(c=>c.action==='queryFacilityOverviewContactList').length,600);
assert.equal(state.contactAccessToken,'token-40');
state.facilityKey='DDH';
await context.refreshFacilityOverviewContactList();
assert.equal(renders,0,'unchanged DDH expiry calculations must not redraw the page');
mode='heartbeat';
await context.refreshFacilityOverviewContactList();
assert.equal(renders,0,'healthy sync timestamps must not redraw unchanged allocations');
mode='changed';
await context.refreshFacilityOverviewContactList();
assert.equal(state.content,'12020');assert.equal(renders,1,'a changed allocation updates the displayed phone');
// Restore the original allocation for renewal failure/revocation scenarios.
state.contactList={...state.contactList,revision:'contacts',contacts:[{phone:'12010'}]};state.content='12010';renders=0;

expired=true;mode='network';
await context.refreshFacilityOverviewContactList();
assert.equal(state.content,'12010'); assert.equal(renders,0);
mode='normal';
await context.refreshFacilityOverviewContactList();
assert.equal(expired,false,'temporary renewal failure remains retryable');

const before=renewals;stale=true;
await context.refreshFacilityOverviewContactList();
assert.equal(renewals,before,'a late denial from a previous site must not renew or clear the new site');
assert.equal(state.contactList.contacts[0].phone,'12010');
stale=false;expired=true;mode='denied';
await context.refreshFacilityOverviewContactList();
assert.equal(state.contactAccessToken,''); assert.equal(state.contactList,null);
assert.equal(state.onShiftData,null,'actual authenticated access denial removes the protected view');
assert.equal(context.facilityOverviewContactRefreshIsActive(),false);

const loading=vm.createContext({facilityOverviewState:{contactList:null}});
vm.runInContext(extract('renderFacilityOverviewContactListStatus'),loading);
assert.equal(loading.renderFacilityOverviewContactListStatus([],[]),'','loading is not a contact sync failure');
console.log('Whole-shift contact session passed: automatic renewal, unchanged views, transient failures and revoked access.');
