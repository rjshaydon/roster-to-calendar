import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import vm from 'node:vm';
const source=await readFile(new URL('../public/static/app.js',import.meta.url),'utf8');
const section=(start,end)=>source.slice(source.indexOf(start),source.indexOf(end,source.indexOf(start)));
const cards=['alice@example.test','bob@example.test'].map(email=>({dataset:{adminUserEmail:email},classList:{toggle(key,value){this.hidden=value;}}}));
const inputs=cards.map(card=>({dataset:{toggleUserFacilityOverview:card.dataset.adminUserEmail},checked:false,disabled:false}));
const count={textContent:''};
const pending=[];
const context=vm.createContext({
  serverUsers:[{email:'alice@example.test',realName:'Alice Jones',seniorities:['HMO'],claims:[],facilityOverviewEnabled:false},{email:'bob@example.test',realName:'Bob Smith',seniorities:['SMS'],claims:[],facilityOverviewEnabled:false}],
  adminUserSearchQuery:'JONES',adminUserSeniorityFilter:'',adminUserDirectoryRevision:0,pendingAdminPermissions:new Map(),serverUsersUnavailable:false,serverUsersLoaded:true,serverUsersLoading:false,
  accountsBody:{querySelectorAll(selector){return selector==='[data-admin-user-email]'?cards:inputs;},querySelector(){return count;}},
  normalizeEmail:v=>String(v||'').toLowerCase(),normalizeServerUser:v=>v,sanitizeRosterClaims:v=>v||[],isCreatorAuthenticated:()=>true,
  authUserEmail:'owner@example.test',authUserPassword:'fixture',currentUserEmail:'owner@example.test',currentUserPassword:'fixture',
  setStatus(){},fetch(url,options){return new Promise(resolve=>pending.push({resolve,body:JSON.parse(options.body)}));},
  async readJsonResponse(response){if(response.error)throw new Error(response.error);return response;},
});
vm.runInContext(section('function filterAdminUserCards()', 'function renderAccountsModal(')+section('function applyAdminPermissionChoice(', 'async function setUserInsightsEnabled('),context);
context.filterAdminUserCards();
assert.equal(cards[0].classList.hidden,false,'surname search matches locally');
assert.equal(cards[1].classList.hidden,true);
assert.equal(count.textContent,'1 account');
assert.equal(pending.length,0,'typing must make no requests');
const inputHandler=section('accountsBody.addEventListener("input"', 'accountsBody.addEventListener("toggle"');
assert.doesNotMatch(inputHandler,/renderAccountsModal|innerHTML|setSelectionRange/,'typing must retain the actual input and caret');
assert.doesNotMatch(section('async function saveAdminPermission(', 'async function confirmSuggestedClaim('),/renderAccountsModal/,'permission changes must retain their controls');
const alice=context.saveAdminPermission('alice@example.test','facilityOverviewEnabled','setUserFacilityOverviewEnabled',true);
assert.equal(inputs[0].checked,true);assert.equal(inputs[0].disabled,false,'taps remain available during saves');
const aliceOff=context.saveAdminPermission('alice@example.test','facilityOverviewEnabled','setUserFacilityOverviewEnabled',false);
const aliceOn=context.saveAdminPermission('alice@example.test','facilityOverviewEnabled','setUserFacilityOverviewEnabled',true);
assert.equal(await aliceOff,false,'superseded unsent choices are coalesced');
assert.equal(inputs[0].checked,true,'new selection is visible immediately');
assert.equal(pending.length,1,'one request per user may be in flight');
const bob=context.saveAdminPermission('bob@example.test','facilityOverviewEnabled','setUserFacilityOverviewEnabled',true);
pending[1].resolve({user:{...context.serverUsers[1],facilityOverviewEnabled:true}});await bob;
pending[0].resolve({error:'503'});await alice;
assert.equal(context.serverUsers[0].facilityOverviewEnabled,true,'failure cannot remove a newer queued tick');
assert.equal(context.serverUsers[1].facilityOverviewEnabled,true,'failure cannot revert a different successful save');
assert.equal(pending.length,3,'latest queued choice is sent after the first request settles');
pending[2].resolve({user:{...context.serverUsers[0],facilityOverviewEnabled:true}});await aliceOn;
assert.equal(inputs[0].disabled,false);assert.equal(inputs[1].checked,true);
assert.equal(context.pendingAdminPermissions.size,0);
// Two different permissions for one person can be selected immediately.
const who=context.saveAdminPermission('alice@example.test','insightsEnabled','setUserInsightsEnabled',true);
const overview=context.saveAdminPermission('alice@example.test','facilityOverviewEnabled','setUserFacilityOverviewEnabled',false);
assert.equal(inputs[0].checked,false);
assert.equal(pending.length,4,'different permissions still save in order');
pending[3].resolve({user:{...context.serverUsers[0],insightsEnabled:true,facilityOverviewEnabled:true}});await who;
assert.equal(inputs[0].checked,false,'old full-user response retains the newer overview choice');
assert.equal(pending.length,5);
pending[4].resolve({error:'503'});await overview;
assert.equal(inputs[0].checked,true,'failed final choice rolls back to confirmed state');
assert.equal(context.serverUsers[0].insightsEnabled,true,'rollback keeps the successfully saved other permission');
assert.match(section('async function loadServerUsers()', 'function adminParserSurfaceReady'),/directoryRevision !== adminUserDirectoryRevision/,'old directory responses cannot overwrite newer permission changes');
console.log('Admin search and permission interaction regressions passed.');

const provider = await readFile(new URL('../functions/_lib/findmyshift.js', import.meta.url), 'utf8');
assert.doesNotMatch(provider, /^import.*from ["']xlsx["']/m, 'cold account requests must not initialize the spreadsheet library');

// Opening before authentication readiness must not manufacture an empty list.
const loads=[];let panels=0;
const loading=vm.createContext({
  serverUsersRequest:null,serverUsersLoaded:false,serverUsersLoading:false,serverUsersUnavailable:false,
  adminUserDirectoryRevision:0,serverUsers:[],pendingAdminPermissions:new Map(),
  adminViewingEmail:'',currentUserEmail:'owner@example.test',currentUserPassword:'fixture',authUserEmail:'owner@example.test',authUserPassword:'fixture',OWNER_EMAIL:'owner@example.test',cloudAvailable:false,
  normalizeEmail:v=>String(v||'').toLowerCase(),isCreatorAuthenticated:()=>true,isViewingCreatorAccount:()=>true,
  accountsModal:{classList:{contains:()=>false}},renderAccountsModal(){panels++;},
  fetch(){return new Promise(resolve=>loads.push(resolve));},async readJsonResponse(r){if(r.error)throw Error(r.error);return r;},
  calendarLoadResponseIsRetryable:async()=>false,applyAuthoritativeAvailableDoctors(){},applyIssueConfig(){},syncAccountsButton(){},
});
vm.runInContext(section('async function loadServerUsers()', 'function adminParserSurfaceReady'),loading);
await loading.loadServerUsers();assert.equal(loads.length,0);assert.equal(loading.serverUsersLoaded,false);
loading.cloudAvailable=true;
const firstLoad=loading.loadServerUsers(),reopened=loading.loadServerUsers();
assert.equal(loads.length,1,'reopening shares the pending request');assert.equal(loading.serverUsersLoading,true);
loads[0]({status:200,users:[{email:'stella@example.test',realName:'Stella Robinson'}]});await Promise.all([firstLoad,reopened]);
assert.equal(loading.serverUsersLoaded,true);assert.equal(loading.serverUsers[0].realName,'Stella Robinson');assert.equal(loading.serverUsersLoading,false);assert.ok(panels>=2,'completion updates the open panel');
const failedLoad=loading.loadServerUsers();loads[1]({status:500,error:'network failure'});await failedLoad;
assert.equal(loading.serverUsersUnavailable,true,'failure is distinct from no accounts');assert.equal(loading.serverUsers.length,1,'failed refresh retains cached users');
const retriedLoad=loading.loadServerUsers();loads[2]({status:200,users:[{email:'vidya@example.test',realName:'Vidya Achan'}]});await retriedLoad;
assert.equal(loading.serverUsersUnavailable,false);assert.equal(loading.serverUsers[0].realName,'Vidya Achan');
console.log('Directory readiness, coalescing, completion and recovery tests passed.');

// Approved durable links are shown even when legacy claims are empty.
const labels=vm.createContext({sanitizeRosterClaims:v=>v||[],escapeHtml:v=>String(v)});
vm.runInContext(section('function renderAdminUserClaims(', 'function adminIssueCount('),labels);
assert.match(labels.renderAdminUserClaims({realName:'Stella Robinson',claims:[],rosterLinks:[{sourceType:'mch',displayName:'Stella Robinson'}]}),/MCH/);
assert.doesNotMatch(section('function syncMobileChrome()', 'function normalizeAdminFilesSortOrder('),/loadServerUsers/);
assert.match(source.slice(source.indexOf('function renderLoginState()'),source.indexOf('function renderLoginState()')+600),/!serverUsersLoading && !serverUsersRequest/);

// A failed creation response may follow committed writes. Recover once with
// login; never re-create, loop, or retry the database daily allowance.
const recoveryCalls=[];
let replies=[];
const recovery=vm.createContext({fetch:async(url,options)=>{recoveryCalls.push(JSON.parse(options.body));return replies.shift();}});
vm.runInContext(section('async function requestAccountLogin(', 'async function restoreCloudState(')+section('async function calendarLoadResponseIsRetryable(', 'function cloudCalendarEventRange('),recovery);
const reply=(status,payload={})=>({status,ok:status===200,clone:()=>({json:async()=>payload})});
replies=[reply(503),reply(200)];
assert.equal((await recovery.requestAccountLogin({action:'login',mode:'create',email:'fixture@example.test',password:'fixture'})).status,200);
assert.deepEqual(recoveryCalls.map(c=>c.mode),['create','login']);
assert.equal(recoveryCalls[1].allowInlineBuild,false);
for(const errorType of ['request-statement-limit','account-daily-quota']) {
 recoveryCalls.length=0;replies=[reply(503,{errorType})];
 assert.equal((await recovery.requestAccountLogin({action:'login',mode:'create'})).status,503);
 assert.equal(recoveryCalls.length,1,'database safeguard must not be retried');
}
recoveryCalls.length=0;replies=[reply(503),reply(401)];
assert.equal((await recovery.requestAccountLogin({action:'login',mode:'create'})).status,503,'if creation never committed retain the original error');
assert.equal(recoveryCalls.length,2);
recoveryCalls.length=0;replies=[reply(503)];
await recovery.requestAccountLogin({action:'login',mode:'login'});assert.equal(recoveryCalls.length,1);
console.log('Signup recovery passed one login-only retry, no repeat creation and no database-limit retries.');
