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
  adminUserSearchQuery:'JONES',adminUserSeniorityFilter:'',adminUserDirectoryRevision:0,pendingAdminPermissions:new Map(),serverUsersUnavailable:false,
  accountsBody:{querySelectorAll(selector){return selector==='[data-admin-user-email]'?cards:inputs;},querySelector(){return count;}},
  normalizeEmail:v=>String(v||'').toLowerCase(),normalizeServerUser:v=>v,sanitizeRosterClaims:v=>v||[],isCreatorAuthenticated:()=>true,
  authUserEmail:'owner@example.test',authUserPassword:'fixture',currentUserEmail:'owner@example.test',currentUserPassword:'fixture',
  setStatus(){},fetch(url,options){return new Promise(resolve=>pending.push({resolve,body:JSON.parse(options.body)}));},
  async readJsonResponse(response){if(response.error)throw new Error(response.error);return response;},
});
vm.runInContext(section('function filterAdminUserCards()', 'function renderAccountsModal()')+section('function applyAdminPermissionChoice(', 'async function setUserInsightsEnabled('),context);
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
