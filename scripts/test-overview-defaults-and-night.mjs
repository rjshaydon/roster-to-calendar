import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import vm from 'node:vm';
const app=await readFile(new URL('../public/static/app.js',import.meta.url),'utf8');
const section=(a,b)=>app.slice(app.indexOf(a),app.indexOf(b,app.indexOf(a)));
const normalize=v=>String(v||'').toUpperCase().trim();
const people=[{key:'ALICE',identity:'ALICE',displayName:'Alice User'},{key:'BOB',identity:'BOB',displayName:'Bob Colleague'}];
const state={togetherStaffKeys:[''],togetherPinnedDoctors:[],togetherUserClearedAll:false,togetherContext:null};
const defaults=vm.createContext({currentNonClinical:false,currentDefaultDoctorKey:'ALICE',currentRosterClaims:[{key:'ALICE',sourceType:'mmc'}],
 currentAccount:()=>({realName:'Alice User'}),selectedDoctor:()=>people[1],normalizeRosterName:normalize,rosterIdentityKey:normalize,doctorIdentityKey:d=>d.identity,
 facilityOverviewState:state,facilityOverviewTogetherStaffOptions:()=>state.togetherContext?people:[],facilityOverviewTogetherContextKey:()=> 'period',
 facilityOverviewTogetherDoctorKeys:d=>d?[d.key,...(d.aliases||[]).map(a=>a.key)]:[],
 facilityOverviewTogetherOptionFor:t=>state.togetherContext?people.find(p=>p.key===t.doctorKey):null,
 facilityOverviewTogetherFallbackOption:t=>({...t,key:t.doctorKey,identity:t.doctorKey}),canUseFacilityOverview:()=>true,canUseFullFacilityOverview:()=>true,
 resetFacilityOverviewScroll(){},renderFacilityOverview(){},loadFacilityOverviewTogether(){},});
vm.runInContext(section('function facilityOverviewTogetherViewer()', 'function facilityOverviewTogetherTermOptions()')+section('function openFacilityOverviewWorkingTogether(', 'function closeFacilityOverviewStaffActionMenu()'),defaults);
defaults.initializeFacilityOverviewTogetherState();
assert.deepEqual([...state.togetherStaffKeys],['ALICE',''],'own identity prepopulates before the network response');
state.togetherContext={key:'period'};defaults.initializeFacilityOverviewTogetherState();
assert.equal(state.togetherStaffKeys[0],'ALICE');
defaults.openFacilityOverviewWorkingTogether({doctorKey:'BOB',displayName:'Bob Colleague'});
assert.deepEqual([...state.togetherStaffKeys],['BOB','ALICE'],'clicked person first, signed-in user second, independent of selected calendar');
state.togetherStaffKeys=[''];state.togetherUserClearedAll=true;defaults.initializeFacilityOverviewTogetherState();
assert.equal(state.togetherStaffKeys[0],'','explicit all-staff search is not overwritten');
// A change in the context identity key must retain the clicked-person selection.
state.togetherPinnedDoctors=[{key:'BOB',identity:'OLD_BOB'}];state.togetherStaffKeys=['OLD_BOB','ALICE'];state.togetherUserClearedAll=false;
defaults.initializeFacilityOverviewTogetherState();assert.equal(state.togetherStaffKeys[0],'BOB');

const rank={SR:0,IR:1,JR:2,HMO:3,I:4};
const night=vm.createContext({normalizeWhoRole:v=>({ 'Senior Registrar':'SR','Junior Registrar':'JR',Intern:'I'}[v]||v),
 facilityOverviewEffectiveTeam:a=>a.team,renderFacilityOverviewGenericOnShiftPeriod:a=>a.length?'Other':'',escapeHtml:v=>v,
 clinicalSupportRosterMode:()=>'',clinicalSupportModeRank:()=>0,facilityOverviewAssignmentText:a=>a.rawValue||'',facilityOverviewOnShiftTimeLabel:()=>'',
 compareFacilityOverviewPeople:(a,b)=>(rank[a.seniority]-rank[b.seniority])||a.displayName.localeCompare(b.displayName),
 renderFacilityOverviewStaffName:p=>p.displayName,renderFacilityOverviewOnShiftSeniority:()=>'',renderFacilityOverviewContactAllocation:()=>'',});
vm.runInContext(section('function renderFacilityOverviewMmcNightPeriod(', 'function facilityOverviewAssignmentText(')+section('function renderFacilityOverviewStreamCard(', 'function facilityOverviewStandaloneServiceContacts(')+section('function renderFacilityOverviewOnShiftNames(', 'function clinicalSupportModeRank('),night);
const assignment=(name,role,team='Night',charge=false)=>({team,role,nightIcRank:charge?0:1,period:'Night',rawValue:team,person:{doctorKey:name,displayName:name,seniority:role}});
const html=night.renderFacilityOverviewMmcNightPeriod([
 assignment('Senior B','SR'),assignment('Junior IC','JR','Night',true),assignment('Senior A','SR'),
 assignment('Hub junior','HMO','NHJ'),assignment('SSU doctor','HMO','SSU'),assignment('Main HMO','HMO'),assignment('Main intern','I')]);
const labels=['Registrars','Night Hub','SSU team','HMOs','Interns'];
assert(labels.every((label,i)=>i===0||html.indexOf(labels[i-1])<html.indexOf(label)),'MMC panels follow the requested order');
assert(html.indexOf('Junior IC')<html.indexOf('Senior A')&&html.indexOf('Senior A')<html.indexOf('Senior B'),'in-charge registrar before other senior registrars');
assert(html.indexOf('Hub junior')>html.indexOf('Night Hub')&&html.indexOf('Hub junior')<html.indexOf('SSU team'),'NHJ appears in Night Hub');
const srCharge=night.renderFacilityOverviewMmcNightPeriod([assignment('Senior A','SR'),assignment('Senior B IC','SR','Night',true)]);
assert(srCharge.indexOf('Senior B IC')<srCharge.indexOf('Senior A'),'in-charge senior registrar wins alphabetical ordering');
console.log('Working together viewer defaults, clicked-person comparison, selection retention and MMC night grouping/order passed.');
