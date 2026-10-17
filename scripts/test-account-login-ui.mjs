import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {runInNewContext} from 'node:vm';
import {installPasswordAutofillSubmit} from '../public/static/password-autofill-submit.js';
const app=await readFile('public/static/app.js','utf8'),html=await readFile('public/index.html','utf8');
const classes=(...initial)=>{const set=new Set(initial);return{contains:x=>set.has(x),add:x=>set.add(x),remove:x=>set.delete(x)};};
let mobile=true,focuses=0,resets=0;
const loginEmail={value:'',focus(){focuses++;}},loginPassword={value:''},loginForm={classList:classes()};
const entrancePage={classList:classes('hidden')};
const context={loginAutofillSubmit:{reset(){}},currentUserEmail:'',normalizeEmail:s=>String(s).trim().toLowerCase(),entrancePage,loginForm,loginEmail,loginPassword,
 setEntranceTab(){},forceConsoleSkin(){},clearLoginPasswordUpgrade(){resets++;},localStorage:{getItem(){return null;}},STAY_LOGGED_IN_PREFERENCE_KEY:'fixture',stayLoggedIn:{checked:false},
 document:{querySelector(){return null;}},loginTabButton:null,createTabButton:null,appShell:{classList:classes()},mobileActionBar:{classList:classes()},
 accountTransport:{options:async()=>({recovery:false})},setTimeout:fn=>fn(),isMobileLayout:()=>mobile};
const start=app.indexOf('function openLoginModal('),end=app.indexOf('function closeLoginModal(',start);
runInNewContext(app.slice(start,end)+';this.open=openLoginModal;',context);
context.open();assert.equal(resets,1);assert.equal(focuses,0,'Home Screen login must not force keyboard focus');
loginEmail.value='fixture@example.test';loginPassword.value='Synthetic autofill draft';
context.open();assert.equal(loginEmail.value,'fixture@example.test');assert.equal(loginPassword.value,'Synthetic autofill draft');assert.equal(resets,1);
context.open('fixture@example.test');assert.equal(loginPassword.value,'Synthetic autofill draft');
context.open('fixture@example.test',{resetPassword:true});assert.equal(loginPassword.value,'','rejected submissions still clear the password');
loginPassword.value='Another synthetic draft';context.open('other@example.test');assert.equal(loginEmail.value,'other@example.test');assert.equal(loginPassword.value,'');
entrancePage.classList.add('hidden');mobile=false;context.open('other@example.test');assert.equal(focuses,1,'desktop keeps initial username focus');
assert.match(html,/<input id="loginEmail"[^>]*autocomplete="username"/);
assert.match(html,/<input id="loginPassword"[^>]*name="password"[^>]*autocomplete="current-password"/);
assert.match(html,/<form id="loginForm"[^>]*method="post"/,'named credentials never use a default GET submission');
assert.match(html,/<input id="loginUpgradePassword"[^>]*disabled/);
// Exercise the real submit handler: OS submission plus a tap creates one login.
let handler,releaseLogin,logins=0;const submit={disabled:false};
const submitContext={installPasswordAutofillSubmit:()=>({reset(){}}),loginPasswordUpgradeField:null,loginForm:{addEventListener(_type,fn){handler=fn;},querySelector(){return submit;}},setEntranceTab(){},normalizeEmail:s=>s.trim().toLowerCase(),
 loginEmail:{value:'fixture@example.test'},loginPassword:{value:'Synthetic fixture'},stayLoggedIn:{checked:true},
 document:{querySelector(){return{classList:classes('hidden')};}},loginWithEmail:async()=>{logins++;await new Promise(r=>releaseLogin=r);}};
const submitStart=app.indexOf('let loginSubmissionPending = false;'),submitEnd=app.indexOf('createAccountForm.addEventListener',submitStart);
runInNewContext(app.slice(submitStart,submitEnd),submitContext);
const first=handler({preventDefault(){}});await handler({preventDefault(){}});assert.equal(logins,1);assert.equal(submit.disabled,true);
releaseLogin();await first;assert.equal(submit.disabled,false);
console.log('Login UI passed autofill draft preservation, explicit rejection/account-change clearing, mobile/desktop focus, inactive upgrade input and duplicate submit coalescing.');
{
 const listeners={},formListeners={};let task=null,submits=0,filled=false,valid=true,eligible=true,pending=false;
 const password={value:'Synthetic fixture',matches:()=>filled,addEventListener:(type,fn)=>listeners[type]=fn};
 const form={addEventListener:(type,fn)=>formListeners[type]=fn,checkValidity:()=>valid,requestSubmit(){submits++;formListeners.submit();}};
 const controller=installPasswordAutofillSubmit({form,password,eligible:()=>eligible,pending:()=>pending,schedule:fn=>(task=fn,1),cancel:()=>task=null});
 const flush=()=>{const fn=task;task=null;fn?.();};
 listeners.input({type:'input',inputType:'insertText'});flush();assert.equal(submits,0);
 listeners.change({type:'change'});flush();assert.equal(submits,0,'paste/manual change without autofill never signs in');
 filled=true;listeners.input({type:'input',inputType:'insertText'});assert.equal(task,null,'typing must not trigger even with stale autofill styling');
 listeners.animationstart({type:'animationstart',animationName:'login-password-autofilled'});flush();assert.equal(submits,1);
 listeners.change({type:'change'});flush();assert.equal(submits,1,'one automatic submission per presentation');
 controller.reset();listeners.change({type:'change'});formListeners.submit();flush();assert.equal(submits,1,'native Safari submit cancels the fallback');
 for(const state of ['hidden','pending','invalid']){
  controller.reset();eligible=state!=='hidden';pending=state==='pending';valid=state!=='invalid';
  listeners.change({type:'change'});flush();assert.equal(submits,1);
 }
 controller.reset();eligible=true;pending=false;valid=true;listeners.change({type:'change'});filled=false;flush();assert.equal(submits,1,'cleared credentials cannot submit a delayed callback');
 assert.match(app,/openLoginModal\("",\{resetPassword:true\}\)/,'deliberate logout explicitly clears the previous credentials');
}
{
 let handler,confirmed=false,requests=0;const actor='fixture@example.test',button={disabled:false,closest:()=>null};
 const context={accountsBody:{addEventListener:(_type,fn)=>handler=fn},authenticatedAccountEmail:()=>actor,
  showAppDialog:async options=>{assert.equal(options.requirePassword,undefined);return confirmed;},
  accountTransport:{signOutOtherDevices:async()=>{requests++;return{};}},readJsonResponse:async()=>({revokedSessions:1})};
 const start=app.indexOf("accountsBody.addEventListener('click',async event=>{"),end=app.indexOf('function normalizeAuthMessage',start);
 runInNewContext(app.slice(start,end),context);
 const event={target:{closest:selector=>selector==='[data-signout-other-devices]'?button:null}};
 await handler(event);assert.equal(requests,0,'cancel leaves every session untouched');
 confirmed=true;await handler(event);assert.equal(requests,1,'confirmation needs one revocation request and no password call');assert.equal(button.disabled,false);
}
// The legacy envelope must propagate provider password-upgrade signals to the
// outer login prompt instead of swallowing them and reopening a blank form.
const upgradeError=Object.assign(new Error('Synthetic upgrade required'),{code:'auth-password-upgrade-required'});
const restoreContext={currentUserEmail:'fixture@example.test',currentUserPassword:'Synthetic old password',OWNER_EMAIL:'owner@example.test',authUserEmail:'fixture@example.test',authUserPassword:'',
 normalizeEmail:s=>String(s||'').trim().toLowerCase(),setActiveCalendarContext(){},calendarTransitionStillCurrent:()=>true,cancelScheduledCloudStateSave(){},
 requestAccountLogin:async()=>{throw upgradeError;}};
const restoreStart=app.indexOf('async function restoreCloudState('),restoreEnd=app.indexOf('async function hydrateAuthenticatedWorkspace(',restoreStart);
runInNewContext(app.slice(restoreStart,restoreEnd)+';this.restore=restoreCloudState;',restoreContext);
await assert.rejects(()=>restoreContext.restore({mode:'login'}),error=>error===upgradeError);
assert.equal(restoreContext.currentUserEmail,'fixture@example.test');
console.log('Inline password-upgrade signals reach the login prompt without erasing its draft.');

// Recent-auth confirmation is a visible, separate username/current-password
// form; cancellation and submission always clear its transient values.
{
 const nodes=new Map();const field=()=>({value:'',disabled:true,required:false,focus(){},reportValidity(){return Boolean(this.value);}});
 const password=field(),username=field(),fields={hidden:true};
 nodes.set('#appDialogPassword',password);nodes.set('#appDialogUsername',username);nodes.set('#appDialogPasswordFields',fields);
 const classList=classes('hidden');classList.toggle=(value,force)=>force?classList.add(value):classList.remove(value);
 const dialog={classList,setAttribute(){}};
 const buttonClasses=classes();buttonClasses.toggle=(value,force)=>force?buttonClasses.add(value):buttonClasses.remove(value);
 const dialogContext={appDialog:dialog,document:{querySelector:key=>nodes.get(key)},appDialogTitle:{},appDialogMessage:{},appDialogConfirmButton:{classList:buttonClasses},appDialogCancelButton:{focus(){}},isMobileLayout:()=>true};
 const start=app.indexOf('let appDialogResolve = null;'),end=app.indexOf('appDialogCloseButton?.addEventListener',start);
 runInNewContext(app.slice(start,end)+';this.show=showAppDialog;this.close=closeAppDialog;',dialogContext);
 const plain=dialogContext.show({title:'Generic confirm',message:'Fixture',confirmLabel:'Confirm'});assert.equal(password.disabled,true);dialogContext.close(true);assert.equal(await plain,true);
 const cancelled=dialogContext.show({title:'Confirm other devices',message:'Fixture',confirmLabel:'Confirm sign out',requirePassword:true,email:'fixture@example.test'});
 assert.equal(password.disabled,false);assert.equal(username.value,'fixture@example.test');password.value='Transient fixture password';dialogContext.close(false);
 assert.equal(await cancelled,false);assert.equal(password.value,'');assert.equal(username.value,'');assert.equal(password.disabled,true);
 const submitted=dialogContext.show({title:'Confirm',message:'Fixture',confirmLabel:'Confirm sign out',requirePassword:true,email:'fixture@example.test'});
 dialogContext.close(true);assert.equal(dialog.classList.contains('hidden'),false,'empty password cannot confirm');password.value='Transient fixture password';dialogContext.close(true);
 assert.equal(await submitted,'Transient fixture password');assert.equal(password.value,'');assert.equal(password.disabled,true);
 assert.match(html,/<form id="appDialogForm" method="post">/);
 assert.match(html,/<input id="appDialogPassword"[^>]*autocomplete="current-password"[^>]*disabled/);
}
console.log('Recent-auth dialog passed explicit confirmation, empty-password refusal, generic-dialog compatibility and transient credential cleanup.');
// A post-auth render exception used to clear the hint while leaving the private
// UI visible, making status checks silently stop before reaching the server.
{
 let hint=true,workspaceVisible=false,forgotten=0;const noop=()=>{},field={value:'',classList:{contains:()=>true,remove(){}}};
 const context={performance:{now:()=>1},currentUserEmail:'fixture@example.test',OWNER_EMAIL:'fixture@example.test',OWNER_DOCTOR_KEY:'FIXTURE',
  beginFacilityOverviewAccountSession:noop,cancelScheduledCloudStateSave:noop,clearActiveViewedAccountState:noop,ensureLocalAccountLogin:noop,normalizeEmail:s=>String(s).toLowerCase(),setActiveCalendarContext:noop,primeInsightsAccessForCurrentView:noop,
  beginCalendarTransition:()=>({}),calendarTransitionStillCurrent:()=>true,localStorage:{setItem(){}},CURRENT_EMAIL_KEY:'email',STAY_LOGGED_IN_PREFERENCE_KEY:'remember',forceConsoleSkin:noop,setStatus:noop,setEntranceStatus:noop,
  accountCalendarContextForEmail:()=>({}),loadCachedCalendarSnapshotForContext:()=>null,restoreCloudState:async()=>({}),
  accountTransport:{expected:()=>hint,validated:()=>hint,email:()=>hint?'fixture@example.test':'',forget(){hint=false;forgotten++;}},loginPassword:{value:'Synthetic draft'},createPassword:{value:''},clearLoginPasswordUpgrade:noop,
  renderLoginState(){workspaceVisible=hint;},closeLoginModal(){workspaceVisible=true;},markLoginPhase:noop,creatorStartupContainmentActive:()=>true,
  finishContainedCreatorStartup(){throw Error('Synthetic local rendering failure');},document:{querySelector:()=>field},normalizeAuthMessage:s=>s};
 const start=app.indexOf('async function loginWithEmail('),end=app.indexOf('async function requestAccountLogin(',start);
 runInNewContext(app.slice(start,end)+';this.login=loginWithEmail;',context);
 await context.login('fixture@example.test','Synthetic password',{stayLoggedIn:true});assert.equal(hint,true);assert.equal(forgotten,0);assert.equal(workspaceVisible,true);
 context.restoreCloudState=async()=>{throw Error('Synthetic pre-auth rejection');};
 await context.login('fixture@example.test','Synthetic invalid password');assert.equal(hint,false);assert.equal(workspaceVisible,false);
}
{
 let forgotten=0,loggedOut=0,requests=0,creator=false,cloud=false;const context={
  isCreatorAuthenticated:()=>creator,get cloudAvailable(){return cloud;},document:{hidden:false},calendarStoreStatusRequest:null,
  accountTransport:{forget(){forgotten++;}},logoutCurrentUser:async options=>{assert.equal(options.expired,true);loggedOut++;},setEntranceStatus(){},
  performCalendarStoreStatusRefresh:async()=>{requests++;return{ok:true};}};
 const start=app.indexOf('async function refreshCalendarStoreStatus('),end=app.indexOf('async function performCalendarStoreStatusRefresh(',start);
 runInNewContext(app.slice(start,end)+';this.refresh=refreshCalendarStoreStatus;',context);
 await context.refresh({silent:true});assert.equal(loggedOut,0);await context.refresh({silent:false});assert.equal(forgotten,1);assert.equal(loggedOut,1);assert.equal(requests,0);
 creator=true;await context.refresh({silent:false});assert.equal(requests,1,'explicit status checks can recover from a transient cloud-availability flag');
}
console.log('Post-auth rendering failures preserve valid sessions; failed authentication hides the workspace; explicit status checks never silently ignore missing authority.');
