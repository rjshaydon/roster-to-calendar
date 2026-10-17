// Cookies carry authority. These storage records are hints for cold startup only.
export const COOKIE_AUTH_HINT = 'roster-cookie-auth';
export const AUTH_NOTICE = 'roster-auth-notice';
const RENEW_INTERVAL = 6 * 3600000;
export function createAccountTransport({ fetchImpl = (...args) => fetch(...args), persistentStorage, tabStorage, onAuthenticated = () => {}, onExpired = () => {}, now = () => Date.now() } = {}) {
  function readHint() { try { const x=JSON.parse(tabStorage?.getItem(COOKIE_AUTH_HINT)||persistentStorage?.getItem(COOKIE_AUTH_HINT)||'null');return x&&typeof x.email==='string'&&x.email.length<=254?x:null; } catch { return null; } }
  let hint=readHint(),validated=false,generation=0,authIntent=0,activityPromise=null,lastActivity=0,upgradeAvailable=Boolean(readHint()),cohortRollout=false,optionsPromise;
  const responseLeases=new WeakMap();
  function removeHints(){try{persistentStorage?.removeItem(COOKIE_AUTH_HINT);}catch{}try{tabStorage?.removeItem(COOKIE_AUTH_HINT);}catch{}}
  function notify(email){try{persistentStorage?.setItem(AUTH_NOTICE,JSON.stringify({email,at:now(),nonce:Math.random()}));}catch{}}
  function accept(email,persistent,fresh=false){
    if(validated&&hint?.email===email&&!fresh)return;
    const changed=hint?.email!==email||fresh;if(changed)generation++;
    hint={email,persistent:persistent===true};validated=true;removeHints();
    onAuthenticated(email);
    try{(hint.persistent?persistentStorage:tabStorage)?.setItem(COOKIE_AUTH_HINT,JSON.stringify(hint));}catch{}
    if(fresh){lastActivity=now();notify(email);}
  }
  function expire(){authIntent++;generation++;hint=null;validated=false;removeHints();onExpired();}
  function changedResponse(){return Response.json({error:'The account changed. Please try again.',code:'auth-session-changed'},{status:401});}
  async function auth(body){
    const lease=generation;
    const payload={...body};
    if(hint&&['status','activity','logout','changePassword','signOutOtherDevices','retryAuthenticationDeletion','reauthenticate'].includes(body.action))payload.expectedAccount=hint.email;
    const response=await fetchImpl('/api/auth/session',{method:'POST',credentials:'same-origin',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
    responseLeases.set(response,lease);
    if(lease===generation&&hint&&[401,403].includes(response.status)){
      const data=await response.clone().json().catch(()=>null);
      if(['auth-session-expired','auth-identity-mismatch'].includes(data?.code))expire();
    }
    return response;
  }
  async function discardAuthenticationResponse(response){
    // A browser applies Set-Cookie before JS can discard a stale response.
    // Revoke its account only when it differs from the latest selected account
    // (or sign-out won). The server expectation protects a newer shared cookie.
    if(response.ok){
      const data=await response.clone().json().catch(()=>null);
      if(typeof data?.email==='string'&&data.email.length<=254&&(!hint||hint.email!==data.email)){
        await fetchImpl('/api/auth/session',{method:'POST',credentials:'same-origin',headers:{'Content-Type':'application/json'},body:JSON.stringify({action:'logout',expectedAccount:data.email})}).catch(()=>{});
      }
    }
    return changedResponse();
  }
  async function explicitAuth(body,persistent){
    const intent=++authIntent,lease=++generation;
    let response=await auth(body);
    if(['signup','login'].includes(body.action)&&response.status===202){
      if(intent!==authIntent||lease!==generation)return changedResponse();
      const prepared=await response.clone().json().catch(()=>null);
      if(prepared?.code!=='auth-account-prepared')return Response.json({error:'Sign-up could not be confirmed. Please try again.',code:'auth-unavailable'},{status:503});
      if(intent!==authIntent||lease!==generation)return changedResponse();
      // One bounded continuation. The password stays in this pending call;
      // no cookie/hint is accepted until the provider verifies it on login.
      response=await auth({action:'login',email:body.email,password:body.newPassword||body.password,remember:body.remember===true});
    }
    if(intent!==authIntent||lease!==generation)return discardAuthenticationResponse(response);
    if(response.ok){const data=await response.clone().json();if(intent!==authIntent||lease!==generation)return discardAuthenticationResponse(response);if(response.status!==200||data?.ok!==true||typeof data.email!=='string'||!data.email)return Response.json({error:'Sign-in could not be confirmed.',code:'auth-unavailable'},{status:503});accept(data.email,persistent,true);}
    responseLeases.set(response,generation);return response;
  }
  const login=(email,password,remember,newPassword)=>explicitAuth({action:'login',email,password,remember:remember===true,...(newPassword?{newPassword}:{})},remember);
  const signup=body=>explicitAuth({action:'signup',email:body.email,password:body.password,realName:body.realName,remember:body.remember===true},body.remember);
  async function accountFetch(input,init={}){
    if(input!=='/api/state')return fetchImpl(input,init);
    let body;try{body=JSON.parse(init.body);}catch{return fetchImpl(input,init);}
    let lease=generation;
    if(upgradeAvailable&&body.action==='acceptInvite')return auth({action:'acceptInvite',inviteToken:body.inviteToken,newPassword:body.newPassword});
    if((hint||upgradeAvailable)&&body.action==='login'&&body.password&&body.mode==='create'){const r=await signup(body);if(!r.ok)return r;body={...body,mode:'login'};delete body.password;lease=generation;}
    // During a named canary, let the state route retain non-cohort legacy
    // sign-in. Its closed 409 bridge upgrades only eligible/retired accounts.
    if((hint||upgradeAvailable&&!cohortRollout)&&body.action==='login'&&body.password&&body.mode!=='create'){
      const response=await login(body.email,body.password,body.remember===true,body.newPassword);
      if(!response.ok)return response;lease=generation;
    }
    function send(){const payload={...body};if(hint){delete payload.password;delete payload.newPassword;delete payload.email;payload.expectedAccount=hint.email;}return fetchImpl(input,{...init,credentials:'same-origin',body:JSON.stringify(payload)});}
    let response=await send();
    if(body.action==='acceptInvite'&&response.status===409){const data=await response.clone().json().catch(()=>null);if(data?.code==='auth-upgrade-required'){upgradeAvailable=true;return auth({action:'acceptInvite',inviteToken:body.inviteToken,newPassword:body.newPassword});}}
    if(!hint&&body.action==='login'&&response.status===409){
      const data=await response.clone().json().catch(()=>null);
      if(data?.code==='auth-upgrade-required'&&body.mode==='create'&&body.password){upgradeAvailable=true;const r=await signup(body);if(!r.ok)return r;body={...body,mode:'login'};delete body.password;lease=generation;response=await send();}
      if(data?.code==='auth-upgrade-required'&&body.mode!=='create'&&body.password){
        upgradeAvailable=true;
        response=await login(body.email,body.password,body.remember===true,body.newPassword);
        if(!response.ok)return response;lease=generation;response=await send();
      }
    }
    if(lease!==generation)return changedResponse();
    if(response.ok&&response.headers.get('X-Roster-Authentication')==='cookie'){
      accept(hint?.email||body.email,hint?.persistent===true);
    }else if(hint&&[401,403].includes(response.status)){
      const data=await response.clone().json().catch(()=>null);
      if(['auth-session-expired','auth-identity-mismatch'].includes(data?.code))expire();
    }
    responseLeases.set(response,generation);
    return response;
  }
  return {
    fetch:accountFetch,
    forget(){authIntent++;generation++;hint=null;validated=false;removeHints();},
    responseCurrent:response=>!responseLeases.has(response)||responseLeases.get(response)===generation,
    options:()=>optionsPromise||=(async()=>{try{const r=await fetchImpl('/api/auth/session?options=1',{credentials:'same-origin'});const data=await r.json();upgradeAvailable=data.enabled===true;cohortRollout=data.cohortRollout===true;return data;}catch{return{enabled:false};}})(),
    requestRecovery:email=>auth({action:'requestRecovery',email}),
    completeEmail:body=>auth({action:'completeEmail',...body}),
    changePassword:(currentPassword,newPassword)=>explicitAuth({action:'changePassword',currentPassword,newPassword},hint?.persistent===true),
    signOutOtherDevices:()=>auth({action:'signOutOtherDevices'}),
    reauthenticate:password=>explicitAuth({action:'reauthenticate',email:hint?.email,password,remember:hint?.persistent===true},hint?.persistent===true),
    expected:()=>Boolean(hint),validated:()=>validated,email:()=>hint?.email||'',
    async logout(){authIntent++;generation++;if(hint){const r=await auth({action:'logout'});if(!r.ok)throw Error('Sign out could not be confirmed. Please try again.');}generation++;hint=null;validated=false;removeHints();notify('');},
    async noteActivity(){
      if(!validated||!hint||activityPromise||now()-lastActivity<RENEW_INTERVAL)return activityPromise;
      lastActivity=now();activityPromise=(async()=>{try{const r=await auth({action:'activity'});if(r.status===401)expire();else if(!r.ok)lastActivity=0;}catch{lastActivity=0;}finally{activityPromise=null;}})();return activityPromise;
    },
    notice(value){try{const data=JSON.parse(value);if(hint&&typeof data.email==='string'&&data.email!==hint.email)expire();}catch{}},
  };
}
