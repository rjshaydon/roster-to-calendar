// Cookies carry authority. These storage records are hints for cold startup only.
export const COOKIE_AUTH_HINT = 'roster-cookie-auth';
export const AUTH_NOTICE = 'roster-auth-notice';
const RENEW_INTERVAL = 6 * 3600000;
export function createAccountTransport({ fetchImpl = (...args) => fetch(...args), persistentStorage, tabStorage, onAuthenticated = () => {}, onExpired = () => {}, now = () => Date.now() } = {}) {
  function readHint() { try { const x=JSON.parse(tabStorage?.getItem(COOKIE_AUTH_HINT)||persistentStorage?.getItem(COOKIE_AUTH_HINT)||'null');return x&&typeof x.email==='string'&&x.email.length<=254?x:null; } catch { return null; } }
  let hint=readHint(),validated=false,generation=0,activityPromise=null,lastActivity=0,upgradeAvailable=Boolean(readHint()),optionsPromise;
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
  function expire(){generation++;hint=null;validated=false;removeHints();onExpired();}
  async function auth(body){return fetchImpl('/api/auth/session',{method:'POST',credentials:'same-origin',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});}
  async function login(email,password,remember){
    const response=await auth({action:'login',email,password,remember:remember===true});
    if(response.ok){const data=await response.clone().json();accept(data.email,remember,true);}
    return response;
  }
  async function signup(body){
    const response=await auth({action:'signup',email:body.email,password:body.password,realName:body.realName,remember:body.remember===true});
    if(response.ok){const data=await response.clone().json();accept(data.email,body.remember===true,true);}
    return response;
  }
  async function accountFetch(input,init={}){
    if(input!=='/api/state')return fetchImpl(input,init);
    let body;try{body=JSON.parse(init.body);}catch{return fetchImpl(input,init);}
    let lease=generation;
    if(upgradeAvailable&&body.action==='acceptInvite')return auth({action:'acceptInvite',inviteToken:body.inviteToken,newPassword:body.newPassword});
    if((hint||upgradeAvailable)&&body.action==='login'&&body.password&&body.mode==='create'){const r=await signup(body);if(!r.ok)return r;body={...body,mode:'login'};delete body.password;lease=generation;}
    if((hint||upgradeAvailable)&&body.action==='login'&&body.password&&body.mode!=='create'){
      const response=await login(body.email,body.password,body.remember===true);
      if(!response.ok)return response;lease=generation;
    }
    function send(){const payload={...body};if(hint){delete payload.password;delete payload.email;}return fetchImpl(input,{...init,credentials:'same-origin',body:JSON.stringify(payload)});}
    let response=await send();
    if(body.action==='acceptInvite'&&response.status===409){const data=await response.clone().json().catch(()=>null);if(data?.code==='auth-upgrade-required'){upgradeAvailable=true;return auth({action:'acceptInvite',inviteToken:body.inviteToken,newPassword:body.newPassword});}}
    if(!hint&&body.action==='login'&&response.status===409){
      const data=await response.clone().json().catch(()=>null);
      if(data?.code==='auth-upgrade-required'&&body.mode==='create'&&body.password){upgradeAvailable=true;const r=await signup(body);if(!r.ok)return r;body={...body,mode:'login'};delete body.password;lease=generation;response=await send();}
      if(data?.code==='auth-upgrade-required'&&body.mode!=='create'&&body.password){
        upgradeAvailable=true;
        response=await login(body.email,body.password,body.remember===true);
        if(!response.ok)return response;lease=generation;response=await send();
      }
    }
    if(lease!==generation)return Response.json({error:'The account changed. Please try again.',code:'auth-session-changed'},{status:401});
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
    forget(){generation++;hint=null;validated=false;removeHints();},
    responseCurrent:response=>!responseLeases.has(response)||responseLeases.get(response)===generation,
    options:()=>optionsPromise||=(async()=>{try{const r=await fetchImpl('/api/auth/session?options=1',{credentials:'same-origin'});const data=await r.json();upgradeAvailable=data.enabled===true;return data;}catch{return{enabled:false};}})(),
    requestRecovery:email=>auth({action:'requestRecovery',email}),
    completeEmail:body=>auth({action:'completeEmail',...body}),
    async changePassword(currentPassword,newPassword){const r=await auth({action:'changePassword',currentPassword,newPassword});if(r.ok){const data=await r.clone().json();accept(data.email,hint?.persistent===true,true);}return r;},
    signOutOtherDevices:()=>auth({action:'signOutOtherDevices'}),
    async reauthenticate(password){const r=await auth({action:'reauthenticate',email:hint?.email,password,remember:hint?.persistent===true});if(r.ok){const data=await r.clone().json();accept(data.email,hint?.persistent===true,true);}return r;},
    expected:()=>Boolean(hint),validated:()=>validated,email:()=>hint?.email||'',
    async logout(){if(hint){const r=await auth({action:'logout'});if(!r.ok)throw Error('Sign out could not be confirmed. Please try again.');}generation++;hint=null;validated=false;removeHints();notify('');},
    async noteActivity(){
      if(!validated||!hint||activityPromise||now()-lastActivity<RENEW_INTERVAL)return activityPromise;
      lastActivity=now();activityPromise=(async()=>{try{const r=await auth({action:'activity'});if(r.status===401)expire();else if(!r.ok)lastActivity=0;}catch{lastActivity=0;}finally{activityPromise=null;}})();return activityPromise;
    },
    notice(value){try{const data=JSON.parse(value);if(hint&&typeof data.email==='string'&&data.email!==hint.email)expire();}catch{}},
  };
}
