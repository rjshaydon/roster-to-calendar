import assert from 'node:assert/strict';
import { createAccountTransport, COOKIE_AUTH_HINT } from '../public/static/account-transport.js';
const store=()=>{const m=new Map();return{getItem:k=>m.get(k)||null,setItem:(k,v)=>m.set(k,String(v)),removeItem:k=>m.delete(k)};};
const persistent=store(),tab=store(),calls=[];let accepted=0,expired=0,clock=1_000_000_000,active=false,release;
const fetchImpl=async(url,init)=>{
 const body=JSON.parse(init.body);calls.push({url,body});
 if(url==='/api/auth/session'){
  if(body.action==='login'){active=true;return Response.json({ok:true,email:body.email});}
  return Response.json({ok:true});
 }
 if(body.action==='late')return new Promise(resolve=>{release=resolve;});
 if(body.action==='expire')return Response.json({code:'auth-session-expired'},{status:401});
 if(body.action==='recent')return Response.json({code:'auth-reauthentication-required'},{status:401});
 if(!active)return Response.json({code:'auth-upgrade-required'},{status:409});
 assert.equal(body.password,undefined);assert.equal(body.email,undefined);
 return Response.json({ok:true},{headers:{'X-Roster-Authentication':'cookie'}});
};
const t=createAccountTransport({fetchImpl,persistentStorage:persistent,tabStorage:tab,onAuthenticated(){accepted++;},onExpired(){expired++;},now:()=>clock});
const send=body=>t.fetch('/api/state',{method:'POST',body:JSON.stringify(body)});
assert.equal(t.expected(),false);
assert.equal((await send({action:'login',email:'user@example.test',password:'fixture',remember:true})).status,200);
assert.equal(calls.length,3);assert.equal(accepted,1);assert.equal(t.validated(),true);
assert.equal(JSON.stringify([...Object.values(persistent)]).includes('fixture'),false);
const savedHint=persistent.getItem(COOKIE_AUTH_HINT);assert.equal(savedHint.includes('fixture'),false);assert.equal(savedHint.includes('token'),false);
assert.equal(calls[1].body.password,'fixture','password is used only for explicit authentication');
const before=calls.length;for(let i=0;i<100;i++)await send({action:'query',email:'target@example.test',password:'should-never-leave-browser',targetEmail:'reviewed@example.test'});
assert.equal(calls.length,before+100,'normal calls gain no preflight/status/polling requests');assert.equal(accepted,1,'unchanged authenticated responses do not repeat cleanup/storage work');
assert.equal(calls.at(-1).body.targetEmail,'reviewed@example.test','viewed target is distinct from session actor');
assert.equal((await send({action:'recent'})).status,401);assert.equal(expired,0,'recent authentication does not wipe the workspace');
await t.noteActivity();assert.equal(calls.length,before+101,'fresh authentication does not immediately renew');
clock+=6*3600000;await Promise.all(Array.from({length:10},()=>t.noteActivity()));
assert.equal(calls.filter(x=>x.body.action==='activity').length,1,'foreground activity is coalesced');
const late=send({action:'late'});await t.logout();release(Response.json({private:'old-account'}));assert.equal((await late).status,401,'late private responses cannot repaint a signed-out workspace');
assert.equal(persistent.getItem(COOKIE_AUTH_HINT),null);
// A cold hint is never considered validated. Its first state request verifies it.
persistent.setItem(COOKIE_AUTH_HINT,savedHint);const cold=createAccountTransport({fetchImpl,persistentStorage:persistent,tabStorage:tab,onExpired(){expired++;},now:()=>clock});
assert.equal(cold.expected(),true);assert.equal(cold.validated(),false);
await cold.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'login',email:'user@example.test'})});assert.equal(cold.validated(),true);
const beforeExpiry=calls.length;await cold.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'expire',password:'legacy-must-not-be-retried'})});assert.equal(calls.length,beforeExpiry+1);assert.equal(cold.expected(),false);assert.equal(expired,1);
const tabOnly=store();tabOnly.setItem(COOKIE_AUTH_HINT,JSON.stringify({email:'user@example.test',persistent:false}));const peer=createAccountTransport({fetchImpl,persistentStorage:store(),tabStorage:tabOnly,onExpired(){expired++;}});peer.notice(JSON.stringify({email:'different@example.test'}));assert.equal(peer.expected(),false);
// Both a cold old client and a client that already knows cookie mode must
// create once, accept its cookie and fetch the normal calendar envelope.
for(const informed of [false,true]){
 const sequence=[];let signed=false;
 const transport=createAccountTransport({persistentStorage:store(),tabStorage:store(),fetchImpl:async(url,init)=>{
   if(url.includes('?options'))return Response.json({enabled:true});
   const b=JSON.parse(init.body);sequence.push({url,body:b});
   if(url==='/api/auth/session'){assert.equal(b.action,'signup');signed=true;return Response.json({ok:true,email:b.email});}
   if(!signed)return Response.json({code:'auth-upgrade-required'},{status:409});
   assert.equal(b.mode,'login');assert.equal(b.password,undefined);assert.equal(b.email,undefined);
   return Response.json({ok:true,state:{calendar:'fixture'}},{headers:{'X-Roster-Authentication':'cookie'}});
 }});
 if(informed)await transport.options();
 const response=await transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'login',mode:'create',email:'new@example.test',password:'Strong fixture only',realName:'New Person',remember:true})});
 assert.equal((await response.json()).state.calendar,'fixture');assert.equal(transport.validated(),true);
 assert.equal(sequence.filter(x=>x.url==='/api/auth/session').length,1,'no extra login after signup');
}
// Explicit auth completions must not overwrite a newer selected account.
{
 const pending=new Map(),signals=new Map(),stateRequests=[],cleanup=[];
 const arrived=email=>new Promise(r=>signals.set(email,r));
 const transport=createAccountTransport({persistentStorage:store(),tabStorage:store(),fetchImpl:async(url,init={})=>{
  if(url.includes('?options'))return Response.json({enabled:true});
  const b=JSON.parse(init.body);
  if(url==='/api/auth/session'&&b.action==='login'){signals.get(b.email)?.();return new Promise(r=>pending.set(b.email,r));}
  if(url==='/api/auth/session'){cleanup.push(b);return Response.json({ok:true});}
  stateRequests.push(b);return Response.json({ok:true},{headers:{'X-Roster-Authentication':'cookie'}});
 }});
 await transport.options();
 const firstArrival=arrived('first@example.test');
 const first=transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'login',email:'first@example.test',password:'fixture'})});await firstArrival;
 const secondArrival=arrived('second@example.test');
 const second=transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'login',email:'second@example.test',password:'fixture'})});await secondArrival;
 pending.get('second@example.test')(Response.json({ok:true,email:'second@example.test'}));assert.equal((await second).status,200);
 pending.get('first@example.test')(Response.json({ok:true,email:'first@example.test'}));assert.equal((await first).status,401);
 assert.equal(transport.email(),'second@example.test');assert.equal(stateRequests.length,1);
 assert.equal(stateRequests[0].expectedAccount,'second@example.test');assert.deepEqual(cleanup,[{action:'logout',expectedAccount:'first@example.test'}],'stale Set-Cookie is cleaned up only under its exact account expectation');
 // Logout during another in-flight authentication cannot restore a local hint.
 const thirdArrival=arrived('third@example.test');
 const third=transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'login',email:'third@example.test',password:'fixture'})});await thirdArrival;
 await transport.logout();pending.get('third@example.test')(Response.json({ok:true,email:'third@example.test'}));
 assert.equal((await third).status,401);assert.equal(transport.expected(),false);
 assert.equal(cleanup.at(-1).expectedAccount,'third@example.test');
}
// A prepared account is not an authenticated session, and cancelling its
// pending request must prevent the automatic verification from being sent.
{
 let releasePrepare,prepareReached;const reached=new Promise(r=>prepareReached=r),requests=[];
 const transport=createAccountTransport({persistentStorage:store(),tabStorage:store(),fetchImpl:async(url,init)=>{
  if(url.includes('?options'))return Response.json({enabled:true});
  const body=JSON.parse(init.body);requests.push(body.action);
  if(body.action==='signup'){prepareReached();return new Promise(r=>releasePrepare=r);}
  throw Error('Cancelled preparation must not continue to login');
 }});
 await transport.options();
 const pending=transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'login',mode:'create',email:'prepared@example.test',password:'Strong fixture password',realName:'Prepared Person'})});
 await reached;await transport.logout();releasePrepare(Response.json({ok:false,code:'auth-account-prepared'},{status:202}));
 assert.equal((await pending).status,401);assert.equal(transport.expected(),false);assert.deepEqual(requests,['signup']);
}
console.log('Browser account transport passed: one-time upgrade, bounded signup verification, no routine passwords/preflights/storage rewrites, cookie hints require validation, Creator target separation, coalesced foreground activity, recent-auth preservation, logout/expiry and stale-response/other-tab isolation.');

// Local authority disappears immediately, while a new login waits for the old
// response's Set-Cookie to settle. No new cookie can be cleared by late logout.
{
 const storage=store();storage.setItem(COOKIE_AUTH_HINT,JSON.stringify({email:'old@example.test',persistent:true}));
 let releaseLogout;const requests=[];
 const transport=createAccountTransport({persistentStorage:storage,tabStorage:store(),fetchImpl:async(_url,init)=>{
  const body=JSON.parse(init.body);requests.push(body);
  if(body.action==='logout')return new Promise(resolve=>{releaseLogout=resolve;});
  return Response.json({ok:true,email:body.email});
 }});
 const outgoing=transport.logout();assert.equal(transport.expected(),false);assert.equal(storage.getItem(COOKIE_AUTH_HINT),null);
 const incoming=transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'login',email:'new@example.test',password:'fixture password',remember:true})});assert.equal(requests.length,1,'new login waits for pending cookie clearance');
 releaseLogout(Response.json({ok:true}));await outgoing;assert.equal((await incoming).status,200);assert.equal(transport.email(),'new@example.test');
 assert.equal(requests[0].expectedAccount,'old@example.test');
}
{
 const storage=store();storage.setItem(COOKIE_AUTH_HINT,JSON.stringify({email:'fixture@example.test'}));let checks=0,expired=0;let time=100000;
 const transport=createAccountTransport({persistentStorage:storage,tabStorage:store(),now:()=>time,onExpired(){expired++;},fetchImpl:async(_url,init)=>{
  const body=JSON.parse(init.body);checks++;return body.action==='status'&&checks>1?Response.json({code:'auth-session-expired'},{status:401}):Response.json({ok:true},{headers:{'X-Roster-Authentication':'cookie'}});
 }});
 await transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'loadAccountContext'})});
 await Promise.all([transport.checkSession(),transport.checkSession()]);assert.equal(checks,2);assert.equal(expired,1);assert.equal(transport.expected(),false);
 await transport.checkSession();assert.equal(checks,2,'signed-out tabs never poll');
}
console.log('Immediate local logout, serialized cookie clearance/new login and coalesced foreground revocation checks passed.');

// Activity denial has a single owner; one revoked response opens login once.
{
 const storage=store();storage.setItem(COOKIE_AUTH_HINT,JSON.stringify({email:'fixture@example.test'}));let expired=0;
 const transport=createAccountTransport({persistentStorage:storage,tabStorage:store(),now:()=>1e9,onExpired(){expired++;},fetchImpl:async(_url,init)=>JSON.parse(init.body).action==='activity'?Response.json({code:'auth-session-expired'},{status:401}):Response.json({ok:true},{headers:{'X-Roster-Authentication':'cookie'}})});
 await transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'loadAccountContext'})});await transport.noteActivity();assert.equal(expired,1);
}
// A prior session's denial cannot sign out a successful newer login.
{
 const storage=store();storage.setItem(COOKIE_AUTH_HINT,JSON.stringify({email:'old@example.test'}));let releaseActivity,expired=0;
 const transport=createAccountTransport({persistentStorage:storage,tabStorage:store(),now:()=>1e9,onExpired(){expired++;},fetchImpl:async(url,init)=>{
  const b=JSON.parse(init.body);if(b.action==='activity')return new Promise(r=>releaseActivity=r);
  return url==='/api/auth/session'?Response.json({ok:true,email:b.email}):Response.json({ok:true},{headers:{'X-Roster-Authentication':'cookie'}});
 }});
 await transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'loadAccountContext'})});const old=transport.noteActivity();await transport.logout();
 await transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'login',email:'new@example.test',password:'Synthetic fixture only'})});
 releaseActivity(Response.json({code:'auth-session-expired'},{status:401}));await old;assert.equal(transport.email(),'new@example.test');assert.equal(expired,0);
}
// Coalesce focus bursts, but never drop a real background/foreground return.
{
 const storage=store();storage.setItem(COOKIE_AUTH_HINT,JSON.stringify({email:'fixture@example.test'}));let clock=1e6,checks=0,revoked=false,expired=0;
 const transport=createAccountTransport({persistentStorage:storage,tabStorage:store(),now:()=>clock,onExpired(){expired++;},fetchImpl:async(_url,init)=>{
  if(JSON.parse(init.body).action==='status'){checks++;return revoked?Response.json({code:'auth-session-expired'},{status:401}):Response.json({ok:true});}
  return Response.json({ok:true},{headers:{'X-Roster-Authentication':'cookie'}});
 }});
 await transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'loadAccountContext'})});await transport.checkSession();clock+=5000;
 await transport.checkSession();assert.equal(checks,1);revoked=true;
 await Promise.all([transport.checkSession({resume:true}),transport.checkSession({resume:true}),transport.checkSession()]);assert.equal(checks,2);assert.equal(expired,1);
}
// Response decoding is asynchronous too: recheck generation after JSON decoding.
for(const endpoint of ['/api/auth/session','/api/state']){
 const storage=store();storage.setItem(COOKIE_AUTH_HINT,JSON.stringify({email:'fixture@example.test'}));let releaseJson,jsonStarted,expired=0;
 const started=new Promise(r=>jsonStarted=r);let deny=false;
 const transport=createAccountTransport({persistentStorage:storage,tabStorage:store(),now:()=>1e9,onExpired(){expired++;},fetchImpl:async(url,init)=>{
  const b=JSON.parse(init.body);
  if(deny&&url===endpoint&&(b.action==='status'||b.action==='oldRead')){
   const response=Response.json({},{status:401});response.clone=()=>({json(){jsonStarted();return new Promise(r=>releaseJson=r);}});return response;
  }
  return url==='/api/auth/session'?Response.json({ok:true,email:b.email}):Response.json({ok:true},{headers:{'X-Roster-Authentication':'cookie'}});
 }});
 await transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'loadAccountContext'})});deny=true;
 const old=endpoint==='/api/state'?transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'oldRead'})}):transport.checkSession();await started;
 await transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'login',email:'fixture@example.test',password:'Synthetic fixture only'})});
 releaseJson({code:'auth-session-expired'});await old;assert.equal(transport.email(),'fixture@example.test');assert.equal(expired,0);
}
// Passwords/Face ID returning during explicit sign-in must not validate its old cookie.
{
 const storage=store();storage.setItem(COOKIE_AUTH_HINT,JSON.stringify({email:'fixture@example.test'}));let releaseLogin;const calls=[];
 const transport=createAccountTransport({persistentStorage:storage,tabStorage:store(),now:()=>1e9,fetchImpl:async(url,init)=>{
  const b=JSON.parse(init.body);calls.push(b.action);if(url==='/api/auth/session'&&b.action==='login')return new Promise(r=>releaseLogin=r);
  return Response.json({ok:true},{headers:{'X-Roster-Authentication':'cookie'}});
 }});
 await transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'loadAccountContext'})});
 const login=transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'login',email:'fixture@example.test',password:'Synthetic fixture only'})});
 await transport.checkSession({resume:true});await transport.noteActivity();assert.deepEqual(calls,['loadAccountContext','login']);
 releaseLogin(Response.json({ok:true,email:'fixture@example.test'}));await login;assert.equal(transport.validated(),true);
}
console.log('Home Screen session regressions passed single expiry, old-response/JSON-decoding fences, genuine resume checks and suppression during explicit authentication.');
// iOS may suspend a response obtained before backgrounding. A resume must queue
// one fresh check after that old response, with no ongoing polling afterward.
{
 const storage=store();storage.setItem(COOKIE_AUTH_HINT,JSON.stringify({email:'fixture@example.test'}));let releaseStatus,checks=0,expired=0;
 const transport=createAccountTransport({persistentStorage:storage,tabStorage:store(),now:()=>1e9,onExpired(){expired++;},fetchImpl:async(_url,init)=>{
  if(JSON.parse(init.body).action==='status'){
   checks++;if(checks===1)return new Promise(r=>releaseStatus=r);
   return Response.json({code:'auth-session-expired'},{status:401});
  }
  return Response.json({ok:true},{headers:{'X-Roster-Authentication':'cookie'}});
 }});
 await transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'loadAccountContext'})});
 const beforeBackground=transport.checkSession();transport.checkSession({resume:true});transport.checkSession({resume:true});
 releaseStatus(Response.json({ok:true}));await beforeBackground;assert.equal(checks,2);assert.equal(expired,1);assert.equal(transport.expected(),false);
}
console.log('Suspended pre-background checks receive one fresh, coalesced resume validation.');
// Requests share the authenticated generation and existing foreground check.
{
 const calls=[],summaries=[];let release,signal;const transport=createAccountTransport({persistentStorage:store(),tabStorage:store(),onSessionStatus:data=>summaries.push(data),fetchImpl:async(url,init)=>{
  const body=JSON.parse(init.body);calls.push({url,body});
  if(body.action==='heldRequest')return new Promise(r=>release=r);
  if(body.action==='status')return Response.json({ok:true,requests:{summary:{version:2,unread:1}}});
  return Response.json({ok:true,email:'requester@example.test'},{headers:{'X-Roster-Authentication':'cookie'}});
 }});
 await transport.fetch('/api/state',{method:'POST',body:JSON.stringify({action:'login',email:'requester@example.test',password:'synthetic'})});
 await transport.requests('summary',{expectedAccount:'other@example.test'});assert.equal(calls.at(-1).url,'/api/requests');assert.equal(calls.at(-1).body.expectedAccount,'requester@example.test');assert.equal(calls.at(-1).body.password,undefined);
 const before=calls.length;await Promise.all([transport.checkSession({resume:true}),transport.checkSession()]);assert.equal(calls.length,before+1);assert.equal(summaries[0].requests.summary.unread,1);
 const held=transport.requests('heldRequest');await Promise.resolve();await transport.logout();release(Response.json({ok:true,body:'private old request'}));assert.equal(transport.responseCurrent(await held),false);
 console.log('Requests transport passed fixed endpoint/actor binding, password-free calls, one coalesced foreground summary and stale private-response rejection.');
}
