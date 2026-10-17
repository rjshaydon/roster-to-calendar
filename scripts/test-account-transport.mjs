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
