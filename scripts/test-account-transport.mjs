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
console.log('Browser account transport passed: one-time upgrade, no routine passwords/preflights/storage rewrites, cookie hints require validation, Creator target separation, coalesced foreground activity, recent-auth preservation, logout/expiry and stale-response/other-tab isolation.');
