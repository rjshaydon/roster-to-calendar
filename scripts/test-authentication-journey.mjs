// Connected browser transport -> real API handlers -> SQLite, with a closed
// in-memory provider. No external network, staff credentials, emails or secrets.
import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {readFile,readdir} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {createAccountTransport,COOKIE_AUTH_HINT} from '../public/static/account-transport.js';
import {onRequestPost as authRoute} from '../functions/api/auth/session.js';
import {onRequestPost as stateRoute} from '../functions/api/state.js';
import {AUTH_PROJECT_ORIGIN} from '../functions/_lib/supabase-auth.js';
import {authenticationLiveProbe} from './security/authentication-live-probe-worker.mjs';

const sql=new DatabaseSync(':memory:');sql.exec('PRAGMA foreign_keys=ON');
for(const name of (await readdir(new URL('../migrations/',import.meta.url))).filter(n=>n.endsWith('.sql')).sort())sql.exec(await readFile(new URL('../migrations/'+name,import.meta.url),'utf8'));
const queries=[];
const db={prepare(q){return{args:[],bind(...a){this.args=a;return this;},async first(){queries.push(q);return sql.prepare(q).get(...this.args)||null;},async all(){queries.push(q);return{results:sql.prepare(q).all(...this.args)};},async run(){queries.push(q);const r=sql.prepare(q).run(...this.args);return{meta:{changes:Number(r.changes)}};}};},async batch(ss){sql.exec('BEGIN');try{const r=[];for(const s of ss)r.push(await s.run());sql.exec('COMMIT');return r;}catch(e){sql.exec('ROLLBACK');throw e;}}};
const config={enabled:true,sessionsEnabled:true,migrationEnabled:true,signupEnabled:true,emailDeliveryReady:false,url:AUTH_PROJECT_ORIGIN,publishableKey:'sb_publishable_fixture',secretKey:'sb_secret_fixture'};
const env={ROSTER_DB:db,ROSTER_AUTH_CONFIG:JSON.stringify(config),AUTH_LOGIN_LIMITER:{async limit(){return{success:true};}},ROSTER_FILES:{async get(){return null;}},ROSTER_CACHE:{async get(){return null;}},IDENTITY_DISCOVERY_ENABLED:'false'};
const users=new Map(),providerCalls=[];const oldFetch=globalThis.fetch;
globalThis.fetch=async(url,init)=>{
 assert.equal(new URL(url).origin,AUTH_PROJECT_ORIGIN,'no outbound destination except the mocked provider');
 const path=new URL(url).pathname,body=init.body?JSON.parse(init.body):{};providerCalls.push({path,method:init.method});
 if(path==='/auth/v1/admin/users'&&init.method==='POST'){
  assert.equal(body.email_confirm,true,'signup/migration preconfirm without sending mail');
  if(body.password.length<6)return Response.json({code:'weak_password'},{status:422});
  if([...users.values()].some(u=>u.email===body.email))return Response.json({code:'email_exists'},{status:422});
  const user={id:body.id,email:body.email,email_confirmed_at:new Date().toISOString()};users.set(user.id,{...user,password:body.password});return Response.json(user);
 }
 if(path.startsWith('/auth/v1/admin/users/')&&init.method==='GET'){
  const user=users.get(path.split('/').at(-1));return user?Response.json(user):Response.json({code:'missing'},{status:404});
 }
 if(path.startsWith('/auth/v1/admin/users/')&&init.method==='DELETE'){
  users.delete(path.split('/').at(-1));return Response.json({});
 }
 if(path==='/auth/v1/user'&&init.method==='PUT'){
  const subject=new Headers(init.headers).get('Authorization')?.replace('Bearer synthetic-access-','');
  const user=users.get(subject);assert.ok(user,'password update must carry the verified provider user token');
  user.password=body.password;return Response.json(user);
 }
 if(path==='/auth/v1/token'){
  const user=[...users.values()].find(u=>u.email===body.email&&u.password===body.password);
  return user?Response.json({user,access_token:'synthetic-access-'+user.id,refresh_token:'synthetic-refresh',expires_in:3600}):Response.json({code:'invalid'},{status:400});
 }
 throw Error('Unexpected mocked provider operation');
};
function storage(){const values=new Map();return{values,getItem:k=>values.get(k)||null,setItem:(k,v)=>values.set(k,String(v)),removeItem:k=>values.delete(k)};}
function browser(){
 let cookie='';const calls=[],persistent=storage(),tab=storage();
 const transport=createAccountTransport({persistentStorage:persistent,tabStorage:tab,fetchImpl:async(path,init={})=>{
  if(path.includes('?options'))return Response.json({enabled:true,recovery:false});
  const body=JSON.parse(init.body);calls.push({path,body});
  const request=new Request('https://app.example.test'+path,{...init,headers:{...init.headers,Origin:'https://app.example.test',...(cookie?{Cookie:cookie}:{})}});
  const context={request,env,data:{stateAction:body.action},waitUntil(){}};
  const response=await(path==='/api/state'?stateRoute(context):authRoute(context));
  if(response.headers.has('Set-Cookie'))cookie=response.headers.get('Set-Cookie').split(';')[0];
  return response;
 }});
 return{transport,calls,persistent,tab,get cookie(){return cookie;},send:body=>transport.fetch('/api/state',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)})};
}
try{
 const legacyEmail='legacy-journey@example.test',password='Legacy synthetic password',salt='c'.repeat(32),hash=createHash('sha256').update(salt+':'+password).digest('hex');
 const claim={email:legacyEmail,source_type:'DDH',doctor_key:'FIXTURE DOCTOR',display_name:'Fixture Doctor',matched_at:'2026-10-16',updated_at:'2026-10-16'};
 sql.prepare("INSERT INTO account_profiles(email,real_name,role,password_hash,password_salt,subscription_token) VALUES(?,'Fixture Doctor','user',?,?,'preserved-feed')").run(legacyEmail,hash,salt);
 sql.prepare('INSERT INTO account_claims VALUES(?,?,?,?,?,?)').run(...Object.values(claim));
 const b=browser();
 const bad=await b.send({action:'login',email:legacyEmail,password:'wrong'});assert.equal(bad.status,401);assert.equal(users.size,0);
 const login=await b.send({action:'login',email:legacyEmail,password,responseMode:'fast',remember:true});
 const data=await login.json();assert.equal(login.status,200);assert.equal(data.ok,true);assert.equal(data.realName,'Fixture Doctor');
 assert.equal(JSON.stringify(data).includes('synthetic-access'),false);assert.equal(b.transport.validated(),true);
 const record=sql.prepare('SELECT * FROM account_profiles WHERE email=?').get(legacyEmail);
 assert.equal(record.password_hash,'');assert.equal(record.password_salt,'');assert.equal(record.subscription_token,'preserved-feed');assert.deepEqual({...sql.prepare('SELECT * FROM account_claims WHERE email=?').get(legacyEmail)},claim);
 assert.equal(sql.prepare('SELECT phase FROM account_authentication_links WHERE email=?').get(legacyEmail).phase,'ready');
 const providerCount=providerCalls.length,beforeQueries=queries.length;
 assert.equal((await b.send({action:'loadAccountContext'})).status,200);
 assert.equal(providerCalls.length,providerCount,'ordinary account reads never contact the auth provider');
 assert.equal(queries.slice(beforeQueries).some(q=>/^\s*(UPDATE|INSERT|DELETE)/.test(q)),false,'ordinary reads never renew/write');
 assert.equal(b.calls.at(-1).body.password,undefined);assert.equal(b.calls.at(-1).body.email,undefined);
 const hint=b.persistent.getItem(COOKIE_AUTH_HINT);assert.equal(hint.includes(password),false);assert.equal(hint.includes('token'),false);
 // A paused migration still accepts the exact completed provider identity.
 env.ROSTER_AUTH_CONFIG=JSON.stringify({...config,migrationEnabled:false,migrationPaused:true,migrationEmails:[]});
 const returning=browser();assert.equal((await returning.send({action:'login',email:legacyEmail,password,responseMode:'fast'})).status,200);
 await returning.transport.logout();assert.equal(returning.transport.expected(),false);
 env.ROSTER_AUTH_CONFIG=JSON.stringify(config);
 const signup=browser(),signupEmail='new-journey@example.test';
 const signed=await signup.send({action:'login',mode:'create',email:signupEmail,password:'New synthetic password',realName:'New Person',responseMode:'fast'});
 const signedData=await signed.json();assert.equal(signed.status,200);assert.equal(signedData.realName,'New Person');assert.equal(signup.transport.validated(),true);
 assert.equal(signup.calls.filter(c=>c.path==='/api/auth/session'&&c.body.action==='signup').length,1);
 assert.equal(signup.calls.filter(c=>c.path==='/api/auth/session'&&c.body.action==='login').length,0,'signup requires no second password login');
 assert.equal(sql.prepare('SELECT password_hash FROM account_profiles WHERE email=?').get(signupEmail).password_hash,'');
 const stolenCookie=signup.cookie;await signup.transport.logout();
 const replay=new Request('https://app.example.test/api/state',{method:'POST',headers:{Origin:'https://app.example.test','Content-Type':'application/json',Cookie:stolenCookie},body:JSON.stringify({action:'loadAccountContext'})});
 assert.equal((await stateRoute({request:replay,env,data:{stateAction:'loadAccountContext'}})).status,401);
 // Only incompatible legacy passwords need an inline upgrade. The old proof
 // must be valid, and only the replacement reaches provider credential storage.
 for(const [i,legacyPassword] of ['four','🔐'.repeat(19)].entries()){
  const upgradeEmail='upgrade-'+i+'@example.test',upgradeSalt='e'.repeat(32);
  const oldHash=createHash('sha256').update(upgradeSalt+':'+legacyPassword).digest('hex');
  sql.prepare("INSERT INTO account_profiles(email,real_name,role,password_hash,password_salt,subscription_token) VALUES(?,'Upgrade Person','user',?,?,'upgrade-feed')").run(upgradeEmail,oldHash,upgradeSalt);
  const upgrade=browser();const needed=await upgrade.send({action:'login',email:upgradeEmail,password:legacyPassword,responseMode:'fast'});
  assert.equal(needed.status,409);assert.equal((await needed.json()).code,'auth-password-upgrade-required');
  assert.equal(sql.prepare('SELECT password_hash FROM account_profiles WHERE email=?').get(upgradeEmail).password_hash,oldHash);
  const upgraded=await upgrade.send({action:'login',email:upgradeEmail,password:legacyPassword,newPassword:'Stronger replacement password',responseMode:'fast'});
  assert.equal(upgraded.status,200);assert.equal(upgrade.transport.validated(),true);
  assert.equal(sql.prepare('SELECT password_hash FROM account_profiles WHERE email=?').get(upgradeEmail).password_hash,'');
  assert.equal(upgrade.calls.at(-1).body.newPassword,undefined,'replacement is not repeated in state/calendar bodies');
 }
 assert.equal(providerCalls.some(c=>c.path.includes('generate_link')||c.path.includes('verify')),false,'no confirmation or recovery operations on signup/migration');
 // Exercise the deployable probe with the same real route/schema adapter before
 // exposing it to live provider credentials. Only its provider remains mocked.
 const gate='f'.repeat(64),probe=authenticationLiveProbe(gate,'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee');
 const probeEnv={TEST_DB:db,ROSTER_AUTH_CONFIG:JSON.stringify({...config,enabled:false})};
 async function probeStep(step,kind='new',cookie=''){
  return probe.fetch(new Request('https://probe.example.test/'+step+'?kind='+kind,{method:'POST',headers:{Authorization:'Bearer '+gate,...(cookie?{Cookie:cookie}:{})}}),probeEnv);
 }
 assert.equal((await probe.fetch(new Request('https://probe.example.test/signup',{method:'POST'}),probeEnv)).status,404);
 assert.equal((await probeStep('seed','legacy')).status,200);
 for(const [step,kind] of [['signup','new'],['login','legacy']]){
  const response=await probeStep(step,kind);assert.equal(response.status,200);
  const cookie=response.headers.get('Set-Cookie').split(';')[0],result=await response.json();
  assert.equal(result.legacyHashRetired,true);assert.equal(result.identityReady,true);assert.equal(result.legacyFeedPreserved,true);
  assert.equal((await probeStep('context',kind,cookie)).status,200);
  if(kind==='legacy'){
   assert.equal((await probeStep('change',kind,cookie)).status,200);
   assert.equal((await probeStep('context',kind,cookie)).status,401);
   assert.equal((await probeStep('oldPassword',kind)).status,401);
   assert.equal((await probeStep('returning',kind)).status,200);
  }
 }
 const cleanup=await probeStep('cleanup');assert.equal(cleanup.status,200);assert.equal((await cleanup.json()).syntheticAccountsRemoved,2);
 console.log('Connected auth journeys passed: password login -> migration -> cookie -> account/calendar envelope; no-mail signup -> cookie -> envelope; unchanged reads without provider work/writes; migration pause preserves ready users; logout replay denied. No external services.');
}finally{globalThis.fetch=oldFetch;sql.close();}
