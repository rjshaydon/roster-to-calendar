import {loadAccountMirror} from '../functions/_lib/d1-calendar.js';
import assert from 'node:assert/strict';
import { DatabaseSync } from 'node:sqlite';
import { readFile, readdir } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { SESSION_COOKIE, SESSION_IDLE_MS, SESSION_OUTER_MS, SESSION_RENEW_MS, accountSessionsEnabled, issueAccountSession, readAccountSession, renewAccountSession, revokeAccountSession, revokeOtherAccountSessions, requireRecentAuthentication, requireSameOrigin, cleanExpiredAccountSessions } from '../functions/_lib/account-sessions.js';
import { onRequestPost as sessionRoute } from '../functions/api/auth/session.js';
import { onRequestPost as stateRoute, prepareAccountResponse, prepareAccountContextResponse } from '../functions/api/state.js';
import { AUTH_PROJECT_ORIGIN } from '../functions/_lib/supabase-auth.js';

const sqlite=new DatabaseSync(':memory:'); sqlite.exec('PRAGMA foreign_keys=ON');
for(const name of (await readdir(new URL('../migrations/',import.meta.url))).filter(x=>x.endsWith('.sql')).sort()) sqlite.exec(await readFile(new URL('../migrations/'+name,import.meta.url),'utf8'));
const queries=[]; const db={prepare(sql){return{args:[],bind(...args){this.args=args;return this;},async first(){queries.push({sql,write:false});return sqlite.prepare(sql).get(...this.args)||null;},async all(){queries.push({sql,write:false});return{results:sqlite.prepare(sql).all(...this.args)};},async run(){queries.push({sql,write:true});const result=sqlite.prepare(sql).run(...this.args);return{meta:{changes:Number(result.changes)}};}};},async batch(statements){sqlite.exec('BEGIN');try{const results=[];for(const stmt of statements)results.push(await stmt.run());sqlite.exec('COMMIT');return results;}catch(e){sqlite.exec('ROLLBACK');throw e;}}};
const email='person@example.test',other='other@example.test',subject='11111111-2222-4333-8444-555555555555';
const now=Date.now();
for(const [e,id] of [[email,subject],[other,'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee']]){
 sqlite.prepare("INSERT INTO account_profiles(email,real_name,role,password_hash,password_salt,subscription_token) VALUES(?,?,'user','','',?)").run(e,'Fixture Person','feed-'+e);
 sqlite.prepare("INSERT INTO account_authentication_links VALUES(?,'supabase',?,'ready','revision','now','now')").run(e,id);
}
const env={ROSTER_DB:db,AUTH_LOGIN_LIMITER:{async limit(){return{success:true};}},ROSTER_AUTH_CONFIG:JSON.stringify({enabled:true,sessionsEnabled:true,url:AUTH_PROJECT_ORIGIN,publishableKey:'sb_publishable_fixture',secretKey:'sb_secret_fixture'})};
function request(cookie,body={action:'status'},origin='https://app.example.test') {return new Request('https://app.example.test/api/auth/session',{method:'POST',headers:{'Content-Type':'application/json',Origin:origin,'Sec-Fetch-Site':'same-origin',...(cookie?{Cookie:cookie.split(';')[0]}:{})},body:JSON.stringify(body)});}
function context(req,environment=env){return{request:req,env:environment,data:{},waitUntil(){}};}
assert.equal(accountSessionsEnabled({}),false);
assert.equal(accountSessionsEnabled({ROSTER_AUTH_CONFIG:'{'}),false);
assert.equal((await sessionRoute(context(request(null),{ROSTER_DB:{prepare(){throw Error('disabled must not read');}}}))).status,503);
// Pages reaches the native limiter through a private service binding. Refusal
// and unavailable admission must precede account/provider work.
const admissionBody={action:'login',email,password:'fixture-password'};
for(const [status,expected] of [[429,429],[503,503]]){
  const before=queries.length;
  const binding={async fetch(url,options){
    assert.equal(url,'https://admission.internal/check');
    assert.match(options.body,/^login:[a-f0-9]{64}$/);
    assert.equal(options.body.includes(email),false);
    assert.equal(options.body.includes(admissionBody.password),false);
    return new Response(null,{status});
  }};
  assert.equal((await sessionRoute(context(request(null,admissionBody),{...env,AUTH_LOGIN_LIMITER:binding}))).status,expected);
  assert.equal(queries.length,before,'refused admission does not read D1');
}
const beforeMissingAdmission=queries.length;
assert.equal((await sessionRoute(context(request(null,admissionBody),{...env,AUTH_LOGIN_LIMITER:undefined}))).status,503);
assert.equal(queries.length,beforeMissingAdmission);
// A named rollout keeps non-cohort existing users on their current path.
// No malformed cookie or new-account creation can select that compatibility path.
const legacyEmail='legacy@example.test',legacySalt='b'.repeat(32),legacyPassword='Legacy fixture password';
const legacyHash=createHash('sha256').update(legacySalt+':'+legacyPassword).digest('hex');
sqlite.prepare("INSERT INTO account_profiles(email,real_name,role,password_hash,password_salt,subscription_token) VALUES(?,'Legacy Person','user',?,?,'unchanged-legacy-feed')").run(legacyEmail,legacyHash,legacySalt);
const rolloutEnv={...env,ROSTER_AUTH_CONFIG:JSON.stringify({...JSON.parse(env.ROSTER_AUTH_CONFIG),migrationEnabled:true,migrationEmails:[email]})};
for(const action of ['loadAccountContext','login']){
 const requestBody={action,email:legacyEmail,password:legacyPassword,responseMode:'fast'};
 const req=new Request('https://app.example.test/api/state',{method:'POST',headers:{Origin:'https://app.example.test','Content-Type':'application/json'},body:JSON.stringify(requestBody)});
 const c=context(req,rolloutEnv);c.data.stateAction=action;
 assert.equal((await stateRoute(c)).status,200,'non-cohort legacy '+action+' remains usable');
}
assert.equal(sqlite.prepare('SELECT password_hash FROM account_profiles WHERE email=?').get(legacyEmail).password_hash,legacyHash);
const createReq=new Request('https://app.example.test/api/state',{method:'POST',headers:{Origin:'https://app.example.test','Content-Type':'application/json'},body:JSON.stringify({action:'login',mode:'create',email:legacyEmail,password:legacyPassword})});
assert.equal((await stateRoute(context(createReq,rolloutEnv))).status,409,'new-account requests cannot create weak hashes during rollout');
const badCookieRequest=new Request('https://app.example.test/api/state',{method:'POST',headers:{Origin:'https://app.example.test','Content-Type':'application/json',Cookie:SESSION_COOKIE+'=bad'},body:JSON.stringify({action:'loadAccountContext',email:legacyEmail,password:legacyPassword})});
const badCookieContext=context(badCookieRequest,rolloutEnv);badCookieContext.data.stateAction='loadAccountContext';
assert.equal((await stateRoute(badCookieContext)).status,401,'non-cohort users cannot downgrade an invalid cookie');
const first=await issueAccountSession(db,{email,subject,persistent:true},{now});
assert.match(first.setCookie,/^__Host-roster_session=[a-f0-9]{64}; Path=\/; Secure; HttpOnly; SameSite=Lax; Max-Age=2592000$/);
const token=first.setCookie.split(';')[0].split('=')[1];
assert.notEqual(sqlite.prepare('SELECT token_hash FROM account_sessions').get().token_hash,token);
assert.equal(sqlite.prepare('SELECT token_hash FROM account_sessions').get().token_hash,createHash('sha256').update(token).digest('hex'));
const mark=queries.length; const session=await readAccountSession(db,request(first.setCookie),{now});
assert.equal(session.email,email); assert.equal(queries.length,mark+1); assert.equal(queries.at(-1).write,false);
assert.equal(session.absolute_expires_at,now+SESSION_OUTER_MS);
assert.equal((await renewAccountSession(db,request(first.setCookie),session,{now:now+1000})).renewed,false);
assert.equal(queries.length,mark+1,'ordinary use and early renewal have no writes');
const renewedAt=now+SESSION_RENEW_MS;
const concurrent=await Promise.all([renewAccountSession(db,request(first.setCookie),session,{now:renewedAt}),renewAccountSession(db,request(first.setCookie),session,{now:renewedAt})]);
assert.equal(concurrent.filter(x=>x.renewed).length,1,'parallel tabs coalesce to one successful activity update');
assert.equal((await readAccountSession(db,request(first.setCookie),{now:renewedAt})).authenticated_at,now,'renewal does not extend recent authentication');
assert.equal(sqlite.prepare('SELECT absolute_expires_at FROM account_sessions').get().absolute_expires_at,now+SESSION_OUTER_MS);
assert.throws(()=>requireRecentAuthentication(session,{now:now+11*60000}),e=>e.code==='auth-reauthentication-required');
for(const bad of [first.setCookie.split(';')[0]+'; '+first.setCookie.split(';')[0],SESSION_COOKIE+'=bad']) await assert.rejects(()=>readAccountSession(db,new Request('https://app.example.test',{headers:{Cookie:bad}})),e=>e.status===401);
assert.throws(()=>requireSameOrigin(request(first.setCookie,{},'https://attacker.test')),e=>e.status===403);
assert.throws(()=>requireSameOrigin(new Request('http://app.example.test',{method:'POST',headers:{Origin:'http://app.example.test'}})),e=>e.status===403);
assert.equal((await sessionRoute(context(request(first.setCookie)))).status,200);
const rawContextAccount=await loadAccountMirror(db,email);
const fullPrepared=await prepareAccountResponse(null,rawContextAccount,{db,identityDiscoveryEnabled:false});
const beforeMinimal=queries.length;const minimalPrepared=await prepareAccountContextResponse(rawContextAccount,{db,identityDiscoveryEnabled:false});
for(const key of ['role','realName','claims','nameMatches','subscription','insightsEnabled','facilityOverviewEnabled','nonClinical','directorViewEnabled','issueConfig','identityDiscoveryUnavailable','availableDoctors'])assert.deepEqual(minimalPrepared[key],fullPrepared[key],'context parity: '+key);
for(const key of ['mode','isSms','workingToday','facilityKey','today'])assert.equal(minimalPrepared.facilityOverviewAccess[key],fullPrepared.facilityOverviewAccess[key]);
assert.equal(queries.slice(beforeMinimal).some(q=>/custom_events|account_states|account_hospital_locations/.test(q.sql)),false,'context does not prepare discarded calendar state');
const privateContextRequest=new Request('https://app.example.test/api/state',{method:'POST',headers:{Origin:'https://app.example.test','Content-Type':'application/json',Cookie:first.setCookie.split(';')[0]},body:JSON.stringify({action:'loadAccountContext',responseMode:'fast'})});
const privateContext=context(privateContextRequest);privateContext.data.stateAction='loadAccountContext';
const privateContextResponse=await stateRoute(privateContext);
assert.equal(privateContextResponse.status,200,'existing state route accepts a verified cookie without password/email fields');
assert.equal(privateContextResponse.headers.get('Cache-Control'),'no-store');
assert.equal((await sessionRoute(context(request(first.setCookie,{},'https://attacker.test')))).status,403);
const legacyCalendarRequest=new Request('https://app.example.test/api/state',{method:'POST',headers:{Origin:'https://app.example.test','Content-Type':'application/json'},body:JSON.stringify({action:'loadCalendarEvents',email,password:'legacy'})});
const legacyCalendarContext=context(legacyCalendarRequest);legacyCalendarContext.data.stateAction='loadCalendarEvents';
const beforeLegacyCalendar=queries.length;assert.equal((await stateRoute(legacyCalendarContext)).status,401);assert.equal(queries.length,beforeLegacyCalendar,'calendar fast path cannot retain legacy routine password authentication');
const badState=new Request('https://app.example.test/api/state',{method:'POST',headers:{Origin:'https://app.example.test','Content-Type':'application/json',Cookie:SESSION_COOKIE+'=bad'},body:JSON.stringify({action:'loadAccountContext',email,password:'irrelevant'})});
assert.equal((await stateRoute(context(badState))).status,401,'bad cookie must not fall back to a supplied password');
const forgedActor=new Request('https://app.example.test/api/state',{method:'POST',headers:{Origin:'https://app.example.test','Content-Type':'application/json',Cookie:first.setCookie.split(';')[0]},body:JSON.stringify({action:'loadAccountContext',email:other})});
assert.equal((await stateRoute(context(forgedActor))).status,403,'body email cannot replace cookie actor');
const currentRole=await readAccountSession(db,request(first.setCookie));

const second=await issueAccountSession(db,{email,subject,persistent:false},{now:now+1});
assert.doesNotMatch(second.setCookie,/Max-Age|Expires|Domain/);
await revokeOtherAccountSessions(db,currentRole,{now:now+10});
await assert.rejects(()=>readAccountSession(db,request(second.setCookie)),e=>e.status===401);
assert.equal((await readAccountSession(db,request(first.setCookie))).email,email);
await assert.rejects(()=>readAccountSession(db,request(first.setCookie),{now:now+SESSION_OUTER_MS}),e=>e.status===401);
const nearEnd={...session,renew_after:0,idle_expires_at:now+SESSION_OUTER_MS-1};
sqlite.prepare('UPDATE account_sessions SET renew_after=0,idle_expires_at=?').run(nearEnd.idle_expires_at);
await renewAccountSession(db,request(first.setCookie),nearEnd,{now:now+SESSION_OUTER_MS-1000});
assert.equal(sqlite.prepare('SELECT idle_expires_at FROM account_sessions').get().idle_expires_at,now+SESSION_OUTER_MS);
assert.throws(()=>sqlite.prepare("UPDATE account_profiles SET password_hash='changed',password_salt='changed' WHERE email=?").run(email),/Retired password/);
sqlite.prepare("UPDATE account_authentication_links SET phase='pending' WHERE email=?").run(email);
sqlite.prepare("UPDATE account_profiles SET password_hash='changed',password_salt='changed' WHERE email=?").run(email);
await assert.rejects(()=>readAccountSession(db,request(first.setCookie)),e=>e.status===401,'legacy password/reset paths revoke sessions through the trigger');
sqlite.prepare("UPDATE account_profiles SET password_hash='',password_salt='' WHERE email=?").run(email);
sqlite.prepare("UPDATE account_authentication_links SET phase='ready' WHERE email=?").run(email);
// Native route login: exact provider subject/email only; no provider tokens in JSON.
const originalFetch=globalThis.fetch; let providerCalls=0;
globalThis.fetch=async()=>{providerCalls++;return Response.json({user:{id:subject,email,email_confirmed_at:'2026-10-16T00:00:00Z'},access_token:'private-provider-access',refresh_token:'private-provider-refresh',expires_in:3600});};
try{
 const login=await sessionRoute(context(request(null,{action:'login',email,password:'fixture',remember:true})));
 assert.equal(login.status,200);const text=await login.text();assert.equal(text.includes('private-provider'),false);assert.equal(text.includes('token'),false);
 const cookie=login.headers.get('Set-Cookie');assert.match(cookie,/HttpOnly/);assert.equal(login.headers.get('Cache-Control'),'no-store');
 assert.equal((await sessionRoute(context(request(cookie,{action:'logout'})))).status,200);
 await assert.rejects(()=>readAccountSession(db,request(cookie)),e=>e.status===401);
 const before=providerCalls;
 const beforeDenied=queries.length;
 const blocked={...env,AUTH_LOGIN_LIMITER:{async limit(){return{success:false};}}};
 assert.equal((await sessionRoute(context(request(null,{action:'login',email,password:'fixture'}),blocked))).status,429);
 assert.equal(queries.length,beforeDenied,'rate-limited guesses do not read/write D1');assert.equal(providerCalls,before);
 assert.equal((await sessionRoute(context(request(null,{action:'login',email,password:'fixture'}),{...env,AUTH_LOGIN_LIMITER:null}))).status,503);
 assert.equal(queries.length,beforeDenied,'missing admission binding fails closed before D1');
 const oversized=await sessionRoute(context(request(null,{action:'login',email,password:'x'.repeat(9000)})));assert.equal(oversized.status,413);assert.equal(providerCalls,before);
 const missingOrigin=new Request('https://app.example.test/api/auth/session',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});assert.equal((await sessionRoute(context(missingOrigin))).status,403);
}finally{globalThis.fetch=originalFetch;}
const roleCookie=await issueAccountSession(db,{email,subject});
sqlite.prepare("UPDATE account_profiles SET role='creator' WHERE email=?").run(email);
await assert.rejects(()=>readAccountSession(db,request(roleCookie.setCookie)),e=>e.status===401,'privilege changes require fresh authentication');
const promoted=await issueAccountSession(db,{email,subject});assert.equal((await readAccountSession(db,request(promoted.setCookie))).role,'creator');
sqlite.prepare("UPDATE account_profiles SET role='user' WHERE email=?").run(email);
await assert.rejects(()=>readAccountSession(db,request(promoted.setCookie)),e=>e.status===401,'demotion immediately revokes cookies');
const deleted=await issueAccountSession(db,{email,subject});sqlite.prepare('DELETE FROM account_profiles WHERE email=?').run(email);
await assert.rejects(()=>readAccountSession(db,request(deleted.setCookie)),e=>e.status===401);
for(let i=0;i<125;i++)sqlite.prepare('INSERT INTO account_sessions VALUES(?,?,?,?,?,?,0)').run(i.toString(16).padStart(64,'0'),other,0,1,2,0);
assert.equal(await cleanExpiredAccountSessions(db,{now:now+SESSION_IDLE_MS}),100);
assert.equal(sqlite.prepare('SELECT COUNT(*) AS n FROM account_sessions').get().n,25,'cleanup is bounded');
for(const sql of ['SELECT * FROM account_sessions WHERE token_hash=?','SELECT token_hash FROM account_sessions WHERE email=? ORDER BY authenticated_at DESC,token_hash DESC','SELECT token_hash FROM account_sessions WHERE idle_expires_at<=? ORDER BY idle_expires_at,token_hash LIMIT 100']){
 const plan=sqlite.prepare('EXPLAIN QUERY PLAN '+sql).all(sql.includes('email=')?other:0).map(x=>x.detail).join(' ');
 assert.match(plan,/SEARCH .* USING (?:COVERING )?INDEX/);assert.doesNotMatch(plan,/SCAN account_sessions/);
}
console.log('Account sessions passed: hashed opaque tokens, exact origin, cookie flags, one indexed read/no routine writes, coalesced activity, idle/outer expiry, recent authentication, current authority, actor binding, no fallback from invalid cookies, password/deletion revocation, bounded cleanup and provider-token containment. No external services.');
