import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {readFile,readdir} from 'node:fs/promises';
import {AUTH_PROJECT_ORIGIN} from '../functions/_lib/supabase-auth.js';
import {loadAuthenticationLink} from '../functions/_lib/authentication-migration.js';
import {beginCredentialChange,verifyAndCompleteCredentialChange} from '../functions/_lib/account-credential-change.js';
import {authenticateExistingAccount} from '../functions/_lib/account-authentication.js';
import {issueAccountSession,readAccountSession} from '../functions/_lib/account-sessions.js';
import {onRequestPost as authRoute} from '../functions/api/auth/session.js';
const sql=new DatabaseSync(':memory:');sql.exec('PRAGMA foreign_keys=ON');
for(const name of (await readdir(new URL('../migrations/',import.meta.url))).filter(n=>n.endsWith('.sql')).sort())sql.exec(await readFile(new URL('../migrations/'+name,import.meta.url),'utf8'));
const db={prepare(q){return{args:[],bind(...a){this.args=a;return this;},async first(){return sql.prepare(q).get(...this.args)||null;},async run(){const before=sql.prepare('SELECT total_changes() AS n').get().n;sql.prepare(q).run(...this.args);const after=sql.prepare('SELECT total_changes() AS n').get().n;return{meta:{changes:Number(after-before)}};}};},async batch(ss){sql.exec('BEGIN');try{const r=[];for(const s of ss)r.push(await s.run());sql.exec('COMMIT');return r;}catch(e){sql.exec('ROLLBACK');throw e;}}};
const email='race@example.test',subject='11111111-2222-4333-8444-555555555555',oldPassword='Original synthetic password',newPassword='Replacement synthetic password';
sql.prepare("INSERT INTO account_profiles(email,role,password_hash,password_salt,subscription_token) VALUES(?,'user','','','preserved-feed')").run(email);
sql.prepare("INSERT INTO account_authentication_links(email,provider,provider_subject,phase,credential_revision,created_at,updated_at) VALUES(?,'supabase',?,'ready','revision','now','now')").run(email,subject);
const env={ROSTER_DB:db,AUTH_LOGIN_LIMITER:{async limit(){return{success:true};}},ROSTER_AUTH_CONFIG:JSON.stringify({enabled:true,sessionsEnabled:true,url:AUTH_PROJECT_ORIGIN,publishableKey:'sb_publishable_fixture',secretKey:'sb_secret_fixture'})};
const user={id:subject,email,email_confirmed_at:'2026-10-17T00:00:00Z'};let password=oldPassword,providerCalls=0,mutationTimesOut=false,rejectChange=false,holdLogin=false,releaseLogin,loginReached;
const oldFetch=globalThis.fetch;
globalThis.fetch=async(url,init)=>{
 assert.equal(new URL(url).origin,AUTH_PROJECT_ORIGIN);providerCalls++;const body=JSON.parse(init.body||'{}');
 if(url.endsWith('/token?grant_type=password')){
  const response=body.password===password?Response.json({user,access_token:'fixture-access',refresh_token:'fixture-refresh',expires_in:3600}):Response.json({code:'invalid'},{status:400});
  if(holdLogin){holdLogin=false;loginReached();await new Promise(r=>{releaseLogin=r;});}return response;
 }
 if(url.includes('/admin/users/')&&init.method==='GET')return Response.json(user);
 if(url.includes('/admin/users/')&&init.method==='PUT'){password=body.password;return Response.json(user);}
 if(url.endsWith('/user')&&init.method==='PUT'){if(rejectChange)return Response.json({code:'weak_password'},{status:422});password=body.password;if(mutationTimesOut)throw Error('lost mutation response');return Response.json(user);}
 throw Error('Unexpected provider request');
};
const request=(cookie,body)=>new Request('https://app.example.test/api/auth/session',{method:'POST',headers:{Origin:'https://app.example.test','Content-Type':'application/json',Cookie:cookie.split(';')[0]},body:JSON.stringify(body)});
try{
 const initial=await issueAccountSession(db,{email,subject});
 // Reproduce an old-password provider reply arriving after revocation.
 let reached;const signal=new Promise(r=>{reached=r;});loginReached=reached;holdLogin=true;
 const oldLogin=authenticateExistingAccount(env,{email,password:oldPassword});await signal;
 sql.prepare("INSERT INTO account_email_challenges VALUES('old-reset',?,?, 'recovery','recovery','digest',?,'issued',?)").run(email,subject,Date.now()+100000,Date.now());
 sql.prepare("INSERT INTO account_invites(token_hash,email,expires_at,created_at,updated_at) VALUES('old-invite',?,'2099-01-01','now','now')").run(email);
 const originalLink=await loadAuthenticationLink(db,email);
 const operation=await beginCredentialChange(env,originalLink,newPassword);
 assert.equal(sql.prepare('SELECT COUNT(*) AS n FROM account_sessions').get().n,0,'admission revokes existing sessions atomically');
 assert.equal(sql.prepare("SELECT phase FROM account_email_challenges WHERE id='old-reset'").get().phase,'complete');
 assert.notEqual(sql.prepare("SELECT revoked_at FROM account_invites WHERE token_hash='old-invite'").get().revoked_at,'');
 await assert.rejects(()=>readAccountSession(db,request(initial.setCookie,{})),e=>e.status===401);
 releaseLogin();const staleProof=await oldLogin;
 await assert.rejects(()=>issueAccountSession(db,{email,subject,credentialVersion:staleProof.credentialVersion}),e=>e.status===401,'proof obtained before reset cannot issue a session');
 await assert.rejects(()=>issueAccountSession(db,{email,subject,credentialVersion:operation.credential_version}),e=>e.status===401,'even current-version proof cannot issue while mutation is pending');
 await assert.rejects(async()=>beginCredentialChange(env,await loadAuthenticationLink(db,email),'Competing synthetic password'),e=>e.code==='auth-password-pending');
 const before=providerCalls;
 await assert.rejects(()=>authenticateExistingAccount(env,{email,password:oldPassword}),e=>e.code==='auth-password-pending');assert.equal(providerCalls,before,'old password cannot unlock a pending new credential');
 await assert.rejects(()=>verifyAndCompleteCredentialChange(env,operation,newPassword),e=>e.code==='invalid-credentials','target HMAC is not a substitute for actual provider password proof');
 // Model an uncertain provider update: the change occurred, but no response.
 password=newPassword;
 const resumed=await authenticateExistingAccount(env,{email,password:newPassword});
 assert.equal(resumed.credentialVersion,2);await assert.rejects(()=>beginCredentialChange(env,originalLink,'Stale admission fixture'),e=>e.code==='auth-password-pending','a failed uniquely-keyed CAS must remain closed even with trigger-aware counts');assert.equal((await loadAuthenticationLink(db,email)).credential_operation_id,'');
 assert.equal((await loadAuthenticationLink(db,email)).credential_target_digest,'','no residual password proof after completion');
 await assert.rejects(()=>issueAccountSession(db,{email,subject,credentialVersion:staleProof.credentialVersion}),e=>e.status===401,'stale proof stays revoked after completion');
 const fresh=await issueAccountSession(db,{email,subject,credentialVersion:resumed.credentialVersion});
 assert.equal((await readAccountSession(db,request(fresh.setCookie,{}))).email,email);
 // Exercise the actual password route, including a lost provider response and
 // retry through ordinary password login. No blind provider update is retried.
 mutationTimesOut=true;const nextPassword='Another replacement fixture';
 const failed=await authRoute({env,request:request(fresh.setCookie,{action:'changePassword',expectedAccount:email,currentPassword:newPassword,newPassword:nextPassword})});
 assert.equal(failed.status,503);assert.notEqual((await loadAuthenticationLink(db,email)).credential_operation_id,'');
 assert.equal(sql.prepare('SELECT COUNT(*) AS n FROM account_sessions').get().n,0);
 mutationTimesOut=false;const recovered=await authenticateExistingAccount(env,{email,password:nextPassword});assert.equal(recovered.credentialVersion,4);
 const current=await issueAccountSession(db,{email,subject,credentialVersion:recovered.credentialVersion});
 const changed=await authRoute({env,request:request(current.setCookie,{action:'changePassword',expectedAccount:email,currentPassword:nextPassword,newPassword:'Final synthetic password'})});assert.equal(changed.status,200);
 assert.equal((await loadAuthenticationLink(db,email)).credential_version,6);
 await assert.rejects(()=>readAccountSession(db,request(current.setCookie,{})),e=>e.status===401);
 // A definite provider policy rejection cancels admission and keeps the old
 // password usable. An uncertain write must instead stay fenced.
 const preRejected=await authenticateExistingAccount(env,{email,password:'Final synthetic password'});
 const rejectedCookie=await issueAccountSession(db,{email,subject,credentialVersion:preRejected.credentialVersion});
 rejectChange=true;
 const rejected=await authRoute({env,request:request(rejectedCookie.setCookie,{action:'changePassword',expectedAccount:email,currentPassword:'Final synthetic password',newPassword:'Policy rejected fixture'})});
 assert.equal(rejected.status,400);assert.equal((await rejected.json()).code,'auth-password-policy');
 assert.equal((await readAccountSession(db,request(rejected.headers.get('Set-Cookie'),{}))).email,email,'known rejection preserves the current browser session without accepting stale proof');
 assert.equal((await loadAuthenticationLink(db,email)).credential_operation_id,'');rejectChange=false;
 assert.equal((await authenticateExistingAccount(env,{email,password:'Final synthetic password'})).credentialVersion,8);
 // An uncertain write is never repeated, even for the intended target. A
 // late first write must not be able to overwrite a later password operation.
 const unappliedTarget='Unapplied replacement fixture';
 await beginCredentialChange(env,await loadAuthenticationLink(db,email),unappliedTarget);
 const callsBefore=providerCalls;
 await assert.rejects(()=>authenticateExistingAccount(env,{email,password:unappliedTarget}),e=>e.code==='auth-password-pending');
 assert.equal(providerCalls,callsBefore+1,'only verify; never a blind privileged rewrite');
 assert.equal(password,'Final synthetic password');
 const serialized=JSON.stringify(await changed.json());assert.equal(serialized.includes('fixture-access'),false);assert.equal(serialized.includes('password'),false);
 assert.equal(sql.prepare('SELECT subscription_token FROM account_profiles WHERE email=?').get(email).subscription_token,'preserved-feed');
 console.log('Credential races passed: held old-password login, pending issuance denial, atomic revocation, serialized mutation, exact intended-target/provider proof, lost-response resume, stale-proof rejection after completion and real password-route recovery. No external services.');
}finally{globalThis.fetch=oldFetch;sql.close();}
