import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {readFile,readdir} from 'node:fs/promises';
import {onRequestPost as route} from '../functions/api/auth/passkeys.js';
import {onRequestPost as sessionRoute} from '../functions/api/auth/session.js';
import {issueAccountSession,readAccountSession} from '../functions/_lib/account-sessions.js';
import {AUTH_PROJECT_ORIGIN} from '../functions/_lib/supabase-auth.js';
import {PASSKEY_ORIGIN,PASSKEY_RP,sealPasskeyCookie,readPasskeyCookie,passkeyLoginGrant} from '../functions/_lib/account-passkeys.js';
import {passkeyOptions,passkeyResponse,createPasskeyCeremony} from '../public/static/passkeys.js';
import {createAccountTransport} from '../public/static/account-transport.js';
import admissionService from '../workers/auth-admission/index.js';

const sqlite=new DatabaseSync(':memory:');sqlite.exec('PRAGMA foreign_keys=ON');
for(const name of(await readdir('migrations')).filter(x=>x.endsWith('.sql')).sort())sqlite.exec(await readFile('migrations/'+name,'utf8'));
const queries=[];
const db={prepare(sql){return{args:[],bind(...args){this.args=args;return this;},async first(){queries.push(sql);return sqlite.prepare(sql).get(...this.args)||null;},async all(){queries.push(sql);return{results:sqlite.prepare(sql).all(...this.args)};},async run(){queries.push(sql);return{meta:{changes:Number(sqlite.prepare(sql).run(...this.args).changes)}};}};},async batch(items){sqlite.exec('BEGIN');try{const result=[];for(const item of items)result.push(await item.run());sqlite.exec('COMMIT');return result;}catch(error){sqlite.exec('ROLLBACK');throw error;}}};
const email='passkey@example.test',other='other@example.test',subject='11111111-2222-4333-8444-555555555555',otherSubject='aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee';
const challenge='cccccccc-1111-4222-8333-dddddddddddd',id='dddddddd-1111-4222-8333-eeeeeeeeeeee';
for(const[e,s]of[[email,subject],[other,otherSubject]]){
 sqlite.prepare("INSERT INTO account_profiles(email,role,password_hash,password_salt,subscription_token) VALUES(?,'user','','',?)").run(e,'feed-'+e);
 sqlite.prepare("INSERT INTO account_authentication_links(email,provider,provider_subject,phase,credential_revision,created_at,updated_at) VALUES(?,'supabase',?,'ready','revision','now','now')").run(e,s);
}
const env={ROSTER_DB:db,AUTH_LOGIN_LIMITER:{async limit(){return{success:true};}},ROSTER_AUTH_CONFIG:JSON.stringify({enabled:true,sessionsEnabled:true,passkeysEnabled:true,url:AUTH_PROJECT_ORIGIN,publishableKey:'sb_publishable_fixture',secretKey:'sb_secret_fixture'})};
const request=(body,cookie='',origin=PASSKEY_ORIGIN)=>new Request(origin+'/api/auth/passkeys',{method:'POST',headers:{Origin:origin,'Content-Type':'application/json',Cookie:cookie,'CF-Connecting-IP':'192.0.2.1'},body:JSON.stringify(body)});
const call=(body,cookie='',environment=env,origin=PASSKEY_ORIGIN)=>route({env:environment,request:request(body,cookie,origin)});
const pair=cookie=>cookie.split(';')[0];
let session=await issueAccountSession(db,{email,subject}),cookie=pair(session.setCookie);
const user={id:subject,email,email_confirmed_at:'2026-10-18T00:00:00Z'},token={user,access_token:'synthetic-provider-access',refresh_token:'synthetic-refresh',expires_in:3600};
const optionData={challenge_id:challenge,options:{challenge:'YWJj',rpId:PASSKEY_RP,userVerification:'preferred'}};
const registrationData={challenge_id:challenge,options:{challenge:'YWJj',rp:{id:PASSKEY_RP,name:'Roster'},user:{id:'YWJj',name:email,displayName:'Fixture'},pubKeyCredParams:[{type:'public-key',alg:-7}]}};
const authData=new Uint8Array(37);authData[32]=5;
const credential={id:'fixture-public-id',type:'public-key',response:{clientDataJSON:'e30',authenticatorData:Buffer.from(authData).toString('base64url'),signature:'YWJj',userHandle:Buffer.from(subject).toString('base64url')}};
const original=globalThis.fetch,calls=[];let providerPasskeys=[],onVerifyRegistration=null,onVerifyAuthentication=null;
globalThis.fetch=async(url,init)=>{
 calls.push({url,init});assert.equal(new URL(url).origin,AUTH_PROJECT_ORIGIN);assert.equal(init.redirect,'manual');
 if(url.endsWith('token?grant_type=password'))return Response.json(token);
 if(url.endsWith('/passkeys/registration/options')){assert.equal(init.headers.Authorization,'Bearer '+token.access_token);return Response.json(structuredClone(registrationData));}
 if(url.endsWith('/passkeys/registration/verify')){await onVerifyRegistration?.();providerPasskeys=[{id,friendly_name:'Fixture passkey',created_at:'now'}];return Response.json({id});}
 if(url.endsWith('/passkeys/authentication/options'))return Response.json(structuredClone(optionData));
 if(url.endsWith('/passkeys/authentication/verify')){await onVerifyAuthentication?.();return Response.json(token);}
 if(url.endsWith('/passkeys/'+id)&&init.method==='DELETE'){providerPasskeys=[];return new Response(null,{status:204});}
 if(url.endsWith('/admin/users/'+subject+'/passkeys'))return Response.json(providerPasskeys);
 throw Error('Unexpected synthetic provider route');
};
try{
 const before=queries.length,network=calls.length;
 assert.equal((await call({action:'status'},cookie,{...env,ROSTER_AUTH_CONFIG:'{}'})).status,503);
 assert.equal((await call({action:'status'},cookie,env,'https://preview.example.test')).status,503);
 assert.equal(queries.length,before);assert.equal(calls.length,network,'disabled and preview requests have no D1/provider work');
 const foreign=request({action:'status'},cookie);foreign.headers.set('Origin','https://other.invalid');
 assert.equal((await route({env,request:foreign})).status,403);assert.equal(queries.length,before);
 assert.equal((await call({action:'status'})).status,401);
 assert.equal((await call({action:'status',expectedAccount:other},cookie)).status,403);
 assert.equal(calls.length,network,'other accounts never reach privileged provider listing');
 let response=await call({action:'status',expectedAccount:email},cookie);assert.equal(response.status,200);assert.equal((await response.json()).remind,true);
 response=await call({action:'reminders',enabled:false},cookie);assert.equal((await response.json()).changed,true);
 response=await call({action:'reminders',enabled:false},cookie);assert.equal((await response.json()).changed,false,'unchanged preference does not rewrite the row');
 const count=calls.length;response=await call({action:'status',prompt:true},cookie);assert.equal((await response.json()).remind,false);assert.equal(calls.length,count,'suppressed reminders skip Supabase entirely');
 const otherSession=await issueAccountSession(db,{email,subject});
 assert.equal((await(await call({action:'status',prompt:true},pair(otherSession.setCookie))).json()).remind,false,'preference follows account across devices');
 await call({action:'reminders',enabled:true},cookie);
 response=await call({action:'startRegistration'},cookie);assert.equal(response.status,401);assert.equal((await response.json()).code,'auth-passkey-password-required');
 // Password-login grant is encrypted and bound to the issued app session.
 const grant=await passkeyLoginGrant(env,request({}),{...token,email,subject,credentialVersion:0,accessToken:token.access_token},session);
 assert.match(grant,/Secure; HttpOnly; SameSite=Lax; Max-Age=300/);assert.equal(grant.includes(token.access_token),false);
 response=await call({action:'startRegistration'},cookie+'; '+pair(grant));assert.equal(response.status,200);
 const start=await response.json(),regCookie=pair(response.headers.get('Set-Cookie'));
 assert.equal(start.options.authenticatorSelection.userVerification,'required');assert.equal(JSON.stringify(start).includes(token.access_token),false);
 assert.equal(calls.filter(x=>x.url.includes('grant_type=password')).length,0,'immediate setup reuses verified login, without another password');
 assert.equal((await call({action:'finishRegistration',challengeId:challenge,credential},pair(otherSession.setCookie)+'; '+regCookie)).status,400,'enrollment cannot cross app sessions');
 assert.equal((await call({action:'finishRegistration',challengeId:otherSubject,credential},cookie+'; '+regCookie)).status,400);
 response=await call({action:'finishRegistration',challengeId:challenge,credential},cookie+'; '+regCookie);assert.equal(response.status,200);
 assert.equal((await(await call({action:'status',prompt:true},cookie)).json()).remind,false,'existing passkey prevents reminder');
 response=await call({action:'remove',id},cookie);assert.equal(response.status,200);assert.equal(providerPasskeys.length,0);
 // Tampering, duplicate cookies, expiry and rotation all fail closed.
 const sealed=await sealPasskeyCookie(env,{kind:'authentication',challengeId:challenge});
 for(const value of [pair(sealed).slice(0,-2)+'xx',pair(sealed)+'; '+pair(sealed)])await assert.rejects(()=>readPasskeyCookie(env,request({},value)));
 const expired=await sealPasskeyCookie(env,{kind:'authentication'},Date.now()-300001);
 await assert.rejects(()=>readPasskeyCookie(env,request({},pair(expired))));
 const rotated={...env,ROSTER_AUTH_CONFIG:JSON.stringify({...JSON.parse(env.ROSTER_AUTH_CONFIG),secretKey:'sb_secret_replaced'})};
 await assert.rejects(()=>readPasskeyCookie(rotated,request({},pair(sealed))));
 // Logout while Supabase is verifying enrollment removes the exact new key.
 response=await call({action:'startRegistration',password:'Synthetic password'},cookie);const raceCookie=pair(response.headers.get('Set-Cookie'));
 onVerifyRegistration=()=>db.prepare('DELETE FROM account_sessions WHERE email=?').bind(email).run();
 response=await call({action:'finishRegistration',challengeId:challenge,credential},cookie+'; '+raceCookie);assert.equal(response.status,401);assert.equal(providerPasskeys.length,0);onVerifyRegistration=null;
 // Passkey authentication is anonymous until verified, exact-link bound and
 // creates the same opaque app cookie; provider JWTs never enter the response.
 response=await call({action:'startSignIn',remember:true});const loginCookie=pair(response.headers.get('Set-Cookie'));assert.equal(response.status,200);assert.equal((await response.json()).options.userVerification,'required');
 response=await call({action:'finishSignIn',challengeId:challenge,credential},loginCookie);assert.equal(response.status,200);
 assert.deepEqual(await response.json(),{ok:true,email});const issued=response.headers.getSetCookie().find(x=>x.startsWith('__Host-roster_session='));assert.match(issued,/HttpOnly; SameSite=Lax; Max-Age=/);
 assert.equal((await readAccountSession(db,request({},pair(issued)))).email,email);
 const missingUV=structuredClone(credential);missingUV.response.authenticatorData=Buffer.alloc(37).toString('base64url');
 const beforeUV=calls.length;assert.equal((await call({action:'finishSignIn',challengeId:challenge,credential:missingUV},loginCookie)).status,400);assert.equal(calls.length,beforeUV,'UV must be present in signed flags, not just requested from browser');
 onVerifyAuthentication=()=>db.prepare('UPDATE account_authentication_links SET credential_operation_id=? WHERE email=?').bind('pending-reset',email).run();
 assert.equal((await call({action:'finishSignIn',challengeId:challenge,credential},loginCookie)).status,401,'pending password mutation cannot issue a passkey session');onVerifyAuthentication=null;
 await db.prepare("UPDATE account_authentication_links SET credential_operation_id='',provider_subject=? WHERE email=?").bind(otherSubject.replace('aaaa','bbbb'),email).run();
 assert.equal((await call({action:'finishSignIn',challengeId:challenge,credential},loginCookie)).status,401,'matching email alone cannot adopt a provider identity');
 await db.prepare('UPDATE account_authentication_links SET provider_subject=? WHERE email=?').bind(subject,email).run();
 response=await call({action:'startSignIn'},'',{...env,AUTH_LOGIN_LIMITER:{async limit(){return{success:false};}}});assert.equal(response.status,429);
 response=await call({action:'reminders',enabled:'false'},pair(issued));assert.equal(response.status,400);
 const tooLarge=request({action:'status',padding:'x'.repeat(33000)},pair(issued));assert.equal((await route({env,request:tooLarge})).status,413);
 // Actual password route appends the grant only when enabled at exact RP.
 const loginRequest=new Request(PASSKEY_ORIGIN+'/api/auth/session',{method:'POST',headers:{Origin:PASSKEY_ORIGIN,'Content-Type':'application/json'},body:JSON.stringify({action:'login',email,password:'Synthetic fixture password'})});
 response=await sessionRoute({env,request:loginRequest});assert.equal(response.status,200);assert.equal(response.headers.getSetCookie().length,2);
 assert.equal((await response.text()).includes(token.access_token),false);
 assert.equal(queries.some(sql=>/roster_events|roster_files|doctor_identities/.test(sql)),false,'passkey journeys never scan rosters or doctor identity tables');
}finally{globalThis.fetch=original;sqlite.close();}

const creation=passkeyOptions(registrationData.options,true);assert.deepEqual([...creation.challenge],[97,98,99]);assert.deepEqual([...creation.user.id],[97,98,99]);
const assertion=passkeyOptions(optionData.options,false);assert.deepEqual([...assertion.challenge],[97,98,99]);
const serial=passkeyResponse({id:'YWJj',rawId:Uint8Array.of(97,98,99),response:{clientDataJSON:Uint8Array.of(1,2),authenticatorData:Uint8Array.of(3),signature:Uint8Array.of(4),userHandle:null}},false);assert.equal(serial.response.clientDataJSON,'AQI');assert.equal(serial.rawId,'YWJj');
let activeSignal;const hardware=createPasskeyCeremony({get:async options=>{activeSignal=options.signal;return new Promise((_,reject)=>activeSignal.addEventListener('abort',()=>reject(new DOMException('cancel','AbortError'))));}});
const waiting=hardware.run(optionData.options);hardware.cancel();await assert.rejects(waiting,error=>error.name==='AbortError');
// A cancelled/stale passkey ceremony cannot commit after logout or a newer
// login, including the Set-Cookie response cleanup used by password sign-in.
const storage=()=>{const m=new Map();return{getItem:k=>m.get(k),setItem:(k,v)=>m.set(k,v),removeItem:k=>m.delete(k)};};
let release,finishes=0;const transport=createAccountTransport({persistentStorage:storage(),tabStorage:storage(),fetchImpl:async(url,init)=>{
 const body=JSON.parse(init.body);
 if(body.action==='startSignIn')return Response.json({ok:true,options:optionData.options,challengeId:challenge});
 if(body.action==='finishSignIn'){finishes++;return Response.json({ok:true,email});}
 return Response.json({ok:true});
}});
const pending=transport.passkeySignIn(async()=>new Promise(resolve=>release=resolve),true);
await new Promise(resolve=>setTimeout(resolve,0));await transport.logout();release(credential);assert.equal((await pending).status,401);assert.equal(finishes,0);
assert.equal(transport.expected(),false);
for(const mode of ['login','passkey']){
 let chosen='';const limiter={LOGIN_RATE:{async limit(){chosen='login';return{success:true};}},PASSKEY_RATE:{async limit(){chosen='passkey';return{success:true};}}};
 const r=await admissionService.fetch(new Request('https://admission.internal/check',{method:'POST',body:mode+':'+'a'.repeat(64)}),limiter);
 assert.equal(r.status,204);assert.equal(chosen,mode,'anonymous shared-IP passkeys must not consume per-account password limits');
}
assert.equal((await admissionService.fetch(new Request('https://admission.internal/check',{method:'POST',body:'passkey:'+'a'.repeat(64)}),{})).status,503);
console.log('Passkeys passed: disabled/preview/origin gates, exact actor/session/version, encrypted short-lived grant, unchanged preference writes, account-wide suppression, no roster reads, UV enforcement, enrollment/logout cleanup, opaque login, identity and pending-operation denial, cancellation and logout races. No external services or real credentials.');
