import {registerAccount} from '../functions/_lib/account-registration.js';
import {authenticateExistingAccount} from '../functions/_lib/account-authentication.js';
import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {readFile,readdir} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {sendAccountEmailProof,completeAccountEmailProof} from '../functions/_lib/account-email.js';
import {AUTH_PROJECT_ORIGIN} from '../functions/_lib/supabase-auth.js';
import {takeAccountEmailAction,takeInvitationToken} from '../public/static/account-email-action.js';
import {onRequestPost as route} from '../functions/api/auth/session.js';
const sql=new DatabaseSync(':memory:');sql.exec('PRAGMA foreign_keys=ON');
for(const name of (await readdir(new URL('../migrations/',import.meta.url))).filter(n=>n.endsWith('.sql')).sort())sql.exec(await readFile(new URL('../migrations/'+name,import.meta.url),'utf8'));
const db={prepare(q){return{args:[],bind(...a){this.args=a;return this;},async first(){return sql.prepare(q).get(...this.args)||null;},async run(){const r=sql.prepare(q).run(...this.args);return{meta:{changes:Number(r.changes)}};}};},async batch(ss){sql.exec('BEGIN');try{const r=[];for(const s of ss)r.push(await s.run());sql.exec('COMMIT');return r;}catch(e){sql.exec('ROLLBACK');throw e;}}};
const email='recovery@example.test',password='Original fixture only',salt='a'.repeat(32),hash=createHash('sha256').update(salt+':'+password).digest('hex');
const config={enabled:true,sessionsEnabled:true,emailDeliveryReady:true,emailMonthlyLimit:10,legacyMigrationUntil:new Date(Date.now()+86400000).toISOString(),url:AUTH_PROJECT_ORIGIN,publishableKey:'sb_publishable_fixture',secretKey:'sb_secret_fixture'};
const env={ROSTER_DB:db,POSTMARK_API_TOKEN:'synthetic-only',ROSTER_AUTH_CONFIG:JSON.stringify(config),AUTH_LOGIN_LIMITER:{async limit(){return{success:true};}}};
function account(e){sql.prepare("INSERT INTO account_profiles(email,real_name,password_hash,password_salt,role,subscription_token) VALUES(?,'Fixture',?,?,'user','preserved-feed')").run(e,hash,salt);}
account(email);
let user,providerPassword,emails=[],calls=[],tokens=new Map(),failChange=false;
const oldFetch=globalThis.fetch;
globalThis.fetch=async(url,init)=>{
 calls.push({url,method:init.method});const body=init.body?JSON.parse(init.body):{};
 if(url==='https://api.postmarkapp.com/email'){emails.push(body);return Response.json({ErrorCode:0});}
 if(url.includes('/admin/users/')&&init.method==='GET')return user&&url.endsWith('/'+user.id)?Response.json(user):Response.json({message:'missing'},{status:404});
 if(url.endsWith('/admin/users')&&init.method==='POST'){user={id:body.id,email:body.email,email_confirmed_at:body.email_confirm?new Date().toISOString():null};providerPassword=body.password;return Response.json(user);}
 if(url.endsWith('/admin/generate_link')){const token=crypto.randomUUID().replaceAll('-','')+crypto.randomUUID().replaceAll('-','');tokens.set(token,{subject:user.id,type:body.type});return Response.json({...user,hashed_token:token,verification_type:body.type});}
 if(url.endsWith('/verify')){const token=tokens.get(body.token_hash);if(!token||token.type!==body.type)return Response.json({error:'used'},{status:400});tokens.delete(body.token_hash);user.email_confirmed_at=new Date().toISOString();return Response.json({user,access_token:'private-proof-token'});}
 if(url.endsWith('/user')&&init.method==='PUT'){assert.equal(init.headers.Authorization,'Bearer private-proof-token');if(failChange)throw Error('uncertain provider mutation');providerPassword=body.password;return Response.json(user);}
 if(url.endsWith('/token?grant_type=password'))return body.password===providerPassword?Response.json({user,access_token:'private-access',refresh_token:'private-refresh',expires_in:3600}):Response.json({error:'invalid'},{status:400});
 throw Error('Unexpected test request');
};
try{
 await assert.rejects(()=>sendAccountEmailProof({...env,ROSTER_AUTH_CONFIG:JSON.stringify({...config,emailDeliveryReady:false})},{email,purpose:'recovery',origin:'https://rtc.curiousmind.app'}),e=>e.code==='auth-email-unavailable');assert.equal(calls.length,0);
 await assert.rejects(()=>sendAccountEmailProof(env,{email,purpose:'recovery',origin:'https://attacker.test'}),e=>e.status===403);
 await Promise.all([sendAccountEmailProof(env,{email,purpose:'recovery',origin:'https://rtc.curiousmind.app'}),sendAccountEmailProof(env,{email,purpose:'recovery',origin:'https://rtc.curiousmind.app'})]);
 assert.equal(emails.length,1,'concurrent email requests coalesce');assert.notEqual(providerPassword,password,'recovery provisioning never sets a caller-selected unverified password');assert.equal(user.email_confirmed_at,null);
 const message=emails[0];assert.equal(message.TrackLinks,'None');assert.equal(message.TrackOpens,false);assert.equal(message.To,email);
 const url=message.TextBody.match(/https:\/\/rtc\.curiousmind\.app\/#[^\s]+/)[0];const parsed=new URL(url);assert.equal(parsed.search,'');
 let cleaned='';const action=takeAccountEmailAction(parsed,{replaceState(a,b,value){cleaned=value;}});assert.equal(cleaned,'/');assert.equal(action.purpose,'recovery');
 const challenge=sql.prepare('SELECT * FROM account_email_challenges WHERE id=?').get(action.id);assert.notEqual(challenge.token_digest,action.token);
 await assert.rejects(()=>completeAccountEmailProof(env,{id:action.id,token:'f'.repeat(64),newPassword:'Replacement fixture only'}),e=>e.code==='auth-email-invalid');
 await assert.rejects(()=>completeAccountEmailProof(env,{id:action.id,token:action.token,newPassword:'short'}),e=>e.code==='auth-password');
 const result=await completeAccountEmailProof(env,{id:action.id,token:action.token,newPassword:'Replacement fixture only'});assert.deepEqual(result,{ok:true,email,purpose:'recovery'});
 assert.equal(sql.prepare('SELECT password_hash FROM account_profiles WHERE email=?').get(email).password_hash,'');assert.equal(sql.prepare('SELECT phase FROM account_authentication_links WHERE email=?').get(email).phase,'ready');
 assert.equal(sql.prepare('SELECT COUNT(*) AS n FROM account_sessions').get().n,0,'email proof does not create a browser session');
 assert.equal(sql.prepare('SELECT subscription_token FROM account_profiles WHERE email=?').get(email).subscription_token,'preserved-feed');
 await assert.rejects(()=>completeAccountEmailProof(env,{id:action.id,token:action.token,newPassword:'Replacement fixture only'}),e=>e.code==='auth-email-invalid');
 // A second recovery after cooldown revokes app sessions even when update is uncertain.
 sql.prepare('DELETE FROM account_email_delivery_locks').run();sql.prepare('UPDATE account_email_challenges SET created_at=0').run();
 await sendAccountEmailProof(env,{email,purpose:'recovery',origin:'https://rtc.curiousmind.app'});
 const secondURL=new URL(emails.at(-1).TextBody.match(/https:\/\/rtc\.curiousmind\.app\/#[^\s]+/)[0]);const second=takeAccountEmailAction(secondURL,{replaceState(){}});
 sql.prepare('INSERT INTO account_sessions VALUES(?,?,?,?,?,?,0)').run('a'.repeat(64),email,Date.now(),Date.now()+1000000,Date.now()+2000000,Date.now()+100000);
 failChange=true;await assert.rejects(()=>completeAccountEmailProof(env,{id:second.id,token:second.token,newPassword:'Another replacement fixture'}),e=>e.code==='auth-provider-unavailable');assert.equal(sql.prepare('SELECT COUNT(*) AS n FROM account_sessions').get().n,0);failChange=false;
 // New registrations use provider hashing from the beginning, never legacy SHA.
 const signupEnv={...env,POSTMARK_API_TOKEN:undefined,ROSTER_AUTH_CONFIG:JSON.stringify({...config,signupEnabled:true,emailDeliveryReady:false})};
 const emailsBeforeSignup=emails.length;
 const signupEmail='new-person@example.test',signupPassword='New strong fixture password';
 await registerAccount(signupEnv,{email:signupEmail,password:signupPassword,realName:'New Fixture',origin:'https://rtc.curiousmind.app'});
 assert.equal(sql.prepare('SELECT password_hash FROM account_profiles WHERE email=?').get(signupEmail).password_hash,'');
 assert.equal(sql.prepare('SELECT role FROM account_profiles WHERE email=?').get(signupEmail).role,'user');
 assert.equal(emails.length,emailsBeforeSignup,'signup requires no email or Postmark credential');
 await authenticateExistingAccount(signupEnv,{email:signupEmail,password:signupPassword});
 assert.equal(sql.prepare('SELECT phase FROM account_authentication_links WHERE email=?').get(signupEmail).phase,'ready');
 await assert.rejects(()=>registerAccount(signupEnv,{email:signupEmail,password:'Different attempted password',realName:'Changed',origin:'https://rtc.curiousmind.app'}),e=>e.code==='auth-registration-conflict');
 assert.equal(providerPassword,signupPassword,'repeated signup cannot overwrite a verified credential');
 // No provider work or counter writes after the conservative global mail cap.
 const other='other@example.test';account(other);sql.prepare('UPDATE account_auth_email_budget SET reserved=10').run();const before=calls.length;
 await assert.rejects(()=>sendAccountEmailProof(env,{email:other,purpose:'recovery',origin:'https://rtc.curiousmind.app'}),e=>e.code==='auth-email-unavailable');assert.equal(calls.length,before);
 const req=new Request('https://rtc.curiousmind.app/api/auth/session',{method:'POST',headers:{Origin:'https://rtc.curiousmind.app','Content-Type':'application/json'},body:JSON.stringify({action:'requestRecovery',email:'unknown@example.test'})});const response=await route({request:req,env});assert.equal(response.status,200);assert.equal((await response.text()).includes('unknown@example.test'),false);
 assert.equal(takeAccountEmailAction({hash:'#account=recovery&token=oops',pathname:'/',search:''},{replaceState(){}}).invalid,true);
 assert.equal(takeAccountEmailAction({hash:'#unrelated',pathname:'/',search:''},{replaceState(){throw Error('unrelated fragment must stay');}}),null);
 let invitationPath='';assert.equal(takeInvitationToken({hash:'#invite='+('a'.repeat(48)),pathname:'/',search:'?other=value'},{replaceState(a,b,path){invitationPath=path;}}),'a'.repeat(48));assert.equal(invitationPath,'/?other=value');
 assert.equal(takeInvitationToken({hash:'',pathname:'/',search:'?invite='+('b'.repeat(48))},{replaceState(a,b,path){invitationPath=path;}}),'b'.repeat(48));assert.equal(invitationPath,'/');
}finally{globalThis.fetch=oldFetch;}
console.log('Account email passed: disabled delivery gates, pinned origin/recipient, capped and coalesced sends, random recovery provisioning, digest-only storage, fragment cleanup, explicit one-time verification, strong replacement, verified retirement, session revocation on uncertain update and generic unknown-account replies. No emails or external services.');
