import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {readFile,readdir} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {acceptAccountInvitation} from '../functions/_lib/account-invitations.js';
import {finishAccountAuthenticationDeletion} from '../functions/_lib/account-auth-deletions.js';
import {authenticateExistingAccount} from '../functions/_lib/account-authentication.js';
import {AUTH_PROJECT_ORIGIN} from '../functions/_lib/supabase-auth.js';
const sql=new DatabaseSync(':memory:');sql.exec('PRAGMA foreign_keys=ON');
for(const name of (await readdir(new URL('../migrations/',import.meta.url))).filter(n=>n.endsWith('.sql')).sort())sql.exec(await readFile(new URL('../migrations/'+name,import.meta.url),'utf8'));
const db={prepare(q){return{args:[],bind(...a){this.args=a;return this;},async first(){return sql.prepare(q).get(...this.args)||null;},async run(){const r=sql.prepare(q).run(...this.args);return{meta:{changes:Number(r.changes)}};}};},async batch(ss){sql.exec('BEGIN');try{const r=[];for(const s of ss)r.push(await s.run());sql.exec('COMMIT');return r;}catch(e){sql.exec('ROLLBACK');throw e;}}};
const email='invite@example.test',salt='a'.repeat(32),hash='b'.repeat(64),token='disposable-invitation-token-with-entropy',tokenHash=createHash('sha256').update(token).digest('hex'),newPassword='New invited fixture password';
sql.prepare("INSERT INTO account_profiles(email,real_name,password_hash,password_salt,role,subscription_token) VALUES(?,'Fixture',?,?,'user','preserved-feed')").run(email,hash,salt);
sql.prepare("INSERT INTO account_invites(token_hash,email,expires_at,created_at,updated_at) VALUES(?,?,?,'now','now')").run(tokenHash,email,new Date(Date.now()+86400000).toISOString());
const env={ROSTER_DB:db,ROSTER_AUTH_CONFIG:JSON.stringify({enabled:true,url:AUTH_PROJECT_ORIGIN,publishableKey:'sb_publishable_fixture',secretKey:'sb_secret_fixture'})};
const oldFetch=globalThis.fetch;let user=null,password='',calls=0,failDelete=false;
globalThis.fetch=async(url,init)=>{calls++;const body=init.body?JSON.parse(init.body):{};
 if(url.includes('/admin/users/')&&init.method==='GET')return user?Response.json(user):Response.json({error:'missing'},{status:404});
 if(url.endsWith('/admin/users')&&init.method==='POST'){user={email:body.email,id:body.id,email_confirmed_at:null};password=body.password;return Response.json(user);}
 if(url.includes('/admin/users/')&&init.method==='PUT'){assert.equal(body.email_confirm,true);assert.equal(sql.prepare('SELECT accepted_at FROM account_invites WHERE token_hash=?').get(tokenHash).accepted_at!=='',true,'invitation proof is consumed before privileged password mutation');user.email_confirmed_at=new Date().toISOString();password=body.password;return Response.json(user);}
 if(url.endsWith('/token?grant_type=password'))return body.password===password?Response.json({user,access_token:'private-fixture',refresh_token:'private-refresh',expires_in:3600}):Response.json({error:'invalid'},{status:400});
 if(url.includes('/admin/users/')&&init.method==='DELETE'){if(failDelete)throw Error('synthetic provider outage');user=null;return new Response(null,{status:204});}
 throw Error('Unexpected request');
};
try{
 await assert.rejects(()=>acceptAccountInvitation(env,{inviteToken:'wrong-token-that-is-long-enough',newPassword}),e=>e.code==='auth-invite');assert.equal(calls,0);
 await assert.rejects(()=>acceptAccountInvitation(env,{inviteToken:token,newPassword:'short'}),e=>e.code==='auth-invite');assert.equal(calls,0);
 const result=await acceptAccountInvitation(env,{inviteToken:token,newPassword});assert.equal(result.email,email);
 assert.equal(sql.prepare('SELECT password_hash FROM account_profiles WHERE email=?').get(email).password_hash,'');
 assert.equal(sql.prepare('SELECT phase FROM account_authentication_links WHERE email=?').get(email).phase,'ready');
 assert.equal(sql.prepare('SELECT subscription_token FROM account_profiles WHERE email=?').get(email).subscription_token,'preserved-feed');
 const beforeReplay=calls;await assert.rejects(()=>acceptAccountInvitation(env,{inviteToken:token,newPassword:'A different attempted password'}),e=>e.code==='auth-invite');assert.equal(calls,beforeReplay,'consumed invitation cannot change the provider again');
 assert.equal((await authenticateExistingAccount(env,{email,password:newPassword})).subject,user.id);
 const subject=user.id;
 assert.equal((await finishAccountAuthenticationDeletion(env,email)).pending,false,'no provider deletion before a job exists');
 // Deletion records survive the profile/link cascade and immediately revoke cookies.
 sql.prepare('INSERT INTO account_sessions VALUES(?,?,?,?,?,?,0)').run('a'.repeat(64),email,Date.now(),Date.now()+100000,Date.now()+200000,Date.now()+10000);
 sql.prepare('DELETE FROM account_profiles WHERE email=?').run(email);
 assert.equal(sql.prepare('SELECT COUNT(*) AS n FROM account_sessions').get().n,0);
 assert.equal(sql.prepare('SELECT provider_subject FROM account_auth_deletions WHERE email=?').get(email).provider_subject,subject);
 await assert.rejects(()=>authenticateExistingAccount(env,{email,password:newPassword}),e=>e.status===401,'queued provider deletion cannot be signed into');
 failDelete=true;assert.equal((await finishAccountAuthenticationDeletion(env,email)).pending,true);assert.equal(sql.prepare('SELECT phase FROM account_auth_deletions WHERE email=?').get(email).phase,'queued');
 failDelete=false;assert.equal((await finishAccountAuthenticationDeletion(env,email)).pending,false);assert.equal(user,null);assert.equal(sql.prepare('SELECT phase FROM account_auth_deletions WHERE email=?').get(email).phase,'complete');
 const beforeIdempotent=calls;assert.equal((await finishAccountAuthenticationDeletion(env,email)).pending,false);assert.equal(calls,beforeIdempotent);
 // Provider deletion failure is bounded and never leaks an exception/token.
 sql.prepare("INSERT INTO account_auth_deletions VALUES('aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee','other@example.test','queued',5,'now','now')").run();
 const beforeCap=calls;assert.equal((await finishAccountAuthenticationDeletion(env,'other@example.test')).pending,true);assert.equal(calls,beforeCap);
}finally{globalThis.fetch=oldFetch;}
console.log('Account lifecycle passed: invitation email proof, consume-before-mutate and replay denial, verified password retirement, preserved settings/feed, immediate cookie revocation, durable exact-subject deletion, safe retry and bounded attempts. No external services.');
