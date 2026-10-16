import assert from 'node:assert/strict';
import { DatabaseSync } from 'node:sqlite';
import { readFile,readdir } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { authenticateExistingAccount } from '../functions/_lib/account-authentication.js';
import { AUTH_PROJECT_ORIGIN } from '../functions/_lib/supabase-auth.js';
const sql=new DatabaseSync(':memory:');sql.exec('PRAGMA foreign_keys=ON');
for(const name of (await readdir(new URL('../migrations/',import.meta.url))).filter(n=>n.endsWith('.sql')).sort())sql.exec(await readFile(new URL('../migrations/'+name,import.meta.url),'utf8'));
const db={prepare(q){return{args:[],bind(...a){this.args=a;return this;},async first(){return sql.prepare(q).get(...this.args)||null;},async run(){const r=sql.prepare(q).run(...this.args);return{meta:{changes:Number(r.changes)}};}};},async batch(ss){sql.exec('BEGIN');try{const r=[];for(const s of ss)r.push(await s.run());sql.exec('COMMIT');return r;}catch(e){sql.exec('ROLLBACK');throw e;}}};
const email='migration@example.test',password='Synthetic fixture only',salt='a'.repeat(32),hash=createHash('sha256').update(salt+':'+password).digest('hex');
sql.prepare("INSERT INTO account_profiles(email,real_name,role,password_hash,password_salt,subscription_token) VALUES(?,'Fixture','user',?,?,'unchanged-feed')").run(email,hash,salt);
const cfg={enabled:true,url:AUTH_PROJECT_ORIGIN,publishableKey:'sb_publishable_fixture',secretKey:'sb_secret_fixture',emailDeliveryReady:false,legacyMigrationUntil:new Date(Date.now()+86400000).toISOString()};
const env={ROSTER_DB:db,ROSTER_AUTH_CONFIG:JSON.stringify(cfg)};
const oldFetch=globalThis.fetch;let user=null,posts=0,fetches=0,loseCreateResponse=true,providerPassword=password;
globalThis.fetch=async(url,init)=>{
 fetches++;
 if(url.includes('/admin/users/')&&init.method==='GET')return user?Response.json(user):Response.json({message:'missing'},{status:404});
 if(url.includes('/admin/users/')&&init.method==='PUT'){const b=JSON.parse(init.body);assert.deepEqual(b,{email_confirm:true});user.email_confirmed_at=new Date().toISOString();return Response.json(user);}
 if(url.endsWith('/admin/users')&&init.method==='POST'){
  posts++;const b=JSON.parse(init.body);assert.equal(b.email_confirm,true);user={id:b.id,email:b.email,email_confirmed_at:new Date().toISOString()};
  if(loseCreateResponse){loseCreateResponse=false;throw Error('simulated lost response');}return Response.json(user);
 }
 if(url.endsWith('/token?grant_type=password')){
  const b=JSON.parse(init.body);if(b.password!==providerPassword)return Response.json({error:'invalid'},{status:400});
  return Response.json({user,access_token:'fixture-access',refresh_token:'fixture-refresh',expires_in:3600});
 }
 throw Error('Unexpected test request');
};
try{
 await assert.rejects(()=>authenticateExistingAccount(env,{email,password:'wrong'}),e=>e.status===401);assert.equal(fetches,0);
 await assert.rejects(()=>authenticateExistingAccount({...env,ROSTER_AUTH_CONFIG:JSON.stringify({...cfg,legacyMigrationUntil:'2020-01-01'})},{email,password}),e=>e.status===401);assert.equal(fetches,0);
 await assert.rejects(()=>authenticateExistingAccount(env,{email,password}),e=>e.code==='auth-provider-unavailable');
 const firstSubject=user.id;
 assert.equal(user.id,firstSubject);assert.equal(posts,1,'retry uses exact reserved UUID without resetting provider credentials');
 assert.equal(sql.prepare('SELECT password_hash FROM account_profiles WHERE email=?').get(email).password_hash,hash,'lost response preserves legacy credential');
 user.email_confirmed_at=null;providerPassword='other';
 await assert.rejects(()=>authenticateExistingAccount(env,{email,password}),e=>e.code==='invalid-credentials');
 assert.equal(sql.prepare('SELECT password_hash FROM account_profiles WHERE email=?').get(email).password_hash,hash,'provider password proof is required before retirement');
 providerPassword=password;const proof=await authenticateExistingAccount(env,{email,password});assert.equal(proof.subject,firstSubject);
 assert.equal(sql.prepare('SELECT phase FROM account_authentication_links WHERE email=?').get(email).phase,'ready');
 const account=sql.prepare('SELECT * FROM account_profiles WHERE email=?').get(email);assert.equal(account.password_hash,'');assert.equal(account.password_salt,'');assert.equal(account.subscription_token,'unchanged-feed');assert.equal(account.role,'user');
 assert.throws(()=>sql.prepare('UPDATE account_profiles SET password_hash=?,password_salt=? WHERE email=?').run(hash,salt,email),/Retired password/,'old clients cannot reintroduce weak hashes');
 assert.equal((await authenticateExistingAccount(env,{email,password})).subject,firstSubject);
 assert.equal(posts,1);await assert.rejects(()=>authenticateExistingAccount(env,{email,password:'wrong'}),e=>e.code==='invalid-credentials');
 user={...user,email:'other@example.test'};await assert.rejects(()=>authenticateExistingAccount(env,{email,password}),e=>e.code==='auth-identity-mismatch');
}finally{globalThis.fetch=oldFetch;}
console.log('Authentication coordinator passed: closed rollout/window gates without email dependency, legacy proof before migration, stable UUID after lost response, no duplicate/reset, exact provider-password proof before retirement, unchanged account/feed and refusal to restore retired hashes. No external services.');
