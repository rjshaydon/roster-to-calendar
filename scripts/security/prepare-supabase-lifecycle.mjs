// Prepare only: no deployment, credentials, staff data or emails.
// Usage: node scripts/security/prepare-supabase-lifecycle.mjs <temporary output directory>
import fs from 'node:fs';
import crypto from 'node:crypto';
const output=process.argv[2];if(!output||!output.startsWith('/'))throw Error('Provide an absolute temporary output directory.');
fs.mkdirSync(output,{recursive:true,mode:0o700});
const gate=crypto.randomBytes(32).toString('hex');
fs.writeFileSync(output+'/gate',gate,{mode:0o600});
const modulePath=process.cwd()+'/functions/_lib/supabase-auth.js';
const source=`import {supabaseAuthenticationConfig,createSupabaseMigrationUser,getSupabaseMigrationUser,signInWithSupabasePassword} from ${JSON.stringify(modulePath)};
export default {async fetch(request,env){
if(request.method!=='POST'||new URL(request.url).pathname!=='/run'||request.headers.get('authorization')!=='Bearer ${gate}')return new Response('Not found',{status:404});
let cleanupSubject='',config;const checks={};let stage='configuration';
try{
const supplied=JSON.parse(env.ROSTER_AUTH_CONFIG||'{}');if(supplied.enabled!==false)throw Error();
const testEnv={ROSTER_AUTH_CONFIG:JSON.stringify({...supplied,enabled:true})};config=supabaseAuthenticationConfig(testEnv);
const subject=crypto.randomUUID(),email='batch6-'+subject+'@example.test',password=crypto.randomUUID()+crypto.randomUUID();
async function admin(path,method,body){const r=await fetch(config.url+'/auth/v1/'+path,{method,redirect:'manual',headers:{apikey:config.secretKey,'content-type':'application/json'},...(body?{body:JSON.stringify(body)}:{})});if(!r.ok){await r.body?.cancel();throw Error('provider-response');}return r.json();}
stage='create-unconfirmed';const created=await admin('admin/users','POST',{id:subject,email,password,email_confirm:false});
if(created.email!==email||! /^[a-f0-9-]{36}$/i.test(created.id||''))throw Error();cleanupSubject=created.id;
checks.reservedSubjectPreserved=created.id===subject;checks.initiallyUnconfirmed=!created.email_confirmed_at;if(!checks.reservedSubjectPreserved||!checks.initiallyUnconfirmed)throw Error();
stage='read-exact-user';const found=await getSupabaseMigrationUser(testEnv,{subject,email});checks.exactUserReadable=found.subject===subject&&!found.emailVerified;
stage='unconfirmed-login';try{await signInWithSupabasePassword(testEnv,{email,password,expectedSubject:subject});throw Error('unexpected-login');}catch(e){checks.unconfirmedLoginDenied=e.code==='email-verification-required';}if(!checks.unconfirmedLoginDenied)throw Error();
// Synthetic fixture only: this deliberately does not exercise email delivery.
stage='confirm-synthetic-fixture';await admin('admin/users/'+subject,'PUT',{email_confirm:true});
stage='wrong-password';try{await signInWithSupabasePassword(testEnv,{email,password:crypto.randomUUID(),expectedSubject:subject});throw Error('unexpected-login');}catch(e){checks.wrongPasswordDenied=e.code==='invalid-credentials';}if(!checks.wrongPasswordDenied)throw Error();
stage='correct-password';const signed=await signInWithSupabasePassword(testEnv,{email,password,expectedSubject:subject});checks.correctPasswordAccepted=signed.subject===subject&&signed.emailVerified&&!!signed.accessToken;
stage='subject-mismatch';try{await signInWithSupabasePassword(testEnv,{email,password,expectedSubject:crypto.randomUUID()});throw Error('unexpected-login');}catch(e){checks.wrongSubjectDenied=e.code==='auth-identity-mismatch';}
checks.migrationDisabled=true;checks.noEmailSent=true;
}catch{checks.failedStage=stage;}finally{
if(cleanupSubject&&config){try{const r=await fetch(config.url+'/auth/v1/admin/users/'+cleanupSubject,{method:'DELETE',redirect:'manual',headers:{apikey:config.secretKey}});checks.syntheticAccountRemoved=r.ok;await r.body?.cancel();}catch{checks.syntheticAccountRemoved=false;}}
}
const ok=!checks.failedStage&&Object.values(checks).every(x=>x===true);return Response.json({ok,...checks},{status:ok?200:503,headers:{'Cache-Control':'no-store'}});
}};`;
fs.writeFileSync(output+'/worker.mjs',source);
fs.writeFileSync(output+'/wrangler.toml','name = "roster-auth-lifecycle"\nmain = "worker.mjs"\ncompatibility_date = "2026-10-15"\nworkers_dev = true\npreview_urls = false\n');
