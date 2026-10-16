// Prepare an isolated synthetic-only provider probe; never deploy or send email.
import fs from 'node:fs';
import crypto from 'node:crypto';
const output=process.argv[2];if(!output||!output.startsWith('/'))throw Error('Provide an absolute temporary output directory.');fs.mkdirSync(output,{recursive:true,mode:0o700});
const gate=crypto.randomBytes(32).toString('hex'),subject=crypto.randomUUID(),email='batch6-mail-'+subject+'@example.test',password=crypto.randomUUID()+crypto.randomUUID(),replacement=crypto.randomUUID()+crypto.randomUUID();
fs.writeFileSync(output+'/gate',gate,{mode:0o600});
fs.writeFileSync(output+'/worker.mjs',`import {supabaseAuthenticationConfig,createSupabaseMigrationUser,generateSupabaseEmailProof,verifySupabaseEmailProof,signInWithSupabasePassword,changeSupabasePassword,deleteSupabaseAccount,acceptSupabaseInvitation} from ${JSON.stringify(process.cwd()+'/functions/_lib/supabase-auth.js')};
const subject=${JSON.stringify(subject)},email=${JSON.stringify(email)},password=${JSON.stringify(password)},replacement=${JSON.stringify(replacement)};
export default {async fetch(request,env){
if(request.method!=='POST'||request.headers.get('authorization')!=='Bearer ${gate}')return new Response(null,{status:404});
const stage=new URL(request.url).pathname.slice(1);const checks={};
try{const supplied=JSON.parse(env.ROSTER_AUTH_CONFIG||'{}');if(supplied.enabled!==false)throw Error();const test={ROSTER_AUTH_CONFIG:JSON.stringify({...supplied,enabled:true})};supabaseAuthenticationConfig(test);
if(stage==='create'){const x=await createSupabaseMigrationUser(test,{email,subject,password});checks.exactUnconfirmed=x.subject===subject&&!x.emailVerified;}
else if(stage==='confirm'){const link=await generateSupabaseEmailProof(test,{email,subject,type:'magiclink'});const proof=await verifySupabaseEmailProof(test,{email,subject,tokenHash:link.tokenHash,type:link.type});checks.confirmed=proof.subject===subject&&proof.emailVerified;}
else if(stage==='password'){const x=await signInWithSupabasePassword(test,{email,password,expectedSubject:subject});checks.passwordAccepted=x.subject===subject;}
else if(stage==='recover'){const link=await generateSupabaseEmailProof(test,{email,subject,type:'recovery'});const proof=await verifySupabaseEmailProof(test,{email,subject,tokenHash:link.tokenHash,type:link.type});await changeSupabasePassword(test,{email,subject,password:replacement,accessToken:proof.accessToken});checks.replacementAccepted=(await signInWithSupabasePassword(test,{email,password:replacement,expectedSubject:subject})).subject===subject;try{await signInWithSupabasePassword(test,{email,password,expectedSubject:subject});checks.oldPasswordDenied=false;}catch(e){checks.oldPasswordDenied=e.code==='invalid-credentials';}try{await verifySupabaseEmailProof(test,{email,subject,tokenHash:link.tokenHash,type:link.type});checks.reusedProofDenied=false;}catch(e){checks.reusedProofDenied=e.code==='invalid-credentials';}}
else if(stage==='invite'){await acceptSupabaseInvitation(test,{email,subject,password});checks.invitePasswordAccepted=(await signInWithSupabasePassword(test,{email,password,expectedSubject:subject})).subject===subject;}
else if(stage==='cleanup'){await deleteSupabaseAccount(test,{email,subject});checks.fixtureRemoved=true;}
else return new Response(null,{status:404});
const ok=Object.values(checks).length>0&&Object.values(checks).every(x=>x===true);return Response.json({ok,stage,...checks,noEmailSent:true},{status:ok?200:503,headers:{'Cache-Control':'no-store'}});
}catch(e){return Response.json({ok:false,stage,code:['auth-provider-response','auth-provider-unavailable','invalid-credentials','auth-migration-unconfirmed','auth-identity-mismatch','auth-rate-limited'].includes(e.code)?e.code:'probe-failed',noEmailSent:true},{status:503});}
}};`);
fs.writeFileSync(output+'/wrangler.toml','name = "roster-auth-email-validation"\nmain = "worker.mjs"\ncompatibility_date = "2026-10-15"\nworkers_dev = true\npreview_urls = false\n');
