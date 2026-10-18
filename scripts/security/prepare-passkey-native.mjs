// Synthetic native session/crypto cost probe. No provider credentials, network
// authentication, staff records or production bindings are used.
import fs from 'node:fs';import crypto from 'node:crypto';
const root=process.cwd(),out=process.argv[2],id=process.argv[3];
if(!out?.startsWith('/')||!/^[a-f0-9-]{36}$/.test(id||''))throw Error('Provide an absolute disposable directory and NEW empty database UUID.');
fs.mkdirSync(out,{recursive:true,mode:0o700});const gate=crypto.randomBytes(32).toString('hex');
fs.writeFileSync(out+'/gate',gate,{mode:0o600});
fs.writeFileSync(out+'/schema.sql',fs.readdirSync('migrations').filter(x=>x.endsWith('.sql')).sort().map(x=>fs.readFileSync('migrations/'+x,'utf8')).join('\n'));
fs.writeFileSync(out+'/wrangler.toml',`name = "roster-passkey-native-20261018"\nmain = "worker.mjs"\ncompatibility_date = "2026-04-22"\nworkers_dev = true\npreview_urls = false\n[observability]\nenabled = true\nhead_sampling_rate = 1\n[[d1_databases]]\nbinding = "TEST_DB"\ndatabase_name = "roster-passkey-native-20261018"\ndatabase_id = "${id}"\n`);
fs.writeFileSync(out+'/worker.mjs',`import {issueAccountSession,readAccountSession} from ${JSON.stringify(root+'/functions/_lib/account-sessions.js')};
import {sealPasskeyCookie,readPasskeyCookie,passkeyAccount,grantMatches,PASSKEY_ORIGIN} from ${JSON.stringify(root+'/functions/_lib/account-passkeys.js')};
export default {async fetch(request,bindings){
if(request.method!=='POST'||request.headers.get('Authorization')!==${JSON.stringify('Bearer '+gate)})return new Response(null,{status:404});
const db=bindings.TEST_DB,email='passkey-native@example.test',subject='77777777-1111-4222-8333-444444444444';
const env={ROSTER_DB:db,ROSTER_AUTH_CONFIG:JSON.stringify({enabled:true,sessionsEnabled:true,passkeysEnabled:true,url:'https://nncmejyfkujpfuivlart.supabase.co',publishableKey:'sb_publishable_synthetic',secretKey:'sb_secret_synthetic'})};
const stage=new URL(request.url).pathname;
if(stage==='/seed'){
await db.batch([db.prepare("INSERT INTO account_profiles(email,role,password_hash,password_salt,subscription_token) VALUES(?,'user','','','synthetic-feed')").bind(email),db.prepare("INSERT INTO account_authentication_links(email,provider,provider_subject,phase,credential_revision,created_at,updated_at) VALUES(?,'supabase',?,'ready','revision','now','now')").bind(email,subject)]);return Response.json({ok:true});}
if(stage==='/issue'){
const issued=await issueAccountSession(db,{email,subject});
const grant=await sealPasskeyCookie(env,{kind:'grant',email,subject,version:0,session:issued.tokenHash,accessToken:'synthetic-token-'+'.'.repeat(1100)});
return Response.json({ok:true,cookie:issued.setCookie.split(';')[0]+'; '+grant.split(';')[0]});}
if(stage==='/registration'){
const body=await request.json();const internal=new Request(PASSKEY_ORIGIN+'/api/auth/passkeys',{headers:{Cookie:body.cookie}});
const session=await readAccountSession(db,internal),link=await passkeyAccount(db,session),grant=await readPasskeyCookie(env,internal);
if(!grantMatches(grant,session,link))return Response.json({ok:false},{status:400});
await sealPasskeyCookie(env,{...grant,kind:'registration',challengeId:'cccccccc-1111-4222-8333-dddddddddddd'});
return Response.json({ok:true});}
if(stage==='/status'){
const body=await request.json();const internal=new Request(PASSKEY_ORIGIN+'/api/auth/passkeys',{headers:{Cookie:body.cookie}});
const session=await readAccountSession(db,internal);await passkeyAccount(db,session);
const preference=await db.prepare('SELECT passkey_reminder_dismissed FROM account_profiles WHERE email=?').bind(session.email).first();
return Response.json({ok:true,remindersEnabled:preference.passkey_reminder_dismissed===0});}
return new Response(null,{status:404});
}};`);
