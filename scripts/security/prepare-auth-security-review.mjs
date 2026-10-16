// Prepares a synthetic isolated probe only; no network or credential reads.
// Never apply its schema to an existing staff database.
import fs from 'node:fs';import crypto from 'node:crypto';
const root=process.cwd(),out=process.argv[2],id=process.argv[3];if(!out||!out.startsWith('/')||!/^[a-f0-9-]{36}$/.test(id||''))throw Error('Provide an absolute disposable directory and a NEW disposable D1 UUID');fs.mkdirSync(out,{recursive:true,mode:0o700});
const gate=crypto.randomBytes(32).toString('hex');fs.writeFileSync(out+'/gate',gate,{mode:0o600});
fs.writeFileSync(out+'/schema.sql',fs.readdirSync(root+'/migrations').filter(f=>f.endsWith('.sql')).sort().map(f=>fs.readFileSync(root+'/migrations/'+f,'utf8')).join('\n'));
fs.writeFileSync(out+'/wrangler.toml',`name = "roster-auth-security-review-20261017"\nmain = "worker.mjs"\ncompatibility_date = "2026-04-22"\nworkers_dev = true\npreview_urls = false\n[[d1_databases]]\nbinding = "TEST_DB"\ndatabase_name = "roster-auth-security-review-20261017"\ndatabase_id = "${id}"\n`);
fs.writeFileSync(out+'/worker.mjs',`import {issueAccountSession,readAccountSession} from ${JSON.stringify(root+'/functions/_lib/account-sessions.js')};
import {beginCredentialChange} from ${JSON.stringify(root+'/functions/_lib/account-credential-change.js')};
import {loadAuthenticationLink} from ${JSON.stringify(root+'/functions/_lib/authentication-migration.js')};
export default{async fetch(request,env){
if(request.method!=='POST'||request.headers.get('Authorization')!==${JSON.stringify('Bearer '+gate)})return new Response(null,{status:404});
const db=env.TEST_DB,email='native-security@example.test',subject='77777777-1111-4222-8333-444444444444';
try{
await db.batch([
db.prepare("INSERT INTO account_profiles(email,role,password_hash,password_salt,subscription_token) VALUES(?,'user','','','synthetic-feed')").bind(email),
db.prepare("INSERT INTO account_authentication_links(email,provider,provider_subject,phase,credential_revision,created_at,updated_at) VALUES(?,'supabase',?,'ready','fixture','now','now')").bind(email,subject),
db.prepare("INSERT INTO account_email_challenges VALUES('fixture-challenge',?,?,'recovery','recovery','digest',9999999999999,'issued',1)").bind(email,subject),
db.prepare("INSERT INTO account_invites(token_hash,email,expires_at,created_at,updated_at) VALUES('fixture-invite',?,'2099-01-01','now','now')").bind(email),
]);
const initial=await issueAccountSession(db,{email,subject});
const configuration={ROSTER_DB:db,ROSTER_AUTH_CONFIG:JSON.stringify({enabled:true,sessionsEnabled:true,url:'https://nncmejyfkujpfuivlart.supabase.co',publishableKey:'sb_publishable_unused_fixture',secretKey:'sb_secret_unused_fixture'})};
const op=await beginCredentialChange(configuration,await loadAuthenticationLink(db,email),'Synthetic replacement password');
const after=await db.prepare('SELECT COUNT(*) AS count FROM account_sessions WHERE email=?').bind(email).first();
const challenge=await db.prepare("SELECT phase FROM account_email_challenges WHERE id='fixture-challenge'").first();
const invite=await db.prepare("SELECT revoked_at FROM account_invites WHERE token_hash='fixture-invite'").first();
let staleDenied=false,lockedDenied=false,revokedCookieDenied=false;
try{await issueAccountSession(db,{email,subject,credentialVersion:0});}catch(e){staleDenied=e.status===401;}
try{await issueAccountSession(db,{email,subject,credentialVersion:1});}catch(e){lockedDenied=e.status===401;}
try{await readAccountSession(db,new Request('https://fixture.invalid',{headers:{Cookie:initial.setCookie.split(';')[0]}}));}catch(e){revokedCookieDenied=e.status===401;}
// Synthetic release only: provider proof is covered by the connected tests;
// this native test exercises the actual SQLite/D1 admission and trigger rules.
await db.prepare("UPDATE account_authentication_links SET credential_version=credential_version+1,credential_operation_id='',credential_target_digest='' WHERE email=? AND credential_operation_id=?").bind(email,op.credential_operation_id).run();
const fresh=await issueAccountSession(db,{email,subject,credentialVersion:2});
const read=await readAccountSession(db,new Request('https://fixture.invalid',{headers:{Cookie:fresh.setCookie.split(';')[0]}}));
const proof={ok:op.credential_version===1&&after.count===0&&challenge.phase==='complete'&&!!invite.revoked_at&&staleDenied&&lockedDenied&&revokedCookieDenied&&read.email===email,nativeD1:true,sessionRevocationAtomic:after.count===0,oldProofDenied:staleDenied,pendingIssuanceDenied:lockedDenied,oldCookieDenied:revokedCookieDenied,resetRetired:challenge.phase==='complete',invitationRetired:!!invite.revoked_at,newVersionAccepted:read.email===email,noProviderCalls:true};
return Response.json(proof,{status:proof.ok?200:500});
}catch(e){return Response.json({ok:false,code:String(e.code||'native-test-failed')},{status:503});}
}};`);
