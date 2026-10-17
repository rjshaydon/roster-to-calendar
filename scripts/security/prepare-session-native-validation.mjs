// Prepares only: no network, deployment or credential reads.
// Use a newly created empty roster-session-native-validation D1 database.
import fs from 'node:fs';import crypto from 'node:crypto';
const out=process.argv[2],databaseId=process.argv[3];if(!out||!out.startsWith('/')||!/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i.test(databaseId||''))throw Error('Provide an absolute temporary output directory and the UUID of a newly created disposable database.');fs.mkdirSync(out,{recursive:true,mode:0o700});
const gate=crypto.randomBytes(32).toString('hex'),root=process.cwd();
fs.writeFileSync(out+'/gate',gate,{mode:0o600});
const migrations=fs.readdirSync(root+'/migrations').filter(x=>x.endsWith('.sql')).sort();
let schema=migrations.map(name=>fs.readFileSync(root+'/migrations/'+name,'utf8')).join('\n');
schema+=`\nWITH RECURSIVE n(x) AS(SELECT 0 UNION ALL SELECT x+1 FROM n WHERE x<499)
INSERT INTO account_profiles(email,real_name,role,subscription_token,password_hash,password_salt,created_at,updated_at)
SELECT 'session-fixture-'||x||'@example.test','Synthetic Fixture '||x,CASE WHEN x=0 THEN 'creator' ELSE 'user' END,'synthetic-feed-'||x,'','','2026-10-16','2026-10-16' FROM n;
WITH RECURSIVE n(x) AS(SELECT 0 UNION ALL SELECT x+1 FROM n WHERE x<499)
INSERT INTO account_authentication_links(email,provider,provider_subject,phase,credential_revision,created_at,updated_at)
SELECT 'session-fixture-'||x||'@example.test','supabase',printf('00000000-0000-4000-8000-%012d',x),'ready','synthetic-fixture','2026-10-16','2026-10-16' FROM n;\n`;
fs.writeFileSync(out+'/schema.sql',schema);
fs.writeFileSync(out+'/worker.mjs',`import {issueAccountSession} from ${JSON.stringify(root+'/functions/_lib/account-sessions.js')};
import {onRequestPost as stateHandler} from ${JSON.stringify(root+'/functions/api/state.js')};
import {onRequestPost as sessionHandler} from ${JSON.stringify(root+'/functions/api/auth/session.js')};
import {onRequest as middleware} from ${JSON.stringify(root+'/functions/_middleware.js')};
export default {async fetch(request,env,ctx){
if(request.method!=='POST'||request.headers.get('authorization')!=='Bearer ${gate}')return new Response(null,{status:404});
const url=new URL(request.url),mode=url.pathname.slice(1),i=Number(url.searchParams.get('i')||0);
if(!Number.isInteger(i)||i<0||i>499)return new Response(null,{status:400});
const meter={statements:0,rowsRead:0,rowsWritten:0};function record(r){meter.statements++;meter.rowsRead+=Number(r.meta?.rows_read||0);meter.rowsWritten+=Number(r.meta?.rows_written||0);return r;}
const db={prepare(sql){return{stmt:env.TEST_DB.prepare(sql),bind(...args){this.stmt=this.stmt.bind(...args);return this;},async all(){return record(await this.stmt.all());},async run(){return record(await this.stmt.run());},async first(name){const r=await this.all();return name?r.results?.[0]?.[name]??null:r.results?.[0]||null;}};},async batch(ss){const rr=await env.TEST_DB.batch(ss.map(s=>s.stmt));rr.forEach(record);return rr;}};
const environment={ROSTER_DB:db,ROSTER_AUTH_CONFIG:JSON.stringify({enabled:true,sessionsEnabled:true,url:'https://nncmejyfkujpfuivlart.supabase.co',publishableKey:'sb_publishable_unused_fixture',secretKey:'sb_secret_unused_fixture'}),CREATOR_STARTUP_HYDRATION_ENABLED:'false',CREATOR_DIRECTORY_ENABLED:'false',IDENTITY_REVIEW_ENABLED:'false',IDENTITY_DISCOVERY_ENABLED:'false',FACILITY_OVERVIEW_MAINTENANCE_MODE:'true',ROSTER_FILES:{async get(){return null;}},ROSTER_CACHE:{async get(){return null;}}};
try{
if(mode==='issue'){const issued=await issueAccountSession(db,{email:'session-fixture-'+i+'@example.test',subject:'00000000-0000-4000-8000-'+String(i).padStart(12,'0'),persistent:true});return Response.json({ok:true,meter},{headers:{'Set-Cookie':issued.setCookie,'Cache-Control':'no-store'}});}
const state=mode==='state'||mode==='forged',action=state?'loadAccountContext':mode;
const headers={'Content-Type':'application/json',Origin:'https://session-test.invalid','Cookie':request.headers.get('Cookie')||''};
const internal=new Request('https://session-test.invalid/api/'+(state?'state':'auth/session'),{method:'POST',headers,body:JSON.stringify({action,responseMode:'fast',...(mode==='forged'?{email:'session-fixture-499@example.test'}:{})})});
const context={request:internal,env:environment,data:{},waitUntil:p=>ctx.waitUntil(p),next:()=>state?stateHandler(context):sessionHandler(context)};
const response=await middleware(context);await response.body?.cancel();
return Response.json({ok:response.ok,status:response.status,meter},{headers:{'Cache-Control':'no-store',...(response.headers.get('Set-Cookie')?{'Set-Cookie':response.headers.get('Set-Cookie')}:{})}});
}catch{return Response.json({ok:false,code:'native-probe-failed',meter},{status:503});}
}};`);
fs.writeFileSync(out+'/wrangler.toml','name = "roster-session-native-validation"\nmain = "worker.mjs"\ncompatibility_date = "2026-04-22"\nworkers_dev = true\npreview_urls = false\n[[d1_databases]]\nbinding = "TEST_DB"\ndatabase_name = "roster-session-native-validation"\ndatabase_id = "'+databaseId+'"\n');
