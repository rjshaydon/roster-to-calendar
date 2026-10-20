// Builds on the synthetic storage probe. Only an explicitly supplied disposable
// D1 is bound; no source bucket, provider secret or production account is used.
import fs from 'node:fs';import {execFileSync} from 'node:child_process';
const [out,id]=process.argv.slice(2);if(!out?.startsWith('/')||!/^[a-f0-9-]{36}$/.test(id||''))throw Error('Absolute disposable directory and database UUID required.');
execFileSync(process.execPath,['scripts/prepare-staff-history-native.mjs',out],{stdio:'ignore'});
const root=process.cwd();let source=fs.readFileSync(out+'/worker.mjs','utf8');
source=`import {onRequestPost as stateHandler} from ${JSON.stringify(root+'/functions/api/state.js')};\nimport {issueAccountSession} from ${JSON.stringify(root+'/functions/_lib/account-sessions.js')};\nimport {onRequest as middleware} from ${JSON.stringify(root+'/functions/_middleware.js')};\n`+source;
source=source.replace('async fetch(request){','async fetch(request,env,ctx){');
source=source.replace(' try{\n let result;',` try{
 if(['issue','issue-sms','issue-trainee','issue-disabled'].includes(mode)){
  const issued=await issueAccountSession(env.TEST_DB,{email:mode==='issue'?'staff-history-probe@example.test':'staff-history-'+mode.slice(6)+'@example.test',subject:'00000000-0000-4000-8000-'+({issue:'000000000001','issue-sms':'000000000002','issue-trainee':'000000000003','issue-disabled':'000000000004'}[mode]),persistent:true});
  return Response.json({ok:true},{headers:{'Set-Cookie':issued.setCookie}});
 }
 if(['full-names','full-histories'].includes(mode)){
  const input={action:'queryFacilityStaffHistory',targetEmail:request.headers.get('X-Test-Target')||undefined,mode:mode==='full-names'?'names':'histories',...(mode==='full-names'?{page:0}:{revision:'${'a'.repeat(64)}',slots:[0,1,2,3,4,5,6,7,8,9,10,11]})};
  const environment={ROSTER_DB:env.TEST_DB,ROSTER_FILES:r2,ROSTER_CACHE:{async get(){return null;}},
   ROSTER_AUTH_CONFIG:JSON.stringify({enabled:true,sessionsEnabled:true,url:'https://synthetic.supabase.co',publishableKey:'sb_publishable_unused_fixture',secretKey:'sb_secret_unused_fixture'}),
   FACILITY_OVERVIEW_MAINTENANCE_MODE:'false',FACILITY_SHARED_ROLLOUT_ACTIVE:'true',FACILITY_SHARED_EMERGENCY_PAUSED:'false',FACILITY_LEGACY_READS_PAUSED:'true',FACILITY_SHARED_READER_COHORT:'all',FACILITY_SHARED_READER_SOURCE_ALLOWLIST:'mmc,mch,ddh,vhh',FACILITY_SHARED_METADATA_ENABLED:'true',FACILITY_SHARED_DAYS_ENABLED:'true',STAFF_HISTORY_ENABLED:'true',CREATOR_DIRECTORY_ENABLED:'false',IDENTITY_REVIEW_ENABLED:'true',IDENTITY_DISCOVERY_ENABLED:'false'};
  const internal=new Request('https://synthetic.invalid/api/state',{method:'POST',headers:{Origin:'https://synthetic.invalid','Content-Type':'application/json',Cookie:request.headers.get('Cookie')||''},body:JSON.stringify(input)});
  const context={request:internal,env:environment,data:{},waitUntil:p=>ctx.waitUntil(p),next:()=>stateHandler(context)};
  const response=await middleware(context);
  const meter=context.data.d1Meter;
  const headers=new Headers(response.headers);headers.set('X-Test-Meter',JSON.stringify({statements:meter.statementCount,rowsRead:meter.rowsRead,rowsWritten:meter.rowsWritten,complete:meter.metadataComplete}));
  return new Response(response.body,{status:response.status,headers,encodeBody:headers.has('Content-Encoding')?'manual':'automatic'});
 }
 let result;`);
fs.writeFileSync(out+'/worker.mjs',source);
fs.writeFileSync(out+'/wrangler.toml',`name = "roster-staff-history-auth-native-20261020"\nmain = "worker.mjs"\ncompatibility_date = "2026-04-22"\nworkers_dev = true\npreview_urls = false\n[observability]\nenabled = true\nhead_sampling_rate = 1\n[[d1_databases]]\nbinding = "TEST_DB"\ndatabase_name = "roster-staff-history-native-20261020"\ndatabase_id = "${id}"\n`);
let schema=fs.readdirSync('migrations').filter(n=>n.endsWith('.sql')).sort().map(n=>fs.readFileSync('migrations/'+n,'utf8')).join('\n');
schema+=`\nINSERT INTO account_profiles(email,real_name,role,facility_overview_enabled,password_hash,password_salt,created_at,updated_at)VALUES('staff-history-probe@example.test','Synthetic Staff History Probe','creator',1,'','','2026-10-20','2026-10-20');
INSERT INTO account_authentication_links(email,provider,provider_subject,phase,credential_revision,created_at,updated_at)VALUES('staff-history-probe@example.test','supabase','00000000-0000-4000-8000-000000000001','ready','synthetic','2026-10-20','2026-10-20');\n`;
fs.writeFileSync(out+'/schema.sql',schema);
console.log('Prepared full authenticated synthetic search validation.');

fs.writeFileSync(out+'/clinical-fixtures.sql',`CREATE INDEX IF NOT EXISTS idx_roster_events_file_doctor ON roster_events(file_id,doctor_key);
INSERT OR IGNORE INTO account_profiles(email,real_name,role,facility_overview_enabled,password_hash,password_salt)VALUES('staff-history-sms@example.test','Synthetic SMS','user',1,'',''),('staff-history-trainee@example.test','Synthetic Trainee','user',1,'',''),('staff-history-disabled@example.test','Synthetic Disabled','user',0,'','');
INSERT OR IGNORE INTO account_authentication_links(email,provider,provider_subject,phase,credential_revision,created_at,updated_at)VALUES('staff-history-sms@example.test','supabase','00000000-0000-4000-8000-000000000002','ready','synthetic','2026-10-20','2026-10-20'),('staff-history-trainee@example.test','supabase','00000000-0000-4000-8000-000000000003','ready','synthetic','2026-10-20','2026-10-20'),('staff-history-disabled@example.test','supabase','00000000-0000-4000-8000-000000000004','ready','synthetic','2026-10-20','2026-10-20');
INSERT OR IGNORE INTO account_claims(email,source_type,doctor_key,display_name)VALUES('staff-history-sms@example.test','mmc','SYNTHETIC SMS','Synthetic SMS'),('staff-history-trainee@example.test','mmc','SYNTHETIC TRAINEE','Synthetic Trainee');
INSERT OR IGNORE INTO roster_files(id,name,source_type,active)VALUES('synthetic-staff-history','Synthetic.xlsx','mmc',1);
INSERT OR IGNORE INTO facility_term_staff_contributions(source_type,term_start,doctor_key,file_id,display_name,seniority)VALUES('mmc','2026-08-03','SYNTHETIC SMS','synthetic-staff-history','Synthetic SMS','SMS'),('mmc','2026-08-03','SYNTHETIC TRAINEE','synthetic-staff-history','Synthetic Trainee','HMO');
`);
