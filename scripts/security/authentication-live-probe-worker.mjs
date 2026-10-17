// Disposable synthetic acceptance only. Deploy through the preparer, never as
// an application route. No staff database or mail service is bound.
import {onRequestPost as authRoute} from '../../functions/api/auth/session.js';
import {onRequestPost as stateRoute} from '../../functions/api/state.js';
import {deleteSupabaseAccount} from '../../functions/_lib/supabase-auth.js';

export function authenticationLiveProbe(gate,fixtureId){return{async fetch(request,env){
 if(request.method!=='POST'||request.headers.get('Authorization')!=='Bearer '+gate)return new Response(null,{status:404});
 const supplied=JSON.parse(env.ROSTER_AUTH_CONFIG||'{}');
 // The existing secret may also be pasted directly into an encrypted binding
 // without copying a whole credential-bearing JSON document through tooling.
 if(env.SUPABASE_SECRET_KEY)supplied.secretKey=env.SUPABASE_SECRET_KEY;
 if(supplied.enabled!==false)return Response.json({ok:false,code:'configuration-must-stay-disabled'},{status:503});
 const meter={statements:0,rowsRead:0,rowsWritten:0};
 function record(r){meter.statements++;meter.rowsRead+=Number(r.meta?.rows_read||0);meter.rowsWritten+=Number(r.meta?.rows_written||0);return r;}
 const db={prepare(q){return{stmt:env.TEST_DB.prepare(q),bind(...a){this.stmt=this.stmt.bind(...a);return this;},async all(){return record(await this.stmt.all());},async run(){return record(await this.stmt.run());},async first(name){const r=await this.all();return name?r.results?.[0]?.[name]??null:r.results?.[0]||null;}};},async batch(ss){return(await env.TEST_DB.batch(ss.map(s=>s.stmt))).map(record);}};
 const config={...supplied,enabled:true,sessionsEnabled:true,migrationEnabled:true,signupEnabled:true,emailDeliveryReady:false};
 const environment={ROSTER_DB:db,ROSTER_AUTH_CONFIG:JSON.stringify(config),AUTH_LOGIN_LIMITER:{async limit(){return{success:true};}},ROSTER_FILES:{async get(){return null;}},ROSTER_CACHE:{async get(){return null;}},IDENTITY_DISCOVERY_ENABLED:'false',CREATOR_DIRECTORY_ENABLED:'false',CREATOR_STARTUP_HYDRATION_ENABLED:'false'};
 const url=new URL(request.url),step=url.pathname.slice(1),kind=url.searchParams.get('kind')==='legacy'?'legacy':'new';
 const email='batch6-'+fixtureId+'-'+kind+'@example.test';
 const password='Synthetic '+gate.slice(0,28),replacement='Replacement '+gate.slice(28);
 const cookie=request.headers.get('Cookie')||'';
 try{
  if(step==='seed'){
   const salt=fixtureId.replaceAll('-','').slice(0,32);
   const digest=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(salt+':'+password));
   const hash=Array.from(new Uint8Array(digest),b=>b.toString(16).padStart(2,'0')).join('');
   await db.prepare("INSERT INTO account_profiles(email,real_name,role,password_hash,password_salt,subscription_token) VALUES(?,'Synthetic Legacy Doctor','user',?,?,'synthetic-preserved-feed')").bind(email,hash,salt).run();
   return Response.json({ok:true,meter});
  }
  if(step==='cleanup'){
   let removed=0;
   for(const type of ['legacy','new']){
    const address='batch6-'+fixtureId+'-'+type+'@example.test';
    const link=await db.prepare('SELECT provider_subject FROM account_authentication_links WHERE email=?').bind(address).first();
    if(link){await deleteSupabaseAccount(environment,{subject:link.provider_subject,email:address});removed++;}
   }
   return Response.json({ok:true,syntheticAccountsRemoved:removed,meter});
  }
  const isState=step==='context';let body;
  if(step==='signup')body={action:'signup',email,password,realName:'Synthetic New Doctor',remember:true};
  else if(step==='login'||step==='oldPassword')body={action:'login',email,password,remember:true};
  else if(step==='returning')body={action:'login',email,password:replacement,remember:true};
  else if(step==='change')body={action:'changePassword',currentPassword:password,newPassword:replacement,expectedAccount:email};
  else if(step==='logout')body={action:'logout',expectedAccount:email};
  else if(isState)body={action:'loadAccountContext',expectedAccount:email,responseMode:'fast'};
  else return new Response(null,{status:404});
  const internal=new Request(url.origin+(isState?'/api/state':'/api/auth/session'),{method:'POST',headers:{Origin:url.origin,'Content-Type':'application/json',...(cookie?{Cookie:cookie}:{})},body:JSON.stringify(body)});
  const response=await(isState?stateRoute:authRoute)({request:internal,env:environment,data:{stateAction:body.action},waitUntil(){}});
  const data=await response.json();
  const profile=await db.prepare('SELECT password_hash,password_salt,subscription_token FROM account_profiles WHERE email=?').bind(email).first();
  const link=await db.prepare('SELECT phase,credential_operation_id,credential_version FROM account_authentication_links WHERE email=?').bind(email).first();
  // Output only acceptance booleans and aggregate costs, never provider replies,
  // tokens, fixture addresses, passwords or private calendar envelopes.
  return Response.json({ok:response.ok,status:response.status,code:typeof data.code==='string'?data.code:null,legacyHashRetired:!!profile&&!profile.password_hash&&!profile.password_salt,identityReady:link?.phase==='ready'&&!link.credential_operation_id,legacyFeedPreserved:kind!=='legacy'||profile?.subscription_token==='synthetic-preserved-feed',meter},{status:response.status,headers:{'Cache-Control':'no-store',...(response.headers.has('Set-Cookie')?{'Set-Cookie':response.headers.get('Set-Cookie')}:{})}});
 }catch(e){return Response.json({ok:false,code:String(e.code||'native-acceptance-failed'),meter},{status:503});}
}};}
