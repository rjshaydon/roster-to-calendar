import {guardedFetch} from '../_lib/outbound-network.js';
import {reserveRosterMaintenanceBudget} from '../_lib/roster-maintenance-budget.js';
import {requestIdentityPublication} from '../_lib/identity-publication-dispatch.js';
import {authenticationSettings} from '../_lib/authentication-settings.js';
import {AuthenticationError} from '../_lib/supabase-auth.js';
import {boundedBody,admit} from '../_lib/authentication-request.js';
import {accountSessionsEnabled,readAccountSession,requireSameOrigin,requireExpectedAccount,requireRecentAuthentication} from '../_lib/account-sessions.js';
import {requestAdmin,requestSummary,inspectSupportCreation,supportMutationOutcome,createSupportRequest,listSupportRequests,readSupportRequest,markSupportRead,replySupportRequest,setSupportStatus,supportIdentityEvidence,previewSupportResolution,commitSupportResolution,inviteSupportHolder,legacySupportReports,convertLegacyReport,attachSupportIdentity} from '../_lib/support-requests.js';
import {melbourneDateKey} from '../../public/static/roster-term-policy.js';

function reply(body,status=200){return Response.json(body,{status,headers:{'Cache-Control':'no-store','Vary':'Cookie'}});}
export async function onRequestPost(context){
 const {env,request}=context;
 try{
  let settings;try{settings=authenticationSettings(env);}catch{settings={};}
  if(!accountSessionsEnabled(env)||settings.requestsEnabled!==true)return reply({code:'requests-disabled',error:'Messages and requests are not enabled.'},503);
  requireSameOrigin(request);const input=context.data?.supportRequestInput??await boundedBody(request,16384),db=env.ROSTER_DB,session=await readAccountSession(db,request);
  requireExpectedAccount(session,input.expectedAccount);
  if(settings.requestsAdminOnly===true&&!requestAdmin(session))throw new AuthenticationError('requests-disabled','Messages and requests are not enabled for this account.',403);
  const action=input.action;
  if(action==='prepare'){
   if(env.ROSTER_ACCOUNT_BUDGET_ENABLED!=='true')return reply({ok:true,preparedUntil:Date.now()+30000});
   await admit(env,session.email,'support');
   const now=new Date(),until=Date.now()+30000;
   const grant=await db.prepare("SELECT valid_until FROM roster_account_budget WHERE utc_day=? AND valid_until>? AND stop_reason=''").bind(now.toISOString().slice(0,10),new Date(until).toISOString()).first();
   if(grant)return reply({ok:true,preparedUntil:until});
   // A quiet roster needs no maintenance grant. Refresh one only for an
   // explicit message review, on the existing separate admission endpoint.
   const token=String(env.ROSTER_AUTOMATION_TOKEN||'');
   if(!token)return reply({code:'requests-budget',error:'The daily database check is unavailable. Your draft is preserved.'},503);
   const response=await guardedFetch(env,new URL('/api/automation/account-budget',request.url),{method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify({reuseCurrentGrant:true}),signal:AbortSignal.timeout(15000)},{label:'Request daily admission'});
   const admission=await response.json().catch(()=>null);
   if(!response.ok||admission?.ok!==true||admission.deferred)return reply({code:'requests-budget',error:'The daily database safeguard has deferred sending. Your draft is preserved; try again later.'},503);
   return reply({ok:true,preparedUntil:Date.now()+30000});
  }
  if(action==='summary')return reply({ok:true,summary:await requestSummary(db,session.email),admin:requestAdmin(session),transfersEnabled:settings.requestTransfersEnabled===true});
  if(action==='outcome')return reply(await supportMutationOutcome(db,session,input));
  if(action==='list')return reply(await listSupportRequests(db,session,input));
  if(action==='detail')return reply(await readSupportRequest(db,session,input));
  if(action==='read')return reply({...await markSupportRead(db,session,input,{async admitWrite(){await admit(env,session.email,'support');return env.ROSTER_ACCOUNT_BUDGET_ENABLED!=='true'||await reserveRosterMaintenanceBudget(db,128,256);}}),summary:await requestSummary(db,session.email)});
  if(action==='evidence')return reply(await supportIdentityEvidence(db,session,input));
  if(action==='terms'){
   const {identityTermHistory}=await import('../_lib/identity-term-history.js');
   const evidence=await supportIdentityEvidence(db,session,input);return reply({ok:true,...(evidence.identity?await identityTermHistory(env.ROSTER_FILES,evidence.identity.aliases,melbourneDateKey(),input.after||0):{terms:[],next:null})});
  }
  if(action==='previewResolution'){
   if(!requestAdmin(session))throw new AuthenticationError('requests-forbidden','Admin access is required.',403);
   await admit(env,session.email,'support');
   if(env.ROSTER_ACCOUNT_BUDGET_ENABLED==='true'&&!await reserveRosterMaintenanceBudget(db,128,512))return reply({code:'requests-budget',error:'Identity review is temporarily paused by the daily database safeguard.'},503);
   return reply(await previewSupportResolution(db,session,input));
  }
  if(action==='legacy')return reply(await legacySupportReports(db,session,input));
  const mutations={attachIdentity:attachSupportIdentity,create:createSupportRequest,reply:replySupportRequest,status:setSupportStatus,inviteHolder:inviteSupportHolder,convertLegacy:convertLegacyReport,commitResolution:commitSupportResolution};
  if(!mutations[action])throw new AuthenticationError('requests-input','Unknown request action.',400);
  if(action==='commitResolution'){
   if(settings.requestTransfersEnabled!==true)throw new AuthenticationError('requests-review-gate','Identity decisions await the final security review.',403);
   if(!requestAdmin(session))throw new AuthenticationError('requests-forbidden','Admin access is required.',403);
   requireRecentAuthentication(session);
  }
  await admit(env,session.email,'support');
  const creation=action==='create'?await inspectSupportCreation(db,session,input):null;
  // An existing open request is a lookup, not a new message. Keep the new
  // explanation unsent in the browser; no write grant/receipt is necessary.
  if(creation?.ok)return reply({...creation,summary:await requestSummary(db,session.email)});
  if(env.ROSTER_ACCOUNT_BUDGET_ENABLED==='true'&&!await reserveRosterMaintenanceBudget(db,action==='commitResolution'?2048:128,action==='commitResolution'?4096:256))return reply({code:'requests-budget',error:'Requests are temporarily paused by the daily database safeguard. Your calendar remains available.'},503);
  const result=action==='create'?await createSupportRequest(db,session,input,{review:creation}):await mutations[action](db,session,input);
  if(result.operationId&&context.waitUntil)context.waitUntil(requestIdentityPublication(context,result.operationId).then(r=>{if(!r.ok)console.warn('Request identity publication deferred');}).catch(()=>console.warn('Request identity publication deferred')));
  return reply({...result,summary:await requestSummary(db,session.email)});
 }catch(error){
  if(['d1-statement-budget-exceeded','d1-request-deadline-exceeded'].includes(error?.code))throw error;
  if(error instanceof AuthenticationError)return reply({code:error.code,error:error.message},error.status);
  if(error?.code?.startsWith('IDENTITY_'))return reply({code:error.code,error:error.message},409);
  return reply({code:'requests-unavailable',error:'Could not load or update this request. Your calendar is still available.'},503);
 }
}
