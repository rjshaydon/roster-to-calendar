import {reserveRosterMaintenanceBudget} from '../_lib/roster-maintenance-budget.js';
import {requestIdentityPublication} from '../_lib/identity-publication-dispatch.js';
import {authenticationSettings} from '../_lib/authentication-settings.js';
import {AuthenticationError} from '../_lib/supabase-auth.js';
import {boundedBody,admit} from '../_lib/authentication-request.js';
import {accountSessionsEnabled,readAccountSession,requireSameOrigin,requireExpectedAccount,requireRecentAuthentication} from '../_lib/account-sessions.js';
import {requestAdmin,requestSummary,createSupportRequest,listSupportRequests,readSupportRequest,markSupportRead,replySupportRequest,setSupportStatus,supportIdentityEvidence,previewSupportResolution,commitSupportResolution,inviteSupportHolder,legacySupportReports,convertLegacyReport,attachSupportIdentity} from '../_lib/support-requests.js';
import {identityTermHistory} from '../_lib/identity-term-history.js';
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
  if(action==='summary')return reply({ok:true,summary:await requestSummary(db,session.email),admin:requestAdmin(session),transfersEnabled:settings.requestTransfersEnabled===true});
  if(action==='list')return reply(await listSupportRequests(db,session,input));
  if(action==='detail')return reply(await readSupportRequest(db,session,input));
  if(action==='read')return reply({...await markSupportRead(db,session,input),summary:await requestSummary(db,session.email)});
  if(action==='evidence')return reply(await supportIdentityEvidence(db,session,input));
  if(action==='terms'){
   const evidence=await supportIdentityEvidence(db,session,input);return reply({ok:true,...(evidence.identity?await identityTermHistory(env.ROSTER_FILES,evidence.identity.aliases,melbourneDateKey(),input.after||0):{terms:[],next:null})});
  }
  if(action==='previewResolution')return reply(await previewSupportResolution(db,session,input));
  if(action==='legacy')return reply(await legacySupportReports(db,session,input));
  const mutations={attachIdentity:attachSupportIdentity,create:createSupportRequest,reply:replySupportRequest,status:setSupportStatus,inviteHolder:inviteSupportHolder,convertLegacy:convertLegacyReport,commitResolution:commitSupportResolution};
  if(!mutations[action])throw new AuthenticationError('requests-input','Unknown request action.',400);
  if(action==='commitResolution'){
   if(settings.requestTransfersEnabled!==true)throw new AuthenticationError('requests-review-gate','Identity decisions await the final security review.',403);
   if(!requestAdmin(session))throw new AuthenticationError('requests-forbidden','Admin access is required.',403);
   requireRecentAuthentication(session);
  }
  await admit(env,session.email,'support');
  if(env.ROSTER_ACCOUNT_BUDGET_ENABLED==='true'&&!await reserveRosterMaintenanceBudget(db,action==='commitResolution'?2048:128,action==='commitResolution'?4096:256))return reply({code:'requests-budget',error:'Requests are temporarily paused by the daily database safeguard. Your calendar remains available.'},503);
  const result=await mutations[action](db,session,input);
  if(result.operationId&&context.waitUntil)context.waitUntil(requestIdentityPublication(context,result.operationId).then(r=>{if(!r.ok)console.warn('Request identity publication deferred');}).catch(()=>console.warn('Request identity publication deferred')));
  return reply({...result,summary:await requestSummary(db,session.email)});
 }catch(error){
  if(['d1-statement-budget-exceeded','d1-request-deadline-exceeded'].includes(error?.code))throw error;
  if(error instanceof AuthenticationError)return reply({code:error.code,error:error.message},error.status);
  if(error?.code?.startsWith('IDENTITY_'))return reply({code:error.code,error:error.message},409);
  return reply({code:'requests-unavailable',error:'Could not load or update this request. Your calendar is still available.'},503);
 }
}
