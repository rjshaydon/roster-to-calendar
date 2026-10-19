import {scheduleStaffHistoryBuild} from '../../_lib/staff-history-cache.js';
import {identityMaintenancePreflight} from '../../_lib/identity-maintenance-preflight.js';
import {publishIdentityOperation} from '../../_lib/doctor-identity.js';
import {refreshAccountMaintenanceBudget} from './account-budget.js';
import {reserveRosterMaintenanceBudget,optionalMaintenanceAvailable,currentAccountMaintenanceGrant} from '../../_lib/roster-maintenance-budget.js';
import {auditIdentityBatch,cachedIdentityRegistrationMissing} from '../../_lib/identity-discovery.js';
import {identityAuditWindow} from '../../../public/static/identity-audit-policy.js';
import {facilityMetadataManifestKey} from '../../_lib/facility-overview-cache.js';
import {melbourneDateKey} from '../../../public/static/roster-term-policy.js';
export async function onRequestPost(context) {
 if(context.env.IDENTITY_REVIEW_ENABLED!=='true') return new Response('Unavailable',{status:503});
 const input=context.data?.identityMaintenanceInput ?? await context.request.json().catch(()=>({}));
 const tokens=[context.env.IDENTITY_MAINTENANCE_TOKEN,context.env.ROSTER_AUTOMATION_TOKEN,...(input.mode==='publish'?[]:[context.env.ROSTER_WATCHDOG_TOKEN])].filter(Boolean).map(String);
 if(!tokens.some(token=>context.request.headers.get('Authorization')===`Bearer ${token}`)) return Response.json({error:'Unauthorized.'},{status:401});
 if(!['publish','audit','register'].includes(input.mode) || input.mode==='publish' && !/^[a-f0-9-]{36}$/.test(String(input.operationId||''))) return new Response('Invalid operation',{status:400});
 if(input.mode==='audit'&&context.env.IDENTITY_EXTERNAL_AUDIT_ENABLED==='true'){
  const {handleExternalIdentityAudit}=await import('./identity-audit.js');
  return handleExternalIdentityAudit(context,{mode:'check'});
 }
 const db=context.env.ROSTER_DB;
 let registryState=null,registryObject=null,registryRevision='';
 let auditState=null,auditObject=null,auditWeek='';
 if(input.mode==='register') {
  if(context.env.IDENTITY_REGISTRY_ENABLED!=='true') return Response.json({status:'registry-paused'});
  const r2=context.env.ROSTER_FILES;
  registryObject=await r2.get('identity/registry-progress.json');
  registryState=registryObject?await registryObject.json():{};
  // Active jobs defer before optional timezone/manifest preparation. Preserve
  // the zero-D1 shortcut for unchanged, completed registry metadata.
  if(registryState.runId&&(await identityMaintenancePreflight(db)).active_import)return Response.json({status:'import-active'});
  if(!registryState.runId || registryState.finalizing) {
   const heads=await Promise.all(['mmc','mch','ddh','vhh'].map(source=>r2.head(facilityMetadataManifestKey(source))));
   if(heads.some(head=>!head)) return Response.json({status:'directory-preparing'});
   registryRevision=JSON.stringify(['cached-names-v2',melbourneDateKey().slice(0,7),...heads.map(head=>head.etag)]);
  } else registryRevision=registryState.revision;
  if(!registryState.runId && registryState.revision===registryRevision) return Response.json({status:'unchanged'});
 }

 if(input.mode==='audit') {
  if(context.env.IDENTITY_SCHEDULED_AUDIT_ENABLED!=='true') return Response.json({status:'outside-audit-window'});
  const window=identityAuditWindow();
  auditObject=await context.env.ROSTER_FILES.get('identity/weekly-audit-progress.json');
  auditState=auditObject?await auditObject.json():{};
  if(!auditState.runId && !window.catchupEligible) return Response.json({status:'outside-audit-window'});
  auditWeek=auditState.runId?auditState.weekKey:window.weekKey;
  if(!auditState.runId && auditState.weekKey===auditWeek && auditState.status==='complete') return Response.json({status:'already-complete'});
  const registration=await context.env.ROSTER_FILES.get('identity/registry-progress.json');
  if(registration && (await registration.json()).runId) return Response.json({status:'registration-active'});
 }
 let preflight;
 if(['audit','register'].includes(input.mode)) {
  const runId=input.mode==='register'?registryState.runId:auditState.runId;
  preflight=await identityMaintenancePreflight(db,{runId,weekKey:input.mode==='audit'?auditWeek:'',accountAware:context.env.ROSTER_ACCOUNT_BUDGET_ENABLED==='true'});
  if(preflight.completed_run) {
   await saveAuditProgress(context.env.ROSTER_FILES,auditObject,{weekKey:auditWeek,runId:'',status:'complete'});
   return Response.json({status:'already-complete'});
  }
  if(preflight.active_import) return Response.json({status:'import-active'});
  if(preflight.live_lease) return Response.json({status:'busy'});
 }
 if(context.env.ROSTER_ACCOUNT_BUDGET_ENABLED==='true') {
  const available=preflight?Number(preflight.unfinished_count)<3:await optionalMaintenanceAvailable(db);
  const granted=preflight?Number(preflight.current_grant)===1:available&&await currentAccountMaintenanceGrant(db);
  if(!available || !granted) {
   // Refresh also reconciles safely expired receipts before applying the
   // unfinished-work circuit breaker. No reservation is granted here.
   const admission=await (await refreshAccountMaintenanceBudget(context)).json();
   if(admission.deferred) return Response.json({status:'deferred'},{status:503});
  }
  if(!available && !await optionalMaintenanceAvailable(db)) return Response.json({status:'deferred',reason:'unfinished-maintenance'},{status:503});
  if(!await reserveRosterMaintenanceBudget(db,input.mode==='publish'?4096:512,131072)) return Response.json({status:'deferred'},{status:503});
 }
 if(['audit','register'].includes(input.mode)) {
  if(input.mode==='register' && registryState.finalizing) {
   // Capture current metadata before rechecking name coverage. A late source
   // publication changes that watermark and gets checked on the next tick.
   const missing=await cachedIdentityRegistrationMissing(db);
   await saveRegistryProgress(context.env.ROSTER_FILES,registryObject,missing
    ?{revision:registryRevision,runId:registryState.runId,restart:true}
    :{revision:registryRevision,runId:''});
   if(!missing) {
    await context.env.ROSTER_FILES.put('identity/revision.json',JSON.stringify({revision:crypto.randomUUID()}),{httpMetadata:{contentType:'application/json'}});
    scheduleStaffHistoryBuild(context);
   }
   return Response.json({status:missing?'running':'complete',runId:registryState.runId,examined:registryState.examined});
  }
  // The indexed names cache retains historical staff. Register one name per
  // call and audit one pending feature, without decompressing every site's
  // published staff directory inside a free Pages request.
  const result=await auditIdentityBatch(db,[],{cachedDirectory:true,runId:input.mode==='register'?(registryState.restart?'':registryState.runId):auditState.runId,register:input.mode==='register',actor:'published-roster-registration',weekKey:input.mode==='audit'?auditWeek:''});
  if(input.mode==='audit' && result.status!=='busy') await saveAuditProgress(context.env.ROSTER_FILES,auditObject,
   {weekKey:auditWeek,runId:result.status==='complete'?'':result.runId,status:result.status});
  if(input.mode==='register' && result.status!=='busy') {
   const progress={revision:registryState.runId?registryState.revision:registryRevision,runId:result.runId,
    ...(result.status==='complete'?{finalizing:true,examined:result.examined}:{})};
   await saveRegistryProgress(context.env.ROSTER_FILES,registryObject,progress);
   if(result.status==='complete') return Response.json({...result,status:'running'});
  }
  return Response.json(result);
 }
 const result=await publishIdentityOperation(context.env.ROSTER_DB,context.env.ROSTER_FILES,input.operationId);
 if(result.status==='complete')scheduleStaffHistoryBuild(context);
 return Response.json(result,{status:result.status==='missing'?404:result.status==='failed'?409:200});
}
async function saveRegistryProgress(r2,previous,value) {
 await r2.put('identity/registry-progress.json',JSON.stringify(value),{httpMetadata:{contentType:'application/json'},
  ...(previous?.etag?{onlyIf:{etagMatches:previous.etag}}:{onlyIf:{etagDoesNotMatch:'*'}})});
}
async function saveAuditProgress(r2,previous,value) {
 await r2.put('identity/weekly-audit-progress.json',JSON.stringify(value),{httpMetadata:{contentType:'application/json'},
  ...(previous?.etag?{onlyIf:{etagMatches:previous.etag}}:{onlyIf:{etagDoesNotMatch:'*'}})});
}
