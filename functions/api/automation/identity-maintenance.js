import {publishIdentityOperation} from '../../_lib/doctor-identity.js';
import {refreshAccountMaintenanceBudget} from './account-budget.js';
import {reserveRosterMaintenanceBudget} from '../../_lib/roster-maintenance-budget.js';
import {auditIdentityBatch} from '../../_lib/identity-discovery.js';
import {publishedIdentityDirectory} from '../../_lib/bounded-identity.js';
import {identityAuditWindow} from '../../../public/static/identity-audit-policy.js';
import {facilityMetadataManifestKey} from '../../_lib/facility-overview-cache.js';
import {melbourneDateKey} from '../../../public/static/roster-term-policy.js';
export async function onRequestPost(context) {
 if(context.env.IDENTITY_REVIEW_ENABLED!=='true') return new Response('Unavailable',{status:503});
 const input=await context.request.json().catch(()=>({}));
 const tokens=[context.env.IDENTITY_MAINTENANCE_TOKEN,context.env.ROSTER_AUTOMATION_TOKEN,...(input.mode==='publish'?[]:[context.env.ROSTER_WATCHDOG_TOKEN])].filter(Boolean).map(String);
 if(!tokens.some(token=>context.request.headers.get('Authorization')===`Bearer ${token}`)) return Response.json({error:'Unauthorized.'},{status:401});
 if(!['publish','audit','register'].includes(input.mode) || input.mode==='publish' && !/^[a-f0-9-]{36}$/.test(String(input.operationId||''))) return new Response('Invalid operation',{status:400});
 const db=context.env.ROSTER_DB;
 const window=identityAuditWindow();
 let registryState=null,registryObject=null,registryRevision='';
 if(input.mode==='register') {
  if(context.env.IDENTITY_REGISTRY_ENABLED!=='true') return Response.json({status:'registry-paused'});
  const r2=context.env.ROSTER_FILES;
  const heads=await Promise.all(['mmc','mch','ddh','vhh'].map(source=>r2.head(facilityMetadataManifestKey(source))));
  if(heads.some(head=>!head)) return Response.json({status:'directory-preparing'});
  registryRevision=JSON.stringify([melbourneDateKey().slice(0,7),...heads.map(head=>head.etag)]);
  registryObject=await r2.get('identity/registry-progress.json');
  registryState=registryObject?await registryObject.json():{};
  if(!registryState.runId && registryState.revision===registryRevision) return Response.json({status:'unchanged'});
 }

 if(input.mode==='audit') {
  if(context.env.IDENTITY_SCHEDULED_AUDIT_ENABLED!=='true' || !window.eligible) return Response.json({status:'outside-audit-window'});
  const done=await db.prepare("SELECT run_id FROM roster_identity_audit_runs WHERE week_key=? AND status='complete' AND mode='audit' LIMIT 1").bind(window.weekKey).first();
  if(done) return Response.json({status:'already-complete'});
 }
 if(['audit','register'].includes(input.mode)) {
  for(const source of ['monash-adults','monash-paeds','vhh-active-medical-roster','dandenong-findmyshift']) {
    if(await db.prepare("SELECT id FROM roster_sync_runs WHERE source_id=? AND status IN ('queued','processing') LIMIT 1").bind(source).first()) return Response.json({status:'import-active'});
  }
 }
 if(context.env.ROSTER_ACCOUNT_BUDGET_ENABLED==='true') {
  const admission=await (await refreshAccountMaintenanceBudget(context)).json();
  if(admission.deferred || !await reserveRosterMaintenanceBudget(db,4096,input.mode==='publish'?4096:8192)) return Response.json({status:'deferred'},{status:503});
 }
 if(['audit','register'].includes(input.mode)) {
  const directory=await publishedIdentityDirectory(context.env.ROSTER_FILES,melbourneDateKey());
  if(directory.preparing || directory.missingSources?.length) return Response.json({status:'directory-preparing'},{status:503});
  const result=await auditIdentityBatch(db,directory.doctors,{runId:input.mode==='register'?registryState.runId:undefined,register:input.mode==='register',actor:'published-roster-registration',weekKey:input.mode==='audit'?window.weekKey:''});
  if(input.mode==='register' && result.status!=='busy') {
   const progress={revision:registryState.runId?registryState.revision:registryRevision,runId:result.status==='complete'?'':result.runId};
   await context.env.ROSTER_FILES.put('identity/registry-progress.json',JSON.stringify(progress),{httpMetadata:{contentType:'application/json'},...(registryObject?.etag?{onlyIf:{etagMatches:registryObject.etag}}:{onlyIf:{etagDoesNotMatch:'*'}})});
   if(result.status==='complete') await context.env.ROSTER_FILES.put('identity/revision.json',JSON.stringify({revision:crypto.randomUUID()}),{httpMetadata:{contentType:'application/json'}});
  }
  return Response.json(result);
 }
 const result=await publishIdentityOperation(context.env.ROSTER_DB,context.env.ROSTER_FILES,input.operationId);
 return Response.json(result,{status:result.status==='missing'?404:result.status==='failed'?409:200});
}
