import {identityMaintenancePreflight} from '../../_lib/identity-maintenance-preflight.js';
import {readIdentityAuditSnapshotPage} from '../../_lib/identity-audit-snapshot.js';
import {pinExternalIdentityAudit,applyExternalIdentityAuditBatch,finishExternalIdentityAudit} from '../../_lib/identity-audit-checkpoints.js';
import {reserveRosterMaintenanceBudget} from '../../_lib/roster-maintenance-budget.js';
import {identityAuditWindow} from '../../../public/static/identity-audit-policy.js';
import {guardedFetch,localFeatureDisabledResponse} from '../../_lib/outbound-network.js';
const KEY='identity/external-audit-progress.json',ALGORITHM='external-pairs-v1',LEASE_MS=20*60*1000;
export async function onRequestPost(context){
 const input=context.data?.identityAuditInput??await context.request.json().catch(()=>({}));
 const auth=context.request.headers.get('authorization');
 const runner=Boolean(context.env.ROSTER_AUTOMATION_TOKEN)&&auth==='Bearer '+context.env.ROSTER_AUTOMATION_TOKEN;
 const watchdog=input.mode==='check'&&Boolean(context.env.ROSTER_WATCHDOG_TOKEN)&&auth==='Bearer '+context.env.ROSTER_WATCHDOG_TOKEN;
 if(!runner&&!watchdog)return Response.json({error:'Unauthorized.'},{status:401});
 return handleExternalIdentityAudit(context,input);
}
export async function handleExternalIdentityAudit(context,input){
 const env=context.env,db=env.ROSTER_DB,r2=env.ROSTER_FILES;
 if(env.IDENTITY_EXTERNAL_AUDIT_ENABLED!=='true'||env.IDENTITY_REVIEW_ENABLED!=='true')return Response.json({status:'external-disabled'});
 if(!['check','start','snapshot','pin','apply','finish','abandon','yield'].includes(input.mode))return Response.json({error:'Invalid audit operation.'},{status:400});
 const local=localFeatureDisabledResponse(env,'External identity audit');if(local)return local;
 try{
  const object=await r2.get(KEY),state=object?await object.json():{},now=new Date(),instant=now.toISOString();
  const active=state.leaseUntil>instant;
  if(['check','start'].includes(input.mode)){
   if(active&&!(input.mode==='start'&&input.dispatchId===state.leaseToken))return Response.json({status:'busy'});
   const registry=await r2.get('identity/registry-progress.json');if(registry&&(await registry.json()).runId)return Response.json({status:'registration-active'});
   const meta=await db.prepare(`SELECT revision,EXISTS(SELECT 1 FROM roster_identity_features INDEXED BY idx_identity_feature_pending WHERE audited_fingerprint<>fingerprint LIMIT 1) AS pending,
     (SELECT stop_reason FROM roster_account_budget WHERE utc_day=?) AS stop_reason FROM roster_identity_revision WHERE id=1`).bind(instant.slice(0,10)).first();
   if(!meta||!Number.isSafeInteger(Number(meta.revision)))throw Error('Identity evidence unavailable.');
   if(meta.stop_reason)return Response.json({status:'deferred'});
   const window=identityAuditWindow(now);
   if(!state.runId&&state.algorithm===ALGORITHM&&!meta.pending&&(!window.catchupEligible||state.weekKey===window.weekKey))return Response.json({status:'unchanged'});
   const preflight=await identityMaintenancePreflight(db,{accountAware:env.ROSTER_ACCOUNT_BUDGET_ENABLED==='true'});
   if(preflight.active_import)return Response.json({status:'import-active'});
   if(input.mode==='start'&&(Number(preflight.unfinished_count)>=3||!Number(preflight.current_grant)))return Response.json({status:'deferred'});
   if(input.mode==='check'&&!env.GITHUB_ACTIONS_TOKEN)return Response.json({status:'dispatch-unconfigured'});
   const resume=state.runId&&state.identityRevision===Number(meta.revision);
   const next={...state,runId:resume?state.runId:crypto.randomUUID(),identityRevision:Number(meta.revision),weekKey:resume?state.weekKey:window.weekKey,
    leaseToken:input.mode==='start'&&active?state.leaseToken:crypto.randomUUID(),leaseUntil:new Date(now.getTime()+LEASE_MS).toISOString(),status:input.mode==='check'?'dispatched':'running'};
   if(!await save(r2,object,next))return Response.json({status:'busy'});
   if(input.mode==='start')return Response.json({status:'running',runId:next.runId,leaseToken:next.leaseToken,identityRevision:next.identityRevision,fullAudit:true});
   const response=await guardedFetch(env,'https://api.github.com/repos/rjshaydon/roster-to-calendar/actions/workflows/identity-maintenance.yml/dispatches',{
    method:'POST',headers:{Accept:'application/vnd.github+json',Authorization:'Bearer '+env.GITHUB_ACTIONS_TOKEN,'Content-Type':'application/json','X-GitHub-Api-Version':'2022-11-28','User-Agent':'roster-to-calendar'},
    body:JSON.stringify({ref:env.ROSTER_GITHUB_WORKFLOW_REF||'main',inputs:{mode:'audit',audit_dispatch_id:next.leaseToken}}),signal:AbortSignal.timeout(15000)}, {label:'Identity workflow dispatch'});
   // A lost/rejected dispatch keeps its lease: never duplicate an uncertain job.
   return Response.json({status:response.status===204?'queued':'dispatch-deferred'});
  }
  if(!active||input.runId!==state.runId||input.leaseToken!==state.leaseToken)return Response.json({status:'lease-expired'},{status:409});
  const preflight=await identityMaintenancePreflight(db,{accountAware:env.ROSTER_ACCOUNT_BUDGET_ENABLED==='true',identityRevision:true});
  if(Number(preflight.identity_revision)!==state.identityRevision&& !['abandon','yield'].includes(input.mode))return Response.json({status:'stale'},{status:409});
  if(preflight.active_import)return Response.json({status:'import-active'});
  if(Number(preflight.unfinished_count)>=3||!Number(preflight.current_grant))return Response.json({status:'deferred'},{status:503});
  const common={runId:state.runId,revision:input.revision};let result;
  if(input.mode==='snapshot'){
   if(input.identityRevision!=null&&input.identityRevision!==state.identityRevision)return Response.json({status:'stale'},{status:409});
   if(!await reserveRosterMaintenanceBudget(db,0,4096))return Response.json({status:'deferred'},{status:503});
   result=await readIdentityAuditSnapshotPage(db,{...input,identityRevision:state.identityRevision});
  }else if(input.mode==='pin'){
   if(input.manifest?.identityRevision!==state.identityRevision||input.manifest?.fullAudit!==true)throw Object.assign(Error('External audit requires its complete pinned snapshot.'),{code:'IDENTITY_AUDIT_EVIDENCE'});
   result=await pinExternalIdentityAudit(db,{...common,manifest:input.manifest});
  }else if(input.mode==='apply')result=await applyExternalIdentityAuditBatch(db,{...common,kind:input.kind,batch:input.batch});
  else if(input.mode==='finish'){
   result=await finishExternalIdentityAudit(db,common);
   if(result.completed){
    await r2.put('identity/revision.json',JSON.stringify({revision:crypto.randomUUID()}),{httpMetadata:{contentType:'application/json'}});
    if(!await save(r2,object,{...state,algorithm:ALGORITHM,runId:'',leaseUntil:'',leaseToken:'',status:'complete'}))return Response.json({status:'busy'});
   }
  }else{
   result={status:'deferred'};
   if(!await save(r2,object,{...state,runId:input.mode==='abandon'?'':state.runId,leaseUntil:'',leaseToken:'',status:'pending'}))return Response.json({status:'busy'});
  }
  return Response.json(result);
 }catch(error){
  // Public workflow logs must not contain roster names, SQL values or bodies.
  console.warn(JSON.stringify({event:'external-identity-audit',mode:input.mode,status:'stopped'}));
  return error.code==='IDENTITY_AUDIT_EVIDENCE'
   ? Response.json({status:'stale',error:'Audit evidence changed or its checkpoint was rejected.'},{status:409})
   : Response.json({status:'interrupted',error:'Audit interrupted; durable progress is retained.'},{status:500});
 }
}
async function save(r2,previous,value){
 const result=await r2.put(KEY,JSON.stringify(value),{httpMetadata:{contentType:'application/json'},...(previous?.etag?{onlyIf:{etagMatches:previous.etag}}:{onlyIf:{etagDoesNotMatch:'*'}})});
 return result!==null;
}
