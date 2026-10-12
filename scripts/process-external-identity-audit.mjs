import {loadIdentityAuditSnapshot,planIdentityAudit} from './identity-audit-plan.mjs';
export async function continueExternalIdentityAudit({request,dispatchId='',maxSteps=600,maxMs=15*60*1000,now=Date.now}){
 const started=now();let count=0,identity=null;
 const call=async input=>{
  if(++count>maxSteps||now()-started>=maxMs)throw Object.assign(Error('Audit checkpoint/time limit'),{boundedPause:true});
  const result=await request({...identity,...input});
  if(['deferred','busy','import-active','registration-active','lease-expired'].includes(result.status)||result.deferred)throw Object.assign(Error('Audit deferred'),{auditDeferred:true});
  return result;
 };
 try{
  const start=await call({mode:'start',dispatchId});
  if(['external-disabled','unchanged','already-complete'].includes(start.status))return {status:start.status};
  identity={runId:start.runId,leaseToken:start.leaseToken};
  const snapshot=await loadIdentityAuditSnapshot(input=>call({mode:'snapshot',...input}));
  const plan=planIdentityAudit(snapshot,{fullAudit:true});
  const progress=await call({mode:'pin',manifest:plan.manifest,revision:plan.revision});
  if(!progress.completed){
   for(const batch of plan.batches.slice(progress.nextCandidateBatch||0))await call({mode:'apply',kind:'candidates',revision:plan.revision,batch});
   for(const batch of plan.featureBatches.slice(progress.nextFeatureBatch||0))await call({mode:'apply',kind:'features',revision:plan.revision,batch});
  }
  const result=await call({mode:'finish',revision:plan.revision});
  return {status:result.completed?'complete':'deferred',examined:plan.manifest.nameCount,proposals:plan.manifest.candidateCount,checkpoints:count};
 }catch(error){
  // Only confirmed bounded pauses or stale evidence may release/reset a lease.
  // Unknown HTTP/CPU outcomes retain it and their receipts for later resumption.
  if(identity&&(error.boundedPause||error.staleEvidence))await request({...identity,mode:error.staleEvidence?'abandon':'yield'}).catch(()=>{});
  if(error.boundedPause||error.auditDeferred||error.staleEvidence)return {status:'deferred',checkpoints:count};
  throw error;
 }
}
