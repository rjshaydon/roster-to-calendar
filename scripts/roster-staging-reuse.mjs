import {planRosterPatch} from './roster-patch-plan.mjs';

// Comparison lives on the runner; Pages only returns small indexed pages and
// applies hash-pinned, typed 16-entry batches to an invisible candidate.
export async function executeRosterRebase(payload,target,range,callback){
 const baseline=await callback({phase:'bounded-rebase-start',targetFileId:target.fileId,range});
 if(baseline.deferred||baseline.completed)return baseline;
 if(baseline.prepared)return promote(baseline.revision,callback);
 if(baseline.revision&&baseline.nextBatch>0)return finish(baseline.revision,baseline.nextBatch,baseline.batchCount,callback);
 const stored={};
 for(const [section,limit] of [['doctors',512],['events',25000],['issues',5000],['staff',750],['catalog',750],...(baseline.presenceMode==='cached'?[]:[['presence',25000]])]){
  const rows=[];let after='';
  do{
   const page=await callback({phase:'bounded-rebase-page',section,after});
   if(page.deferred||page.completed)return page;
   rows.push(...page.rows);after=page.next;
   if(rows.length>limit)throw Error('Staged baseline exceeds its inspection bound.');
  }while(after);
  stored[section]=section==='presence'?rows.flatMap(r=>JSON.parse(r.rows_json)):rows;
 }
 const plan=await planRosterPatch(payload,baseline,stored,{staging:true,presenceMode:baseline.presenceMode});
 console.log(`Staging reuse: ${stored.events.length} retained events; ${plan.manifest.changedFacts} changed facts; ${plan.batches.length} small completion batches.`);
 const pinned=await callback({phase:'bounded-rebase-plan',manifest:plan.manifest,patchRevision:plan.patchRevision});
 if(pinned.deferred||pinned.completed)return pinned;
 // Persist the immutable plan outside D1 before the first fact mutation. A
 // retry must not diff an already partly changed baseline into a new plan.
 for(const batch of plan.batches){
  const saved=await callback({phase:'bounded-rebase-save',batch,patchRevision:plan.patchRevision});
  if(saved.deferred||saved.completed)return saved;
 }
 return finish(plan.patchRevision,pinned.nextBatch||0,plan.batches.length,callback);
}
async function finish(revision,next,count,callback){
 for(let batchIndex=next;batchIndex<count;batchIndex++){
  const result=await callback({phase:'bounded-rebase-stage',batchIndex,patchRevision:revision});
  if(result.deferred||result.completed)return result;
 }
 const prepared=await callback({phase:'bounded-rebase-prepare',patchRevision:revision});
 if(prepared.deferred||prepared.completed)return prepared;
 return promote(prepared.revision,callback);
}
async function promote(revision,callback){
 const active=await callback({phase:'bounded-activate',revision,deferCompletion:true});
 if(active.deferred||active.completed)return active;
 return callback({phase:'bounded-finalize',revision});
}
