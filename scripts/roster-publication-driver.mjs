// Shared by initial delivery and publication-only recovery. The same bounded
// counter survives before/after-import calls in one runner invocation.
export async function drainRosterPublication({step,publishMonth,publishSource,budget={remaining:24}}){
 let reconciliations=0;
 while(budget.remaining>0){
  budget.remaining--;
  const refresh=await step();
  if(refresh.recoverableConflict){
   if(++reconciliations>1)return {deferred:true,reason:'publication-conflict'};
   continue;
  }
  if(refresh.doctorCalendarsEnabled&&refresh.mode==='build-month'&&refresh.revision)await publishMonth(refresh);
  if(refresh.doctorCalendars?.length){for(const pointer of refresh.doctorCalendars)await publishMonth({...refresh,...pointer});continue;}
  if(refresh.doctorSource){await publishSource(refresh);continue;}
  if(refresh.deferred)return {deferred:true,reason:refresh.reason||'account-budget'};
  if(refresh.idle)return {completed:true};
  // Completing one term does not authorize another import while another term
  // is pending. Poll again within the same bounded pass.
 }
 return {deferred:true,reason:'run-limit'};
}
