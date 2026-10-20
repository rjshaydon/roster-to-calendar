import { extractVhhHistoryWindow } from './vhh-history-extract.mjs';
import { planVhhHistoryWindows } from './vhh-history-plan.mjs';
import { buildVhhDerivedRosterPayload } from '../functions/_lib/vhh-roster.js';
import { rosterImportDigest,planRosterImportBatches } from '../functions/_lib/roster-import-batches.js';
import { historyFileName } from '../functions/_lib/vhh-history-preservation.js';
import { attachRosterCompactPlan } from './roster-compact-plan.mjs';
import { executeBoundedRosterImport } from './roster-import-driver.mjs';
export async function preserveVhhHistory({run,extract,request,recoveryOnly=false}) {
  const inspect=await request({phase:'bounded-history-inspect',runId:run.id});
  if(inspect.deferred) return inspect;
  let windows=planVhhHistoryWindows(extract,inspect.active);
  if(recoveryOnly) windows=windows.filter(w=>!w.overlap);
  const prepared=[];
  for(const window of windows) {
    const {from,to}=window;
    const history=extractVhhHistoryWindow(extract,{from,to,fileName:historyFileName(from,to)});
    const digest=await rosterImportDigest([run.id,from,to]);
    const payload=buildVhhDerivedRosterPayload({extract:history,fileId:`vhh-history:${digest}`,contentHash:digest,lastModified:run.lastModified});
    let plan=await attachRosterCompactPlan(await planRosterImportBatches(payload,{maximumBatchBytes:65536,maximumOccurrences:64}),payload);
    if(inspect.cachedPresence)plan.manifest.presenceMode='cached';
    plan.revision=await rosterImportDigest(plan.manifest);
    const result=await executeBoundedRosterImport(plan,step=>request({...step,file:payload.file,vhhHistory:{from,to},prepareOnly:Boolean(window.overlap)}));
    if(result.deferred) return result;
    if(window.overlap) prepared.push({from,to});
    console.log(`VHH history ${from}–${to}: ${payload.eventCount} shifts ${window.overlap?'prepared for atomic rollover':'recovered'}.`);
  }
  return {historyWindows:prepared};
}
