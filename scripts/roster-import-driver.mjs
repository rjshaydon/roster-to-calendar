import {rosterImportDigest} from '../functions/_lib/roster-import-batches.js';
export async function executeBoundedRosterImport(plan, callback) {
  const request = (phase, extra = {}) => callback({ phase, revision: plan.revision, ...extra });
  let progress = await request("bounded-begin", { manifest: plan.manifest, deferStaging: true });
  if (progress.mode === "staging-ready") {
    if (progress.presenceMode === 'cached') {
      plan = {...plan,manifest:{...plan.manifest,presenceMode:'cached'}};
      plan.revision = await rosterImportDigest(plan.manifest);
    }
    progress = await request("bounded-initialize", {manifest:plan.manifest});
  }
  if (progress.mode === "complete" || progress.mode === "patch" || progress.mode === 'rebase' || progress.deferred || progress.completed) return progress;
  for (const batch of plan.batches.slice(Number(progress.nextBatch || 0))) {
    const result = await request("bounded-events", { batch });
    if (result.deferred || result.completed) return result;
  }
  for (const batch of plan.batches.slice(Number(progress.preparedBatch || 0))) {
    const result = await request("bounded-presence", { batch });
    if (result.deferred || result.completed) return result;
    if (result.nextBatch === plan.batches.length) break;
  }
  for (const batch of plan.compactBatches || []) {
    const result = await request("bounded-metadata", { batch });
    if (result.deferred || result.completed) return result;
  }
  const metadata = await request("bounded-metadata");
  if (metadata.deferred || metadata.completed) return metadata;
  const activated = await request("bounded-activate",{deferCompletion:true});
  if (activated.deferred || activated.completed) return activated;
  return request("bounded-finalize");
}
