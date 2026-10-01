export async function executeBoundedRosterImport(plan, callback) {
  const request = (phase, extra = {}) => callback({ phase, revision: plan.revision, ...extra });
  const progress = await request("bounded-begin", { manifest: plan.manifest });
  if (progress.mode === "complete" || progress.deferred || progress.completed) return progress;
  for (const batch of plan.batches.slice(Number(progress.nextBatch || 0))) {
    const result = await request("bounded-events", { batch });
    if (result.deferred || result.completed) return result;
  }
  for (const batch of plan.batches.slice(Number(progress.preparedBatch || 0))) {
    const result = await request("bounded-presence", { batch });
    if (result.deferred || result.completed) return result;
  }
  const metadata = await request("bounded-metadata");
  if (metadata.deferred || metadata.completed) return metadata;
  return request("bounded-activate");
}
