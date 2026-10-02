import { rosterDeliveryOrder, skipSupersededRosterDelivery } from './roster-delivery-order.js';
import { beginBoundedRosterImport, stageBoundedRosterBatch, prepareBoundedRosterPresence, prepareBoundedRosterMetadata, activateBoundedRosterTerm } from "./roster-import-staging.js";
import { loadRosterSyncRun, loadRosterSource, loadRosterFileStatusSummary, finishRosterSyncRun, upsertRosterSource, supersedeDuplicateRosterSyncRuns } from "./d1-calendar.js";
import { reviewedRosterFactLimit } from "./roster-automation-guard.js";
import { automaticFacilityPublicationEnabled } from "./facility-refresh-queue.js";
import { rosterImportDigest, validateRosterImportManifest } from "./roster-import-batches.js";

export async function handleBoundedRosterRequest(context, body, source, resolveTarget) {
  if (String(context.env.ROSTER_AUTOMATION_BOUNDED_IMPORT_ENABLED || "") !== "true") return Response.json({ error: "Bounded roster importing is not enabled." }, { status: 503 });
  const db = context.env.ROSTER_DB;
  try {
    const run = await loadRosterSyncRun(db, String(body.runId || ""));
    if (!run || run.sourceId !== body.sourceId || run.fileId !== body.file?.id) return Response.json({ error: "Bounded import does not match its queued source/file." }, { status: 400 });
    if (reviewedRosterFactLimit(context.env, run.sourceId, run.contentHash) !== 1250) throw new Error("Bounded import requires the reviewed source/content fact budget.");
    if (run.status === "success") return Response.json({ ok: true, completed: true, fileId: run.fileId, doctorCount: run.doctorCount, eventCount: run.eventCount });
    if (run.status === "superseded") return Response.json({ok:true,completed:true,superseded:true,doctorCount:0,eventCount:0});
    const delivery = await rosterDeliveryOrder(db, run);
    if (delivery.superseded) return Response.json(await skipSupersededRosterDelivery(db,run));
    const file = { ...body.file, sourceId: run.sourceId, sourceType: source.sourceType, active: false };
    let result;
    switch (body.phase) {
      case "bounded-begin": {
        const manifest = body.manifest;
        if (!manifest || manifest.contentHash !== run.contentHash || manifest.maximumFacts !== 1250) throw new Error("Bounded plan does not match the queued content or reviewed fact budget.");
        validateRosterImportManifest(manifest);
        if (manifest.fileId !== file.id || manifest.sourceId !== file.sourceId || await rosterImportDigest(manifest) !== body.revision) throw new Error("Bounded plan does not match its queued file or revision.");
        const replacements = context.env.ROSTER_BOUNDED_REPLACEMENT_ENABLED === "true";
        const staged = await db.prepare("SELECT next_batch,prepared_batch FROM roster_import_jobs WHERE run_id=?").bind(run.id).first();
        if ((!replacements || body.forceStaging !== true) && !Number(staged?.next_batch || 0) && !Number(staged?.prepared_batch || 0)) {
          const current = await loadRosterSource(db, run.sourceId);
          const target = await resolveTarget(db, source, run.sourceId, run, { range: [{ start: manifest.startDate }, { start: manifest.rosterEndDate }] }, current);
          if (target.fileId !== run.fileId) return Response.json({ ok: true, mode: "complete", boundedReplacementAvailable: replacements });
        }
        result = await beginBoundedRosterImport(db, run.id, file, { manifest, revision: body.revision }, { allowReplacement: replacements, allowEmptyPlanRecovery: true });
        // An activated job still needs the completion callback if its earlier
        // bookkeeping failed. Only run.status=success means fully completed.
        result = { ...result, activated: Boolean(result.completed), completed: false };
        break;
      }
      case "bounded-events": result = await stageBoundedRosterBatch(db, run.id, file, body.revision, body.batch); break;
      case "bounded-presence": result = await prepareBoundedRosterPresence(db, run.id, file, body.revision, body.batch); break;
      case "bounded-metadata": result = await prepareBoundedRosterMetadata(db, run.id, body.revision); break;
      case "bounded-activate": {
        result = await activateBoundedRosterTerm(db, run.id, body.revision, { publishFacility: automaticFacilityPublicationEnabled(context.env, source.sourceType), deliveryGuard: delivery.guard });
        if (result.deferred) break;
        const summary = await loadRosterFileStatusSummary(db, run.fileId);
        const current = await loadRosterSource(db, run.sourceId);
        const latest = current?.activeFileId ? await loadRosterFileStatusSummary(db, current.activeFileId) : null;
        const now = new Date().toISOString();
        await upsertRosterSource(db, { ...current, ...source, id: run.sourceId, enabled: true, lastSuccessAt: now, lastError: "", activeFileId: latest?.startDate > summary.startDate ? current.activeFileId : run.fileId, updatedAt: now, createdAt: current?.createdAt || now });
        await supersedeDuplicateRosterSyncRuns(db, run, file.name || "");
        await finishRosterSyncRun(db, run.id, { status: "success", fileId: run.fileId, doctorCount: summary.indexedDoctors, eventCount: summary.eventCount, completedAt: now, message: "Bounded next-term import completed." });
        result = { ...result, completed: true, doctorCount: summary.indexedDoctors, eventCount: summary.eventCount };
        break;
      }
      default: return Response.json({ error: "Unknown bounded import phase." }, { status: 400 });
    }
    return Response.json({ ok: true, mode: "bounded", ...result });
  } catch (error) {
    if (error?.code === "ROSTER_MAINTENANCE_DEFERRED") return Response.json({ ok: true, deferred: true });
    // Retain the inactive file and receipts for retry. Do not enter the legacy
    // failed callback's destructive cleanup path after an interrupted batch.
    console.error("Bounded roster import stopped", { runId: body.runId, message: error.message });
    return Response.json({ error: error.message, resumable: true }, { status: 422 });
  }
}
