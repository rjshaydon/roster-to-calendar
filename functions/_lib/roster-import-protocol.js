import {activeAuthorityFile} from './roster-authority.js';
import { beginRosterPatch,readRosterPatchPage,pinRosterPatchPlan,stageRosterPatchBatch,activateRosterPatch } from './roster-patch.js';
import { rosterDeliveryOrder, skipSupersededRosterDelivery } from './roster-delivery-order.js';
import { beginBoundedRosterImport, stageBoundedRosterBatch, prepareBoundedRosterPresence, prepareBoundedRosterMetadata, activateBoundedRosterTerm } from "./roster-import-staging.js";
import { loadRosterSyncRun, loadRosterSource, loadRosterFileStatusSummary, finishRosterSyncRun, upsertRosterSource, supersedeDuplicateRosterSyncRuns } from "./d1-calendar.js";
import { reviewedRosterFactLimit } from "./roster-automation-guard.js";
import { automaticFacilityPublicationEnabled } from "./facility-refresh-queue.js";
import { rosterImportDigest, validateRosterImportManifest } from "./roster-import-batches.js";
import { reusableRosterStage,beginRosterRebase,readRosterRebasePage,pinRosterRebasePlan,saveRosterRebaseBatch,applyRosterRebaseBatch,prepareRosterRebase } from './roster-staging-reuse.js';

export async function handleBoundedRosterRequest(context, body, source, resolveTarget) {
  if (String(context.env.ROSTER_AUTOMATION_BOUNDED_IMPORT_ENABLED || "") !== "true") return Response.json({ error: "Bounded roster importing is not enabled." }, { status: 503 });
  const db = context.env.ROSTER_DB;
  try {
    const run = await loadRosterSyncRun(db, String(body.runId || ""));
    if (!run || run.sourceId !== body.sourceId || (run.fileId !== body.file?.id && !(run.status === "success" && run.sourceFileId === body.file?.id))) return Response.json({ error: "Bounded import does not match its queued source/file." }, { status: 400 });
    if (reviewedRosterFactLimit(context.env, run.sourceId, run.contentHash) !== 1250) throw new Error("Bounded import requires the reviewed source/content fact budget.");
    if (run.status === "success") return Response.json({ ok: true, completed: true, fileId: run.fileId, doctorCount: run.doctorCount, eventCount: run.eventCount });
    if (run.status === "superseded") return Response.json({ok:true,completed:true,superseded:true,doctorCount:0,eventCount:0});
    // A locator request cannot mutate calendars or staging. Defer ordering
    // evidence to initialization/patch start and recheck it at activation.
    const locating = body.phase === 'bounded-begin' && body.deferStaging === true;
    const delivery = locating ? {} : await rosterDeliveryOrder(db, run);
    if (delivery.superseded) return Response.json(await skipSupersededRosterDelivery(db,run));
    const file = { ...body.file, sourceId: run.sourceId, sourceType: source.sourceType, active: false };
    const cachedPresence = context.env.ROSTER_STAGING_REUSE_SOURCES?.split(',').includes(run.sourceId);
    if (body.manifest?.presenceMode && (body.manifest.presenceMode !== 'cached' || !cachedPresence)) throw Error('Cached presence is not enabled for this source.');
    let result;
    switch (body.phase) {
      case "bounded-begin":
      case "bounded-initialize": {
        const manifest = body.manifest;
        if (!manifest || manifest.contentHash !== run.contentHash || manifest.maximumFacts !== 1250) throw new Error("Bounded plan does not match the queued content or reviewed fact budget.");
        validateRosterImportManifest(manifest);
        if (manifest.fileId !== file.id || manifest.sourceId !== file.sourceId || await rosterImportDigest(manifest) !== body.revision) throw new Error("Bounded plan does not match its queued file or revision.");
        const replacements = context.env.ROSTER_BOUNDED_REPLACEMENT_ENABLED === "true";
        if (body.phase === 'bounded-begin' && (!replacements || body.forceStaging !== true)) {
          const target = await resolveTarget(db, source, run.sourceId, run, { range: [{ start: manifest.startDate }, { start: manifest.rosterEndDate }] }, null);
          if (!target.stagedNextBatch && !target.stagedPreparedBatch && target.fileId !== run.fileId) {
            if (body.supportsPatch === true && context.env.ROSTER_AUTOMATION_PATCH_IMPORT_ENABLED === 'true') return Response.json({ok:true,mode:"patch",fileId:target.fileId,presenceMode:cachedPresence?'cached':undefined});
            return Response.json({ ok: true, mode: "complete", boundedReplacementAvailable: replacements });
          }
          if(context.env.ROSTER_STAGING_REUSE_SOURCES?.split(',').includes(run.sourceId)){
            const reusable=await reusableRosterStage(db,run,manifest);
            if(reusable)return Response.json({ok:true,mode:'rebase',fileId:reusable,presenceMode:'cached'});
          }
        }
        if (locating) {
          const existing = await db.prepare('SELECT manifest_json FROM roster_import_jobs WHERE run_id=?').bind(run.id).first();
          return Response.json({ok:true,mode:'staging-ready',presenceMode:existing?JSON.parse(existing.manifest_json).presenceMode:(cachedPresence?'cached':undefined)});
        }
        result = await beginBoundedRosterImport(db, run.id, file, { manifest, revision: body.revision }, { allowReplacement: replacements, allowEmptyPlanRecovery: true, validatedPlan:true });
        // An activated job still needs the completion callback if its earlier
        // bookkeeping failed. Only run.status=success means fully completed.
        result = { ...result, activated: Boolean(result.completed), completed: false };
        break;
      }
      case "bounded-events": result = await stageBoundedRosterBatch(db, run.id, file, body.revision, body.batch); break;
      case "bounded-presence": result = await prepareBoundedRosterPresence(db, run.id, file, body.revision, body.batch); break;
      case "bounded-metadata": result = await prepareBoundedRosterMetadata(db, run.id, body.revision, body.batch); break;
      case "bounded-patch-start": result = await beginRosterPatch(db,run,String(body.targetFileId || ""),cachedPresence); break;
      case "bounded-patch-page": result = await readRosterPatchPage(db,run,body); break;
      case "bounded-patch-plan": result = await pinRosterPatchPlan(db,run,body); break;
      case "bounded-patch-stage": result = await stageRosterPatchBatch(db,run,body); break;
      case "bounded-rebase-start":
      case "bounded-rebase-page":
      case "bounded-rebase-plan":
      case "bounded-rebase-save":
      case "bounded-rebase-stage":
      case "bounded-rebase-prepare": {
        if(!context.env.ROSTER_STAGING_REUSE_SOURCES?.split(',').includes(run.sourceId))throw Error('Staging reuse is not enabled for this source.');
        if(body.phase==='bounded-rebase-start')result=await beginRosterRebase(db,run,String(body.targetFileId||''),body.range||{},cachedPresence);
        if(body.phase==='bounded-rebase-page')result=await readRosterRebasePage(db,run,body);
        if(body.phase==='bounded-rebase-plan')result=await pinRosterRebasePlan(db,run,body);
        if(body.phase==='bounded-rebase-save')result=await saveRosterRebaseBatch(db,run,body,context.env.ROSTER_FILES);
        if(body.phase==='bounded-rebase-stage')result=await applyRosterRebaseBatch(db,run,body,context.env.ROSTER_FILES);
        if(body.phase==='bounded-rebase-prepare')result=await prepareRosterRebase(db,run,body,file);
        break;
      }
      case "bounded-finalize":
      case "bounded-patch-activate":
      case "bounded-activate": {
        if (body.phase === 'bounded-finalize') {
          const patch = Boolean(body.patchRevision);
          const job = await db.prepare(patch
            ? 'SELECT file_id,activated,plan_revision FROM roster_patch_jobs WHERE run_id=?'
            : 'SELECT file_id,activated,plan_revision FROM roster_import_jobs WHERE run_id=?').bind(run.id).first();
          if (!job?.activated || job.plan_revision !== (patch ? body.patchRevision : body.revision)) throw Error('Cannot finalize an unfinished or changed roster.');
          result={activated:true,fileId:job.file_id};
        } else {
          result = body.phase === 'bounded-patch-activate'
            ? await activateRosterPatch(db,run,body,{file,deliveryGuard:delivery.guard,publishFacility:automaticFacilityPublicationEnabled(context.env,source.sourceType),historyWindows:body.historyWindows})
            : await activateBoundedRosterTerm(db, run.id, body.revision, { publishFacility: automaticFacilityPublicationEnabled(context.env, source.sourceType), deliveryGuard: delivery.guard, historyWindows:body.historyWindows });
          if (result.deferred) break;
          // Atomic publication is independent of the small, retryable source/run
          // bookkeeping callback. An interrupted callback does not repeat facts.
          if (body.deferCompletion === true) {result={...result,completed:false};break;}
        }
        const targetFileId = result.fileId || run.fileId;
        if(!await activeAuthorityFile(db,targetFileId))return Response.json(await skipSupersededRosterDelivery(db,run));
        const summary = await loadRosterFileStatusSummary(db, targetFileId);
        const current = await loadRosterSource(db, run.sourceId);
        const latest = current?.activeFileId ? await loadRosterFileStatusSummary(db, current.activeFileId) : null;
        const now = new Date().toISOString();
        await upsertRosterSource(db, { ...current, ...source, id: run.sourceId, enabled: true, lastSuccessAt: now, lastError: "", activeFileId: latest?.startDate > summary.startDate ? current.activeFileId : targetFileId, updatedAt: now, createdAt: current?.createdAt || now });
        await supersedeDuplicateRosterSyncRuns(db, run, file.name || "");
        await finishRosterSyncRun(db, run.id, { status: "success", fileId: targetFileId, doctorCount: summary.indexedDoctors, eventCount: summary.eventCount, completedAt: now, message: "Bounded next-term import completed." });
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
