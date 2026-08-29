import { automationSourceDefinition } from "../../_lib/automation-import.js";
import {
  deleteDerivedRosterFile,
  finishRosterSyncRun,
  hasCalendarDb,
  loadRosterSource,
  loadRosterSyncRun,
  markRosterSyncRunProcessing,
  supersedeDuplicateRosterSyncRuns,
  upsertRosterSource,
} from "../../_lib/d1-calendar.js";
import { runAutomatedDerivedRosterSave } from "../state.js";

export async function onRequestPost(context) {
  if (!hasValidAutomationToken(context.request, context.env.ROSTER_AUTOMATION_TOKEN)) {
    return Response.json({ error: "Unauthorized." }, { status: 401 });
  }
  if (!hasCalendarDb(context.env)) return Response.json({ error: "Roster database is unavailable." }, { status: 503 });
  let body = null;
  try {
    body = await context.request.json();
    const runId = String(body?.runId || "").trim();
    const sourceId = String(body?.sourceId || body?.file?.sourceId || "").trim();
    const phase = String(body?.phase || "").toLowerCase();
    const source = automationSourceDefinition(sourceId);
    const run = await loadRosterSyncRun(context.env.ROSTER_DB, runId);
    if (!run || run.sourceId !== sourceId || run.fileId !== String(body?.file?.id || "")) {
      return Response.json({ error: "Queued roster job does not match the derived payload." }, { status: 400 });
    }
    if (!["start", "events", "finish", "failed"].includes(phase)) {
      return Response.json({ error: "A valid derived-save phase is required." }, { status: 400 });
    }
    if (phase === "failed") {
      const failedAt = new Date().toISOString();
      // An interrupted chunked import may have saved only some events.  Those
      // rows must never become an active calendar source; retain the original
      // workbook in R2 for retry, but remove the incomplete derived copy.
      await deleteDerivedRosterFile(context.env.ROSTER_DB, run.fileId);
      await finishRosterSyncRun(context.env.ROSTER_DB, runId, {
        status: "failed",
        fileId: run.fileId,
        message: String(body?.message || "Background roster processing failed.").slice(0, 300),
        completedAt: failedAt,
      });
      await supersedeDuplicateRosterSyncRuns(context.env.ROSTER_DB, run, body?.file?.name || "");
      // Failure reporting must remain available even when the deployed
      // processor does not yet recognise a newly queued source. Otherwise the
      // run remains queued and the watchdog dispatches it forever.
      if (source) {
        const existing = await loadRosterSource(context.env.ROSTER_DB, sourceId);
        await upsertRosterSource(context.env.ROSTER_DB, {
          ...(existing || {}),
          ...source,
          id: sourceId,
          enabled: true,
          lastError: "Background roster processing failed.",
          updatedAt: failedAt,
          createdAt: existing?.createdAt || failedAt,
        });
      }
      return Response.json({ ok: true, phase, runId, fileId: run.fileId });
    }
    if (!source) {
      return Response.json({ error: "Unknown automation source." }, { status: 400 });
    }
    if (phase === "start") await markRosterSyncRunProcessing(context.env.ROSTER_DB, runId);
    const saved = await runAutomatedDerivedRosterSave(context, {
      phase,
      // A queued source is invisible to calendars until its final phase. A
      // retained-file reparse uses a distinct staging id and names the active
      // source it may replace only after its event comparison passes.
      file: {
        ...body.file,
        sourceId,
        sourceType: source.sourceType,
        active: false,
        staged: true,
        replacesFileId: run.sourceFileId && run.sourceFileId !== run.fileId ? run.sourceFileId : "",
      },
      doctors: Array.isArray(body.doctors) ? body.doctors : [],
      eventsByDoctor: body.eventsByDoctor && typeof body.eventsByDoctor === "object" ? body.eventsByDoctor : {},
      issuesByDoctor: body.issuesByDoctor && typeof body.issuesByDoctor === "object" ? body.issuesByDoctor : {},
    });
    if (phase === "finish") {
      const completedAt = new Date().toISOString();
      const doctorCount = Number(saved?.result?.doctors || 0);
      const eventCount = Number(saved?.result?.events || 0);
      await finishRosterSyncRun(context.env.ROSTER_DB, runId, {
        status: "success",
        fileId: run.fileId,
        doctorCount,
        eventCount,
        message: "Roster indexed by background processor.",
        completedAt,
      });
      await supersedeDuplicateRosterSyncRuns(context.env.ROSTER_DB, run, body?.file?.name || "");
      const existing = await loadRosterSource(context.env.ROSTER_DB, sourceId);
      // A reparse can process several historical retained files. It must not
      // move an automated source's current-file pointer to the last one that
      // happens to finish.
      const preserveActiveFile = ["parser-rule", "creator-reprocess"].includes(run.triggerType)
        && existing?.activeFileId
        && existing.activeFileId !== run.fileId;
      await upsertRosterSource(context.env.ROSTER_DB, {
        ...(existing || {}),
        ...source,
        id: sourceId,
        enabled: true,
        lastSuccessAt: completedAt,
        lastError: "",
        activeFileId: preserveActiveFile ? existing.activeFileId : run.fileId,
        updatedAt: completedAt,
        createdAt: existing?.createdAt || completedAt,
      });
      return Response.json({ ok: true, phase, runId, fileId: run.fileId, doctorCount, eventCount });
    }
    return Response.json({ ok: true, phase, runId, result: saved?.result || null });
  } catch (error) {
    const runId = String(body?.runId || "").trim();
    const sourceId = String(body?.sourceId || body?.file?.sourceId || "").trim();
    const failedAt = new Date().toISOString();
    const diagnostic = String(error?.message || "Background roster processing failed.").slice(0, 300);
    if (runId) {
      await deleteDerivedRosterFile(context.env.ROSTER_DB, String(body?.file?.id || "")).catch(() => null);
      await finishRosterSyncRun(context.env.ROSTER_DB, runId, {
        status: "failed",
        fileId: String(body?.file?.id || ""),
        message: diagnostic,
        completedAt: failedAt,
      }).catch(() => null);
    }
    if (sourceId) {
      const existing = await loadRosterSource(context.env.ROSTER_DB, sourceId).catch(() => null);
      const definition = automationSourceDefinition(sourceId);
      if (definition) {
        await upsertRosterSource(context.env.ROSTER_DB, {
          ...(existing || {}),
          ...definition,
          id: sourceId,
          enabled: true,
          lastError: diagnostic,
          updatedAt: failedAt,
          createdAt: existing?.createdAt || failedAt,
        }).catch(() => null);
      }
    }
    const code = safeDerivedFailureCode(error);
    console.error("Derived roster processing failed", { code, message: diagnostic });
    return Response.json({ error: diagnostic, code }, { status: 422 });
  }
}

function safeDerivedFailureCode(error) {
  const message = String(error?.message || error || "").toLowerCase();
  if (message.includes("d1") || message.includes("sqlite") || message.includes("database")) return "database-save";
  if (message.includes("event") || message.includes("roster save")) return "event-save";
  if (message.includes("supersession")) return "supersession";
  if (message.includes("timeout") || message.includes("timed out")) return "timeout";
  return "derived-save";
}

function hasValidAutomationToken(request, configuredToken) {
  const token = String(configuredToken || "");
  const provided = String(request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!token || !provided || token.length !== provided.length) return false;
  let mismatch = 0;
  for (let index = 0; index < token.length; index += 1) mismatch |= token.charCodeAt(index) ^ provided.charCodeAt(index);
  return mismatch === 0;
}
