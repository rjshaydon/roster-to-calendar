import { automationSourceDefinition } from "../../_lib/automation-import.js";
import {
  finishRosterSyncRun,
  hasCalendarDb,
  loadRosterSource,
  loadRosterSyncRun,
  markRosterSyncRunProcessing,
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
    if (!source || !run || run.sourceId !== sourceId || run.fileId !== String(body?.file?.id || "")) {
      return Response.json({ error: "Queued roster job does not match the derived payload." }, { status: 400 });
    }
    if (!["start", "events", "finish", "failed"].includes(phase)) {
      return Response.json({ error: "A valid derived-save phase is required." }, { status: 400 });
    }
    if (phase === "failed") {
      const failedAt = new Date().toISOString();
      await finishRosterSyncRun(context.env.ROSTER_DB, runId, {
        status: "failed",
        fileId: run.fileId,
        message: String(body?.message || "Background roster processing failed.").slice(0, 300),
        completedAt: failedAt,
      });
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
      return Response.json({ ok: true, phase, runId, fileId: run.fileId });
    }
    if (phase === "start") await markRosterSyncRunProcessing(context.env.ROSTER_DB, runId);
    const saved = await runAutomatedDerivedRosterSave(context, {
      phase,
      file: { ...body.file, sourceId, sourceType: source.sourceType },
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
      const existing = await loadRosterSource(context.env.ROSTER_DB, sourceId);
      await upsertRosterSource(context.env.ROSTER_DB, {
        ...(existing || {}),
        ...source,
        id: sourceId,
        enabled: true,
        lastSuccessAt: completedAt,
        lastError: "",
        activeFileId: run.fileId,
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
    if (runId) {
      await finishRosterSyncRun(context.env.ROSTER_DB, runId, {
        status: "failed",
        fileId: String(body?.file?.id || ""),
        message: "Background roster processing failed.",
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
          lastError: "Background roster processing failed.",
          updatedAt: failedAt,
          createdAt: existing?.createdAt || failedAt,
        }).catch(() => null);
      }
    }
    console.error("Derived roster processing failed", error);
    return Response.json({ error: "Background roster processing failed." }, { status: 422 });
  }
}

function hasValidAutomationToken(request, configuredToken) {
  const token = String(configuredToken || "");
  const provided = String(request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!token || !provided || token.length !== provided.length) return false;
  let mismatch = 0;
  for (let index = 0; index < token.length; index += 1) mismatch |= token.charCodeAt(index) ^ provided.charCodeAt(index);
  return mismatch === 0;
}
