import { findmyshiftConfiguredRosterRange, findmyshiftLastModified, findmyshiftRosterWorkbook } from "../../_lib/findmyshift.js";
import { createRosterSyncRun, findQueuedRosterSyncByHash, findRosterSyncByProviderVersion, hasCalendarDb, listActiveRetainedRosterFiles, loadRosterSource, upsertRosterSource } from "../../_lib/d1-calendar.js";
import { requestQueuedRosterProcessing } from "../../_lib/automation-dispatch.js";
import { reconcileRosterFileSupersessionAndRefresh } from "../state.js";

const SOURCE_ID = "dandenong-findmyshift";
// This version is both part of the retained workbook name and the source
// cursor. Bump it whenever FindMyShift parsing changes, so an unchanged
// provider roster is imported once more with the corrected parser rather than
// leaving its earlier derived events active indefinitely.
const IMPORT_FORMAT = "stream-paired-v2";

export async function onRequestPost(context) {
  if (!hasValidToken(context.request, context.env)) return Response.json({ error: "Unauthorized." }, { status: 401 });
  if (!hasCalendarDb(context.env)) return Response.json({ error: "Roster database is unavailable." }, { status: 503 });
  const apiKey = String(context.env.FINDMYSHIFT_API_KEY || "").trim();
  const teamId = String(context.env.FINDMYSHIFT_TEAM_ID || "").trim();
  if (!apiKey || !teamId) return Response.json({ ok: false, status: "not-configured", error: "FindMyShift API key or team ID is not configured." });

  const now = new Date().toISOString();
  const requestBody = await context.request.json().catch(() => ({}));
  const force = requestBody?.force === true;
  const current = await loadRosterSource(context.env.ROSTER_DB, SOURCE_ID);
  const requestedRange = findmyshiftRequestedRosterRange(requestBody?.range);
  if (requestBody?.range && !requestedRange) {
    return Response.json({ ok: false, status: "invalid-range", error: "A historical FindMyShift range must have valid dates no longer than one term." }, { status: 422 });
  }
  const range = requestedRange || findmyshiftConfiguredRosterRange(context.env);
  let providerVersion = "";
  try {
    providerVersion = await findmyshiftLastModified(apiKey, teamId);
    const rangeState = findmyshiftRangeState(current?.cursor, range, providerVersion);
    const fileName = `Dandenong-FindMyShift-${IMPORT_FORMAT}-${range.from}-to-${range.to}.xlsx`;
    const currentFormatRun = await findRosterSyncByProviderVersion(context.env.ROSTER_DB, SOURCE_ID, providerVersion, fileName);
    // An unpublished upcoming term is checked once per FindMyShift version,
    // then left alone until the provider changes.
    if (current?.providerVersion === providerVersion && rangeState.waiting) {
      await saveSource(context, current, { lastCheckedAt: now, lastError: "" });
      return Response.json({ ok: true, status: "waiting-for-publication", providerModifiedAt: providerVersion });
    }
    // A provider version is current only after it has completed the whole
    // retained-source → background-parser → active-calendar lifecycle.  A
    // failed import must be retried even when FindMyShift has not changed the
    // roster since the failed attempt. A term-window change deliberately
    // bypasses this shortcut so the next term appears four weeks early.
    if (current?.providerVersion
      && current.providerVersion === providerVersion
      && rangeState.requested
      && current.lastSuccessAt
      && currentFormatRun?.status === "success"
      && (!current.lastError || isTransientFindmyshiftRateLimitError(current.lastError))) {
      if (force) {
        const queue = await queueCurrentFindmyshiftReprocess(context.env, providerVersion, now);
        if (!queue.runIds.length) throw new Error("No retained FindMyShift roster file is available to reprocess.");
        await saveSource(context, current, { lastCheckedAt: now, lastError: "" });
        return Response.json({ ok: true, status: "reprocess-queued", providerModifiedAt: providerVersion, queue });
      }
      await reconcileCurrentFindmyshiftRoster(context, current);
      await saveSource(context, current, { lastCheckedAt: now, lastError: "" });
      return Response.json({ ok: true, status: "unchanged", providerModifiedAt: providerVersion });
    }
    // If this exact provider version has already been proved incomplete, do
    // not re-download the full report on every watchdog tick. A future source
    // modification is retried in case the provider starts exposing the stream.
    if (!force && current?.providerVersion === providerVersion && isIncompleteDandenongAssignmentError(current.lastError)) {
      await saveSource(context, current, { lastCheckedAt: now });
      return Response.json({ ok: true, status: "incomplete", providerModifiedAt: providerVersion });
    }
    const workbook = await findmyshiftRosterWorkbook(apiKey, teamId, range);
    const response = await fetch(new URL("/api/automation/ingest", context.request.url), {
      method: "POST",
      headers: { Authorization: `Bearer ${String(context.env.ROSTER_AUTOMATION_TOKEN || "")}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        sourceId: SOURCE_ID,
        fileName,
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        contentBase64: bytesToBase64(new Uint8Array(workbook)),
        providerVersion,
        providerModifiedAt: providerVersion,
      }),
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(String(result.error || `Roster queue returned HTTP ${response.status}.`));
    await saveSource(context, current, {
      providerVersion,
      providerModifiedAt: providerVersion,
      lastCheckedAt: now,
      lastError: "",
      cursor: withFindmyshiftRangeState(current?.cursor, range, providerVersion, "queued"),
    });
    return Response.json({ ok: true, status: String(result.status || "queued"), providerModifiedAt: providerVersion, queue: { runId: String(result.runId || ""), dispatched: result.processorDispatch?.dispatched === true } });
  } catch (error) {
    const errorMessage = String(error?.message || error).slice(0, 300);
    if (error?.code === "findmyshift-no-shifts") {
      await saveSource(context, current, {
        providerVersion: providerVersion || current?.providerVersion || "",
        providerModifiedAt: providerVersion || current?.providerModifiedAt || "",
        lastCheckedAt: now,
        lastError: "",
        cursor: withFindmyshiftRangeState(current?.cursor, range, providerVersion, "waiting"),
      });
      return Response.json({
        ok: true,
        status: "waiting-for-publication",
        providerModifiedAt: providerVersion,
      });
    }
    // A completed source remains current when the provider merely throttles a
    // metadata poll. The next scheduled check will retry it; do not make the
    // Files card look like the successfully imported roster has failed.
    const lastError = isTransientFindmyshiftRateLimitError(errorMessage) && current?.lastSuccessAt ? "" : errorMessage;
    await saveSource(context, current, { lastCheckedAt: now, lastError });
    const incomplete = error?.code === "findmyshift-incomplete-ddh-assignment" || isIncompleteDandenongAssignmentError(error?.message);
    return Response.json({
      ok: false,
      status: incomplete ? "incomplete" : "failed",
      error: incomplete
        ? "FindMyShift did not provide Dandenong stream details for every timed shift, so no ambiguous roster was imported."
        : "FindMyShift roster check failed.",
    }, { status: 502 });
  }
}

function findmyshiftRequestedRosterRange(value) {
  if (!value || typeof value !== "object") return null;
  const from = String(value.from || "").trim();
  const to = String(value.to || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to) || from > to) return null;
  const fromDate = new Date(`${from}T00:00:00Z`);
  const toDate = new Date(`${to}T00:00:00Z`);
  const durationDays = Math.round((toDate.getTime() - fromDate.getTime()) / 86_400_000) + 1;
  return durationDays > 0 && durationDays <= 100 ? { from, to } : null;
}

async function queueCurrentFindmyshiftReprocess(env, providerVersion, now) {
  const files = (await listActiveRetainedRosterFiles(env.ROSTER_DB).catch(() => []))
    .filter((file) => String(file.sourceId || "").trim() === SOURCE_ID);
  const runIds = [];
  for (const file of files) {
    const contentHash = `creator-reprocess:${file.id}:${providerVersion}:${IMPORT_FORMAT}`;
    const existing = await findQueuedRosterSyncByHash(env.ROSTER_DB, SOURCE_ID, contentHash);
    if (existing) {
      runIds.push(existing.id);
      continue;
    }
    const runId = `reprocess:${file.id}:${crypto.randomUUID()}`;
    const stagingFileId = `staged:${file.id}:${crypto.randomUUID()}`;
    await createRosterSyncRun(env.ROSTER_DB, {
      id: runId,
      sourceId: SOURCE_ID,
      triggerType: "creator-reprocess",
      providerVersion: contentHash,
      contentHash,
      fileId: stagingFileId,
      sourceFileId: file.id,
      status: "queued",
      message: "Queued to reprocess the retained FindMyShift roster file.",
      startedAt: now,
    });
    runIds.push(runId);
  }
  const dispatch = runIds.length
    ? await requestQueuedRosterProcessing(env, { reason: "creator-findmyshift-reprocess" })
    : null;
  return {
    runIds,
    dispatched: dispatch?.dispatched === true,
  };
}

function isIncompleteDandenongAssignmentError(value) {
  return /did not include a stream or facility/i.test(String(value || ""));
}

function isTransientFindmyshiftRateLimitError(value) {
  return /FindMyShift .* returned HTTP 429\./i.test(String(value || ""));
}

function findmyshiftRangeState(cursor, range, providerVersion) {
  const saved = cursor && typeof cursor === "object" ? cursor.findmyshiftRange : null;
  const requested = String(saved?.from || "") === String(range?.from || "")
    && String(saved?.to || "") === String(range?.to || "")
    && String(saved?.providerVersion || "") === String(providerVersion || "")
    && String(saved?.importFormat || "") === IMPORT_FORMAT;
  return { requested, waiting: requested && saved?.status === "waiting" };
}

function withFindmyshiftRangeState(cursor, range, providerVersion, status) {
  return {
    ...(cursor && typeof cursor === "object" ? cursor : {}),
    findmyshiftRange: {
      from: String(range?.from || ""),
      to: String(range?.to || ""),
      providerVersion: String(providerVersion || ""),
      importFormat: IMPORT_FORMAT,
      status: String(status || ""),
    },
  };
}

async function reconcileCurrentFindmyshiftRoster(context, current) {
  const activeFileId = String(current?.activeFileId || "").trim();
  if (!activeFileId) return;
  await reconcileRosterFileSupersessionAndRefresh(context, {
    id: activeFileId,
    sourceType: "ddh",
    sourceId: SOURCE_ID,
  }, { reason: "findmyshift-unchanged-reconciliation" });
}

async function saveSource(context, existing, update) {
  const now = new Date().toISOString();
  // Re-read so a just-finished background import cannot have its active-file
  // and success metadata overwritten by this lightweight watchdog update.
  const current = await loadRosterSource(context.env.ROSTER_DB, SOURCE_ID).catch(() => null);
  const base = current || existing || {};
  await upsertRosterSource(context.env.ROSTER_DB, {
    ...base, id: SOURCE_ID, provider: "findmyshift", sourceType: "ddh", label: "Dandenong (Findmyshift)", enabled: true,
    config: base.config || {}, cursor: base.cursor || {}, createdAt: base.createdAt || now, updatedAt: now, ...update,
  });
}

function hasValidToken(request, env) {
  const provided = String(request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  return [env.ROSTER_AUTOMATION_TOKEN, env.ROSTER_WATCHDOG_TOKEN].some((candidate) => timingSafeEqual(provided, String(candidate || "")));
}

function timingSafeEqual(left, right) {
  if (!left || !right || left.length !== right.length) return false;
  let mismatch = 0;
  for (let index = 0; index < left.length; index += 1) mismatch |= left.charCodeAt(index) ^ right.charCodeAt(index);
  return mismatch === 0;
}

function bytesToBase64(bytes) {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}
