import { automationSourceDefinition, sha256Hex } from "../../_lib/automation-import.js";
import { requestQueuedRosterProcessing } from "../../_lib/automation-dispatch.js";
import {
  createRosterSyncRun,
  findQueuedRosterSyncByHash,
  findRosterSyncByProviderVersion,
  findSuccessfulRosterSyncByHash,
  hasCalendarDb,
  loadRosterSource,
  upsertRawRosterFile,
  upsertRosterSource,
} from "../../_lib/d1-calendar.js";
import { normaliseVhhRosterExtract, VHH_ROSTER_SOURCE_ID } from "../../_lib/vhh-roster.js";
import { automatedRosterWritesEnabled, rosterAutomationPausedResponse } from "../../_lib/roster-automation-guard.js";

const MAX_BODY_BYTES = 1024 * 1024;

// Receives only the Office Script roster extract. The SharePoint workbook is
// never uploaded, altered, or retained by this application.
export async function onRequestPost(context) {
  if (!hasValidAutomationToken(context.request, context.env.ROSTER_AUTOMATION_TOKEN, context.env.VHH_AUTOMATION_TOKEN)) return Response.json({ error: "Unauthorized." }, { status: 401 });
  if (!automatedRosterWritesEnabled(context.env)) return rosterAutomationPausedResponse();
  if (!hasCalendarDb(context.env) || !context.env.ROSTER_FILES?.put) return Response.json({ error: "Roster storage is unavailable." }, { status: 503 });
  try {
    const contentLength = Number(context.request.headers.get("content-length") || "0");
    if (Number.isFinite(contentLength) && contentLength > MAX_BODY_BYTES) return Response.json({ error: "VHH roster JSON is too large." }, { status: 413 });
    const extract = normaliseVhhRosterExtract(await context.request.json());
    if (!extract) return Response.json({ error: "Invalid VHH roster extract." }, { status: 400 });
    const source = automationSourceDefinition(VHH_ROSTER_SOURCE_ID);
    if (!source) return Response.json({ error: "VHH automation source is not configured." }, { status: 500 });
    const bytes = new TextEncoder().encode(JSON.stringify(extract));
    if (bytes.byteLength > MAX_BODY_BYTES) return Response.json({ error: "VHH roster JSON is too large." }, { status: 413 });

    const now = new Date().toISOString();
    const db = context.env.ROSTER_DB;
    const sourceRecord = await loadRosterSource(db, VHH_ROSTER_SOURCE_ID);
    const matchingVersion = extract.providerVersion
      ? await findRosterSyncByProviderVersion(db, VHH_ROSTER_SOURCE_ID, extract.providerVersion, extract.fileName)
      : null;
    if (matchingVersion && matchingVersion.status !== "failed") {
      await upsertRosterSource(db, updatedSourceRecord(sourceRecord, source, {
        id: VHH_ROSTER_SOURCE_ID, providerVersion: extract.providerVersion,
        providerModifiedAt: sourceRecord?.providerModifiedAt || extract.providerModifiedAt,
        lastCheckedAt: now, updatedAt: now,
      }));
      return Response.json({ ok: true, status: matchingVersion.status === "success" ? "unchanged" : matchingVersion.status, sourceId: VHH_ROSTER_SOURCE_ID, fileId: matchingVersion.fileId, runId: matchingVersion.id });
    }

    const contentHash = await sha256Hex(bytes);
    await upsertRosterSource(db, updatedSourceRecord(sourceRecord, source, {
      id: VHH_ROSTER_SOURCE_ID, providerVersion: extract.providerVersion, providerModifiedAt: extract.providerModifiedAt,
      lastCheckedAt: now, lastError: "", updatedAt: now,
    }));
    const prior = await findSuccessfulRosterSyncByHash(db, VHH_ROSTER_SOURCE_ID, contentHash, extract.fileName);
    if (prior) return Response.json({ ok: true, status: "unchanged", sourceId: VHH_ROSTER_SOURCE_ID, fileId: prior.fileId, runId: prior.id });
    const queued = await findQueuedRosterSyncByHash(db, VHH_ROSTER_SOURCE_ID, contentHash, extract.fileName);
    if (queued) return Response.json({ ok: true, status: queued.status, sourceId: VHH_ROSTER_SOURCE_ID, fileId: queued.fileId, runId: queued.id }, { status: 202 });

    const runId = `sync:${VHH_ROSTER_SOURCE_ID}:${crypto.randomUUID()}`;
    const fileId = `automation:${VHH_ROSTER_SOURCE_ID}:${contentHash.slice(0, 24)}`;
    const objectKey = `automation/${VHH_ROSTER_SOURCE_ID}/${contentHash}.json`;
    await context.env.ROSTER_FILES.put(objectKey, bytes, { httpMetadata: { contentType: "application/json; charset=utf-8" } });
    await upsertRawRosterFile(db, { id: fileId, name: extract.fileName, sourceType: "vhh", size: bytes.byteLength, lastModified: Date.parse(extract.providerModifiedAt) || Date.now() }, {
      objectKey, type: "application/json; charset=utf-8", uploadedAt: now,
    });
    await createRosterSyncRun(db, { id: runId, sourceId: VHH_ROSTER_SOURCE_ID, triggerType: "sharepoint-json", providerVersion: extract.providerVersion, contentHash, fileId, status: "queued", message: "Queued for VHH JSON processing.", startedAt: now });
    const dispatch = await requestQueuedRosterProcessing(context.env, { reason: "vhh-source-update" });
    return Response.json({ ok: true, status: "queued", sourceId: VHH_ROSTER_SOURCE_ID, fileId, runId, processorDispatch: dispatch?.dispatched === true }, { status: 202 });
  } catch (error) {
    console.error("VHH roster JSON queueing failed", error);
    return Response.json({ error: "VHH roster JSON could not be queued." }, { status: 422 });
  }
}

function updatedSourceRecord(existing, definition, update = {}) {
  return { ...(existing || {}), ...definition, ...update, enabled: true, config: existing?.config || {}, cursor: existing?.cursor || {}, createdAt: existing?.createdAt || new Date().toISOString() };
}

function hasValidAutomationToken(request, ...configuredTokens) {
  const provided = String(request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!provided) return false;
  return configuredTokens.some((configuredToken) => {
    const token = String(configuredToken || "");
    if (!token || token.length !== provided.length) return false;
    let mismatch = 0;
    for (let index = 0; index < token.length; index += 1) mismatch |= token.charCodeAt(index) ^ provided.charCodeAt(index);
    return mismatch === 0;
  });
}
