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

const MAX_ROSTER_BYTES = 15 * 1024 * 1024;

export async function onRequestPost(context) {
  if (!hasValidAutomationToken(context.request, context.env.ROSTER_AUTOMATION_TOKEN)) {
    return Response.json({ error: "Unauthorized." }, { status: 401 });
  }
  if (!hasCalendarDb(context.env)) return Response.json({ error: "Roster database is unavailable." }, { status: 503 });
  try {
    const upload = await readAutomationUpload(context.request);
    const sourceId = upload.sourceId;
    const source = automationSourceDefinition(sourceId);
    const file = upload.file;
    const providerVersion = upload.providerVersion;
    const providerModifiedAt = upload.providerModifiedAt;
    if (!source || !(file instanceof File)) return Response.json({ error: "A configured source and roster file are required." }, { status: 400 });
    if (!file.size || file.size > MAX_ROSTER_BYTES) return Response.json({ error: "Roster file is missing or too large." }, { status: 413 });
    if (!context.env.ROSTER_FILES?.put) return Response.json({ error: "Roster file storage is unavailable." }, { status: 503 });

    const now = new Date().toISOString();
    const sourceRecord = await loadRosterSource(context.env.ROSTER_DB, sourceId);
    const matchingVersion = providerVersion
      ? await findRosterSyncByProviderVersion(context.env.ROSTER_DB, sourceId, providerVersion, file.name)
      : null;
    // A failed version is deliberately re-queued.  It may have been retained
    // successfully while a transient background save failed; treating it as
    // unchanged would otherwise leave that retained roster stranded forever.
    if (matchingVersion && matchingVersion.status !== "failed") {
      await upsertRosterSource(context.env.ROSTER_DB, updatedSourceRecord(sourceRecord, source, {
        id: sourceId,
        providerVersion,
        providerModifiedAt: sourceRecord?.providerModifiedAt || providerModifiedAt,
        lastCheckedAt: now,
        updatedAt: now,
      }));
      const status = matchingVersion.status === "success"
        ? "unchanged"
        : matchingVersion.status === "failed"
          ? "unchanged-failed"
          : matchingVersion.status;
      const dispatch = ["queued", "processing"].includes(status)
        ? await requestQueuedRosterProcessing(context.env, { reason: "duplicate-queue-check" })
        : null;
      return Response.json({
        ok: true,
        status,
        sourceId,
        fileId: matchingVersion.fileId,
        runId: matchingVersion.id,
        processorDispatch: publicDispatchStatus(dispatch),
      }, { status: ["queued", "processing"].includes(status) ? 202 : 200 });
    }

    const bytes = new Uint8Array(await file.arrayBuffer());
    const contentHash = await sha256Hex(bytes);
    await upsertRosterSource(context.env.ROSTER_DB, updatedSourceRecord(sourceRecord, source, {
      id: sourceId, providerVersion, providerModifiedAt, lastCheckedAt: now, lastError: "", updatedAt: now,
    }));
    // The parser/import format is part of the retained filename. Identical
    // workbook bytes must be processed again when that filename changes for a
    // parser revision; otherwise a corrected parser can never replace old
    // derived events until the roster provider changes the workbook itself.
    const prior = await findSuccessfulRosterSyncByHash(context.env.ROSTER_DB, sourceId, contentHash, file.name);
    if (prior) {
      return Response.json({ ok: true, status: "unchanged", sourceId, fileId: prior.fileId, runId: prior.id });
    }
    const queued = await findQueuedRosterSyncByHash(context.env.ROSTER_DB, sourceId, contentHash, file.name);
    if (queued) {
      const dispatch = await requestQueuedRosterProcessing(context.env, { reason: "duplicate-content-check" });
      return Response.json({ ok: true, status: queued.status, sourceId, fileId: queued.fileId, runId: queued.id, processorDispatch: publicDispatchStatus(dispatch) }, { status: 202 });
    }

    const runId = `sync:${sourceId}:${crypto.randomUUID()}`;
    const fileId = `automation:${sourceId}:${contentHash.slice(0, 24)}`;
    const objectKey = `automation/${sourceId}/${contentHash}`;
    await context.env.ROSTER_FILES.put(objectKey, bytes, { httpMetadata: { contentType: file.type || "application/octet-stream" } });
    await upsertRawRosterFile(context.env.ROSTER_DB, {
      id: fileId,
      name: file.name || `${sourceId}.xlsx`,
      sourceType: source.sourceType,
      size: file.size,
      lastModified: file.lastModified,
    }, {
      objectKey,
      type: file.type || "application/octet-stream",
      uploadedAt: now,
    });
    await createRosterSyncRun(context.env.ROSTER_DB, {
      id: runId,
      sourceId,
      triggerType: "sharepoint",
      providerVersion,
      contentHash,
      fileId,
      status: "queued",
      message: "Queued for background processing.",
      startedAt: now,
    });
    const dispatch = await requestQueuedRosterProcessing(context.env, { reason: "source-update" });
    return Response.json({ ok: true, status: "queued", sourceId, fileId, runId, processorDispatch: publicDispatchStatus(dispatch) }, { status: 202 });
  } catch (error) {
    console.error("Automated roster queueing failed", error);
    return Response.json({ error: "Roster could not be queued." }, { status: 422 });
  }
}

function publicDispatchStatus(result) {
  if (!result) return null;
  return {
    dispatched: result.dispatched === true,
    reason: String(result.reason || ""),
    status: String(result.dispatch?.status || ""),
    requestedAt: String(result.dispatch?.requestedAt || ""),
    lastError: String(result.dispatch?.lastError || ""),
  };
}

async function readAutomationUpload(request) {
  const contentType = String(request.headers.get("content-type") || "").toLowerCase();
  if (contentType.includes("application/json")) {
    const body = await request.json();
    const base64 = String(body?.contentBase64 || "").replace(/^data:[^;,]+;base64,/i, "");
    if (!base64) return { sourceId: "", providerVersion: "", providerModifiedAt: "", file: null };
    const binary = atob(base64);
    const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
    const providerModifiedAt = String(body?.providerModifiedAt || body?.lastModified || "").trim();
    const explicitLastModified = Number(body?.lastModified);
    const providerLastModified = Date.parse(providerModifiedAt);
    return {
      sourceId: String(body?.sourceId || "").trim(),
      providerVersion: String(body?.providerVersion || "").trim(),
      providerModifiedAt,
      file: new File([bytes], String(body?.fileName || "roster.xlsx"), {
        type: String(body?.contentType || "application/octet-stream"),
        lastModified: Number.isFinite(explicitLastModified) && explicitLastModified > 0
          ? explicitLastModified
          : Number.isFinite(providerLastModified) ? providerLastModified : Date.now(),
      }),
    };
  }
  const formData = await request.formData();
  return {
    sourceId: String(formData.get("sourceId") || "").trim(),
    providerVersion: String(formData.get("providerVersion") || "").trim(),
    providerModifiedAt: String(formData.get("providerModifiedAt") || formData.get("lastModified") || "").trim(),
    file: formData.get("rosterFile"),
  };
}

function updatedSourceRecord(existing, definition, update = {}) {
  return {
    ...(existing || {}),
    ...definition,
    ...update,
    enabled: true,
    config: existing?.config || {},
    cursor: existing?.cursor || {},
    createdAt: existing?.createdAt || new Date().toISOString(),
  };
}

function hasValidAutomationToken(request, configuredToken) {
  const token = String(configuredToken || "");
  const provided = String(request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!token || !provided || token.length !== provided.length) return false;
  let mismatch = 0;
  for (let index = 0; index < token.length; index += 1) mismatch |= token.charCodeAt(index) ^ provided.charCodeAt(index);
  return mismatch === 0;
}
