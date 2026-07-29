import { buildAutomatedDerivedRosterPayload, automationSourceDefinition, sha256Hex } from "../../_lib/automation-import.js";
import {
  createRosterSyncRun,
  findSuccessfulRosterSyncByHash,
  finishRosterSyncRun,
  hasCalendarDb,
  loadRosterSource,
  upsertRawRosterFile,
  upsertRosterSource,
} from "../../_lib/d1-calendar.js";
import { runAutomatedDerivedRosterSave } from "../state.js";

const MAX_ROSTER_BYTES = 15 * 1024 * 1024;

export async function onRequestPost(context) {
  if (!hasValidAutomationToken(context.request, context.env.ROSTER_AUTOMATION_TOKEN)) {
    return Response.json({ error: "Unauthorized." }, { status: 401 });
  }
  if (!hasCalendarDb(context.env)) return Response.json({ error: "Roster database is unavailable." }, { status: 503 });
  let sourceId = "";
  let source = null;
  let sourceRecord = null;
  let runId = "";
  let providerVersion = "";
  let providerModifiedAt = "";
  let now = "";
  try {
    const upload = await readAutomationUpload(context.request);
    sourceId = upload.sourceId;
    source = automationSourceDefinition(sourceId);
    const file = upload.file;
    providerVersion = upload.providerVersion;
    providerModifiedAt = upload.providerModifiedAt;
    if (!source || !(file instanceof File)) return Response.json({ error: "A configured source and roster file are required." }, { status: 400 });
    if (!file.size || file.size > MAX_ROSTER_BYTES) return Response.json({ error: "Roster file is missing or too large." }, { status: 413 });

    const bytes = new Uint8Array(await file.arrayBuffer());
    const contentHash = await sha256Hex(bytes);
    now = new Date().toISOString();
    sourceRecord = await loadRosterSource(context.env.ROSTER_DB, sourceId);
    await upsertRosterSource(context.env.ROSTER_DB, updatedSourceRecord(sourceRecord, source, {
      id: sourceId, providerVersion, providerModifiedAt, lastCheckedAt: now, lastError: "", updatedAt: now,
    }));
    const prior = await findSuccessfulRosterSyncByHash(context.env.ROSTER_DB, sourceId, contentHash);
    if (prior) {
      return Response.json({ ok: true, status: "unchanged", sourceId, fileId: prior.fileId, runId: prior.id });
    }

    runId = `sync:${sourceId}:${crypto.randomUUID()}`;
    await createRosterSyncRun(context.env.ROSTER_DB, {
      id: runId, sourceId, triggerType: "ingest", providerVersion, contentHash, startedAt: now,
    });
    const derived = await buildAutomatedDerivedRosterPayload({ file, sourceId, contentHash, providerVersion });
    const objectKey = `automation/${sourceId}/${contentHash}`;
    if (context.env.ROSTER_FILES?.put) {
      await context.env.ROSTER_FILES.put(objectKey, bytes, { httpMetadata: { contentType: file.type || "application/octet-stream" } });
    }
    await upsertRawRosterFile(context.env.ROSTER_DB, derived.file, {
      objectKey: context.env.ROSTER_FILES?.put ? objectKey : "",
      dataUrl: "",
      type: file.type || "application/octet-stream",
      uploadedAt: now,
    });
    const saved = await runAutomatedDerivedRosterSave(context, derived);
    const doctors = Number(saved?.result?.doctors || derived.doctors.length || 0);
    const events = Number(saved?.result?.events || derived.eventCount || 0);
    await finishRosterSyncRun(context.env.ROSTER_DB, runId, {
      status: "success", fileId: derived.file.id, doctorCount: doctors, eventCount: events,
      message: "Roster indexed.", completedAt: new Date().toISOString(),
    });
    await upsertRosterSource(context.env.ROSTER_DB, {
      ...updatedSourceRecord(sourceRecord, source, {
        id: sourceId, providerVersion, providerModifiedAt, lastCheckedAt: now,
        lastSuccessAt: new Date().toISOString(), lastError: "", activeFileId: derived.file.id, updatedAt: new Date().toISOString(),
      }),
    });
    return Response.json({ ok: true, status: "imported", sourceId, fileId: derived.file.id, runId, doctors, events });
  } catch (error) {
    console.error("Automated roster ingestion failed", error);
    const failedAt = new Date().toISOString();
    if (runId) {
      await finishRosterSyncRun(context.env.ROSTER_DB, runId, {
        status: "failed", message: "Roster ingestion failed.", completedAt: failedAt,
      }).catch(() => null);
    }
    if (sourceId && source) {
      await upsertRosterSource(context.env.ROSTER_DB, updatedSourceRecord(sourceRecord, source, {
        id: sourceId, providerVersion, providerModifiedAt, lastCheckedAt: now || failedAt,
        lastError: "Roster ingestion failed.", updatedAt: failedAt,
      })).catch(() => null);
    }
    return Response.json({ error: "Roster ingestion failed." }, { status: 422 });
  }
}

async function readAutomationUpload(request) {
  const contentType = String(request.headers.get("content-type") || "").toLowerCase();
  if (contentType.includes("application/json")) {
    const body = await request.json();
    const base64 = String(body?.contentBase64 || "").replace(/^data:[^;,]+;base64,/i, "");
    if (!base64) return { sourceId: "", providerVersion: "", providerModifiedAt: "", file: null };
    const binary = atob(base64);
    const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
    return {
      sourceId: String(body?.sourceId || "").trim(),
      providerVersion: String(body?.providerVersion || "").trim(),
      providerModifiedAt: String(body?.providerModifiedAt || body?.lastModified || "").trim(),
      file: new File([bytes], String(body?.fileName || "roster.xlsx"), {
        type: String(body?.contentType || "application/octet-stream"),
        lastModified: Number(body?.lastModified || Date.now()),
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
