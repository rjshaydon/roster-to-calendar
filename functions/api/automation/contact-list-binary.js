import { ensureCalendarSchema, hasCalendarDb } from "../../_lib/d1-calendar.js";

const SOURCE_ID = "mmc-shift-allocations";
const EXPECTED_FILE_NAME = "shift allocations.xlsx";
const MAX_CONTACT_LIST_BYTES = 50 * 1024 * 1024;

// SHIFT ALLOCATIONS.xlsx is 28 MB. Encoding it as JSON Base64 causes the
// Pages Function to buffer far more than the file itself. This ingress takes
// the binary body and streams it directly into R2 instead.
export async function onRequestPost(context) {
  if (!hasValidAutomationToken(context.request, context.env.ROSTER_AUTOMATION_TOKEN)) {
    return Response.json({ error: "Unauthorized." }, { status: 401 });
  }
  if (!hasCalendarDb(context.env) || !context.env.ROSTER_FILES?.put) {
    return Response.json({ error: "Contact-list storage is unavailable." }, { status: 503 });
  }

  const params = new URL(context.request.url).searchParams;
  const sourceId = requestMetadata(context.request, params, "x-roster-source-id", "sourceId");
  const fileName = requestMetadata(context.request, params, "x-roster-file-name", "fileName");
  const providerVersion = requestMetadata(context.request, params, "x-roster-provider-version", "providerVersion");
  const providerModifiedAt = requestMetadata(context.request, params, "x-roster-provider-modified-at", "providerModifiedAt");
  const contentLength = Number(context.request.headers.get("content-length") || "0");
  if (sourceId !== SOURCE_ID) return Response.json({ error: "Unknown contact-list source." }, { status: 400 });
  if (normaliseFileName(fileName) !== EXPECTED_FILE_NAME) {
    return Response.json({ error: "This source only accepts SHIFT ALLOCATIONS.xlsx." }, { status: 400 });
  }
  if (!context.request.body || (Number.isFinite(contentLength) && contentLength > MAX_CONTACT_LIST_BYTES)) {
    return Response.json({ error: "Contact-list workbook is missing or too large." }, { status: 413 });
  }

  try {
    const db = context.env.ROSTER_DB;
    await ensureCalendarSchema(db);
    if (providerVersion) {
      const matchingVersion = await db.prepare(`
        SELECT id FROM contact_list_files
        WHERE source_id = ? AND provider_version = ? AND LOWER(name) = LOWER(?)
        ORDER BY received_at DESC LIMIT 1
      `).bind(SOURCE_ID, providerVersion, fileName).first();
      if (matchingVersion?.id) {
        return Response.json({ ok: true, status: "unchanged", sourceId: SOURCE_ID, fileId: String(matchingVersion.id) });
      }
    }

    let size = 0;
    const measuredBody = context.request.body.pipeThrough(new TransformStream({
      transform(chunk, controller) {
        size += chunk.byteLength;
        if (size > MAX_CONTACT_LIST_BYTES) throw new Error("Contact-list workbook is too large.");
        controller.enqueue(chunk);
      },
    }));
    const nonce = crypto.randomUUID();
    const fileId = `contact:${SOURCE_ID}:${nonce}`;
    const objectKey = `contact-lists/${SOURCE_ID}/${nonce}`;
    const contentType = String(context.request.headers.get("content-type") || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    await context.env.ROSTER_FILES.put(objectKey, measuredBody, { httpMetadata: { contentType } });
    if (!size) {
      await context.env.ROSTER_FILES.delete(objectKey);
      return Response.json({ error: "Contact-list workbook is missing or too large." }, { status: 413 });
    }

    const now = new Date().toISOString();
    await db.prepare(`
      INSERT INTO contact_list_files (
        id, source_id, name, size, last_modified, object_key, content_type,
        content_hash, provider_version, provider_modified_at, received_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      fileId, SOURCE_ID, fileName, size, Date.parse(providerModifiedAt) || Date.now(), objectKey,
      contentType, `stream:${nonce}`, providerVersion, providerModifiedAt, now,
    ).run();
    return Response.json({ ok: true, status: "stored", sourceId: SOURCE_ID, fileId, receivedAt: now }, { status: 202 });
  } catch (error) {
    console.error("Contact-list binary ingestion failed", error);
    return Response.json({ error: "Contact-list workbook could not be stored." }, { status: 422 });
  }
}

function normaliseFileName(value) {
  return String(value || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function requestMetadata(request, params, headerName, parameterName) {
  return String(request.headers.get(headerName) || params.get(parameterName) || "").trim();
}

function hasValidAutomationToken(request, configuredToken) {
  const token = String(configuredToken || "");
  const provided = String(request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!token || !provided || token.length !== provided.length) return false;
  let mismatch = 0;
  for (let index = 0; index < token.length; index += 1) mismatch |= token.charCodeAt(index) ^ provided.charCodeAt(index);
  return mismatch === 0;
}
