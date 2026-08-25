import { ensureCalendarSchema, hasCalendarDb } from "../../_lib/d1-calendar.js";
import { sha256Hex } from "../../_lib/automation-import.js";
import { extractMmcDoctorContactsFromWorkbook } from "../../_lib/contact-list-workbook.js";
import { normaliseContactListExtract } from "../../../public/static/contact-allocations.js";

const SOURCE_ID = "mmc-shift-allocations";
const EXPECTED_FILE_NAME = "shift allocations.xlsx";
// Power Automate can expand this 28 MB workbook substantially on the wire.
// Keep safely below the 100 MB ingress ceiling while allowing that connector.
const MAX_CONTACT_LIST_BYTES = 80 * 1024 * 1024;

// SHIFT ALLOCATIONS.xlsx is approximately 28 MB. Power Automate transports it
// as Base64, which is decoded incrementally before it is retained in R2.
export async function onRequestPost(context) {
  if (!hasValidAutomationToken(
    context.request,
    context.env.ROSTER_AUTOMATION_TOKEN,
    context.env.DDH_CONTACT_AUTOMATION_TOKEN,
  )) {
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
  if (!context.request.body) {
    return Response.json({ error: "Contact-list workbook is missing." }, { status: 413 });
  }
  if (Number.isFinite(contentLength) && contentLength > MAX_CONTACT_LIST_BYTES) {
    return Response.json({
      error: "Contact-list workbook is too large.",
      contentLength,
      maxBytes: MAX_CONTACT_LIST_BYTES,
    }, { status: 413 });
  }
  if (!Number.isFinite(contentLength) || contentLength <= 0) {
    return Response.json({ error: "Contact-list workbook length is missing." }, { status: 411 });
  }
  const contentEncoding = String(context.request.headers.get("x-roster-content-encoding") || "").toLowerCase();
  const contentType = String(context.request.headers.get("content-type") || "").toLowerCase();
  if (contentEncoding === "base64" || contentType.startsWith("text/plain")) {
    return storeBase64Workbook(context, { sourceId, fileName, providerVersion, providerModifiedAt });
  }

  let stage = "schema";
  try {
    const db = context.env.ROSTER_DB;
    await ensureCalendarSchema(db);
    stage = "duplicate-check";
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
    stage = "r2-put";
    // R2 requires the exact stream length. TransformStream removes that
    // information, so retain Power Automate's Content-Length explicitly.
    const fixedLength = new FixedLengthStream(contentLength);
    await Promise.all([
      context.env.ROSTER_FILES.put(objectKey, fixedLength.readable, { httpMetadata: { contentType } }),
      measuredBody.pipeTo(fixedLength.writable),
    ]);
    if (!size) {
      await context.env.ROSTER_FILES.delete(objectKey);
      return Response.json({ error: "Contact-list workbook is missing or too large." }, { status: 413 });
    }

    stage = "d1-insert";
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
    return Response.json({
      error: "Contact-list workbook could not be stored.",
      stage,
      diagnostic: String(error?.message || error).slice(0, 500),
    }, { status: 422 });
  }
}

async function storeBase64Workbook(context, metadata) {
  try {
    const workbook = await decodeBase64Workbook(context.request.body);
    if (!workbook.size || workbook.size > MAX_CONTACT_LIST_BYTES) {
      return Response.json({ error: "Contact-list workbook is missing or too large." }, { status: 413 });
    }

    const parsed = await extractMmcDoctorContactsFromWorkbook(new Uint8Array(await workbook.arrayBuffer()), {
      providerModifiedAt: metadata.providerModifiedAt,
    });
    const extract = normaliseContactListExtract(parsed);
    if (!extract) throw new Error("Doctor contacts could not be extracted from the workbook.");
    const bytes = new TextEncoder().encode(JSON.stringify(extract));
    const contentHash = await sha256Hex(bytes);

    const db = context.env.ROSTER_DB;
    await ensureCalendarSchema(db);
    const existing = await db.prepare(`
      SELECT id, object_key, content_hash
      FROM contact_list_files
      WHERE source_id = ?
      ORDER BY received_at DESC
    `).bind(SOURCE_ID).all();
    const matchingExtract = existing.results.find((entry) => String(entry.content_hash || "") === contentHash);
    if (matchingExtract?.id) {
      await deleteOtherVersions(context, existing.results, String(matchingExtract.id));
      return Response.json({
        ok: true,
        status: "unchanged",
        sourceId: SOURCE_ID,
        sourceDate: extract.sourceDate,
        contactCount: extract.contacts.filter((contact) => contact.isPopulated).length,
        fileId: String(matchingExtract.id),
      });
    }

    const fileId = `contact:${SOURCE_ID}:${contentHash.slice(0, 24)}`;
    const objectKey = `contact-lists/${SOURCE_ID}/${contentHash}.json`;
    const now = new Date().toISOString();
    const contentType = "application/json; charset=utf-8";
    await context.env.ROSTER_FILES.put(objectKey, bytes, { httpMetadata: { contentType } });
    await db.prepare(`
      INSERT INTO contact_list_files (
        id, source_id, name, size, last_modified, object_key, content_type,
        content_hash, provider_version, provider_modified_at, received_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      fileId, SOURCE_ID, "SHIFT ALLOCATIONS doctors.json", bytes.byteLength,
      Date.parse(metadata.providerModifiedAt) || Date.now(), objectKey, contentType,
      contentHash, metadata.providerVersion, metadata.providerModifiedAt, now,
    ).run();
    await deleteOtherVersions(context, existing.results, "");
    return Response.json({
      ok: true,
      status: "stored",
      sourceId: SOURCE_ID,
      sourceDate: extract.sourceDate,
      contactCount: extract.contacts.filter((contact) => contact.isPopulated).length,
      fileId,
      receivedAt: now,
    }, { status: 202 });
  } catch (error) {
    console.error("Contact-list Base64 ingestion failed", error);
    return Response.json({ error: "Contact-list workbook could not be stored." }, { status: 422 });
  }
}

async function deleteOtherVersions(context, entries, keepId) {
  for (const entry of entries) {
    if (keepId && String(entry.id) === keepId) continue;
    if (entry.object_key) await context.env.ROSTER_FILES.delete(String(entry.object_key));
    await context.env.ROSTER_DB.prepare("DELETE FROM contact_list_files WHERE id = ?").bind(String(entry.id)).run();
  }
}

async function decodeBase64Workbook(body) {
  const reader = body.getReader();
  const decoder = new TextDecoder();
  const chunks = [];
  let pending = "";
  let size = 0;

  const decode = (encoded) => {
    if (!encoded) return;
    if (!/^[A-Za-z0-9+/]*={0,2}$/.test(encoded)) {
      throw new Error("Contact-list workbook is not valid Base64.");
    }
    const binary = atob(encoded);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    size += bytes.byteLength;
    if (size > MAX_CONTACT_LIST_BYTES) throw new Error("Contact-list workbook is too large.");
    chunks.push(bytes);
  };

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    pending += decoder.decode(value, { stream: true });
    const readyLength = pending.length - (pending.length % 4);
    decode(pending.slice(0, readyLength));
    pending = pending.slice(readyLength);
  }
  pending += decoder.decode();
  if (pending.length % 4 !== 0) throw new Error("Contact-list workbook is not valid Base64.");
  decode(pending);
  return new Blob(chunks, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function normaliseFileName(value) {
  return String(value || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function requestMetadata(request, params, headerName, parameterName) {
  return String(request.headers.get(headerName) || params.get(parameterName) || "").trim();
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
