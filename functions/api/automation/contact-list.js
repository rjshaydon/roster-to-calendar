import { ensureCalendarSchema, hasCalendarDb } from "../../_lib/d1-calendar.js";
import { sha256Hex } from "../../_lib/automation-import.js";

const SOURCE_ID = "mmc-shift-allocations";
const EXPECTED_FILE_NAME = "shift allocations.xlsx";
// SHIFT ALLOCATIONS.xlsx is currently about 28 MB. Power Automate sends its
// contents as Base64, so leave practical room for future revisions while
// retaining a bounded upload size at the ingress.
const MAX_CONTACT_LIST_BYTES = 50 * 1024 * 1024;

// This endpoint is intentionally separate from roster ingestion.  It accepts
// the approved Power Automate delivery of the MMC shift-allocation workbook,
// but does not parse or expose its contact data until a dedicated parser and
// access-controlled consumer are added.
export async function onRequestPost(context) {
  if (!hasValidAutomationToken(context.request, context.env.ROSTER_AUTOMATION_TOKEN)) {
    return Response.json({ error: "Unauthorized." }, { status: 401 });
  }
  if (!hasCalendarDb(context.env) || !context.env.ROSTER_FILES?.put) {
    return Response.json({ error: "Contact-list storage is unavailable." }, { status: 503 });
  }

  try {
    const upload = await readUpload(context.request);
    if (!(upload.file instanceof File)) return Response.json({ error: "A contact-list workbook is required." }, { status: 400 });
    if (upload.sourceId !== SOURCE_ID) return Response.json({ error: "Unknown contact-list source." }, { status: 400 });
    if (normaliseFileName(upload.file.name) !== EXPECTED_FILE_NAME) {
      return Response.json({ error: "This source only accepts SHIFT ALLOCATIONS.xlsx." }, { status: 400 });
    }
    if (!upload.file.size || upload.file.size > MAX_CONTACT_LIST_BYTES) {
      return Response.json({ error: "Contact-list workbook is missing or too large." }, { status: 413 });
    }

    const db = context.env.ROSTER_DB;
    await ensureCalendarSchema(db);
    const now = new Date().toISOString();
    const matchingVersion = upload.providerVersion
      ? await db.prepare(`
        SELECT id FROM contact_list_files
        WHERE source_id = ? AND provider_version = ? AND LOWER(name) = LOWER(?)
        ORDER BY received_at DESC LIMIT 1
      `).bind(SOURCE_ID, upload.providerVersion, upload.file.name).first()
      : null;
    if (matchingVersion?.id) {
      return Response.json({ ok: true, status: "unchanged", sourceId: SOURCE_ID, fileId: String(matchingVersion.id) });
    }

    const bytes = new Uint8Array(await upload.file.arrayBuffer());
    const contentHash = await sha256Hex(bytes);
    const matchingHash = await db.prepare(`
      SELECT id FROM contact_list_files
      WHERE source_id = ? AND content_hash = ? AND LOWER(name) = LOWER(?)
      ORDER BY received_at DESC LIMIT 1
    `).bind(SOURCE_ID, contentHash, upload.file.name).first();
    if (matchingHash?.id) {
      return Response.json({ ok: true, status: "unchanged", sourceId: SOURCE_ID, fileId: String(matchingHash.id) });
    }

    const fileId = `contact:${SOURCE_ID}:${contentHash.slice(0, 24)}`;
    const objectKey = `contact-lists/${SOURCE_ID}/${contentHash}`;
    const contentType = upload.file.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    await context.env.ROSTER_FILES.put(objectKey, bytes, { httpMetadata: { contentType } });
    await db.prepare(`
      INSERT INTO contact_list_files (
        id, source_id, name, size, last_modified, object_key, content_type,
        content_hash, provider_version, provider_modified_at, received_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      fileId, SOURCE_ID, upload.file.name, upload.file.size, upload.file.lastModified,
      objectKey, contentType, contentHash, upload.providerVersion, upload.providerModifiedAt, now,
    ).run();
    return Response.json({ ok: true, status: "stored", sourceId: SOURCE_ID, fileId, receivedAt: now }, { status: 202 });
  } catch (error) {
    console.error("Contact-list ingestion failed", error);
    return Response.json({ error: "Contact-list workbook could not be stored." }, { status: 422 });
  }
}

async function readUpload(request) {
  const contentType = String(request.headers.get("content-type") || "").toLowerCase();
  if (contentType.includes("application/json")) {
    const body = await request.json();
    const base64 = String(body?.contentBase64 || "").replace(/^data:[^;,]+;base64,/i, "");
    if (!base64) return { sourceId: "", providerVersion: "", providerModifiedAt: "", file: null };
    const binary = atob(base64);
    const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
    if (bytes.byteLength > MAX_CONTACT_LIST_BYTES) return { sourceId: "", providerVersion: "", providerModifiedAt: "", file: null };
    const providerModifiedAt = String(body?.providerModifiedAt || body?.lastModified || "").trim();
    const explicitLastModified = Number(body?.lastModified);
    const parsedLastModified = Date.parse(providerModifiedAt);
    return {
      sourceId: String(body?.sourceId || "").trim(),
      providerVersion: String(body?.providerVersion || "").trim(),
      providerModifiedAt,
      file: new File([bytes], String(body?.fileName || "SHIFT ALLOCATIONS.xlsx"), {
        type: String(body?.contentType || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        lastModified: Number.isFinite(explicitLastModified) && explicitLastModified > 0
          ? explicitLastModified
          : Number.isFinite(parsedLastModified) ? parsedLastModified : Date.now(),
      }),
    };
  }
  const formData = await request.formData();
  return {
    sourceId: String(formData.get("sourceId") || "").trim(),
    providerVersion: String(formData.get("providerVersion") || "").trim(),
    providerModifiedAt: String(formData.get("providerModifiedAt") || formData.get("lastModified") || "").trim(),
    file: formData.get("contactListFile"),
  };
}

function normaliseFileName(value) {
  return String(value || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function hasValidAutomationToken(request, configuredToken) {
  const token = String(configuredToken || "");
  const provided = String(request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!token || !provided || token.length !== provided.length) return false;
  let mismatch = 0;
  for (let index = 0; index < token.length; index += 1) mismatch |= token.charCodeAt(index) ^ provided.charCodeAt(index);
  return mismatch === 0;
}
