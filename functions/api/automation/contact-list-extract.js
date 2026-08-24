import { ensureCalendarSchema, hasCalendarDb } from "../../_lib/d1-calendar.js";
import { sha256Hex } from "../../_lib/automation-import.js";

const SOURCE_ID = "mmc-shift-allocations";
const FILE_NAME = "SHIFT ALLOCATIONS doctors.json";
const MAX_CONTACTS = 160;
const MAX_BODY_BYTES = 512 * 1024;

// Receives the small, doctors-only result from the Excel Office Script. This
// intentionally has a separate endpoint from the legacy workbook uploads.
export async function onRequestPost(context) {
  if (!hasValidAutomationToken(context.request, context.env.ROSTER_AUTOMATION_TOKEN)) {
    return Response.json({ error: "Unauthorized." }, { status: 401 });
  }
  if (!hasCalendarDb(context.env) || !context.env.ROSTER_FILES?.put) {
    return Response.json({ error: "Contact-list storage is unavailable." }, { status: 503 });
  }

  try {
    const contentLength = Number(context.request.headers.get("content-length") || "0");
    if (Number.isFinite(contentLength) && contentLength > MAX_BODY_BYTES) {
      return Response.json({ error: "Contact-list extract is too large." }, { status: 413 });
    }
    const payload = await context.request.json();
    const extract = normaliseExtract(payload);
    if (!extract) return Response.json({ error: "Invalid doctor contact extract." }, { status: 400 });

    const bytes = new TextEncoder().encode(JSON.stringify(extract));
    if (bytes.byteLength > MAX_BODY_BYTES) {
      return Response.json({ error: "Contact-list extract is too large." }, { status: 413 });
    }

    const db = context.env.ROSTER_DB;
    await ensureCalendarSchema(db);
    const providerVersion = String(payload?.providerVersion || "").trim();
    if (providerVersion) {
      const prior = await db.prepare(`
        SELECT id FROM contact_list_files
        WHERE source_id = ? AND provider_version = ? AND LOWER(name) = LOWER(?)
        ORDER BY received_at DESC LIMIT 1
      `).bind(SOURCE_ID, providerVersion, FILE_NAME).first();
      if (prior?.id) {
        return Response.json({ ok: true, status: "unchanged", sourceId: SOURCE_ID, fileId: String(prior.id) });
      }
    }

    const contentHash = await sha256Hex(bytes);
    const existing = await db.prepare(`
      SELECT id, object_key, content_hash FROM contact_list_files
      WHERE source_id = ?
      ORDER BY received_at DESC
    `).bind(SOURCE_ID).all();
    const matchingHash = existing.results.find((entry) => String(entry.content_hash || "") === contentHash);
    if (matchingHash?.id) {
      for (const entry of existing.results) {
        if (String(entry.id) === String(matchingHash.id)) continue;
        await context.env.ROSTER_FILES.delete(String(entry.object_key));
        await db.prepare("DELETE FROM contact_list_files WHERE id = ?").bind(String(entry.id)).run();
      }
      return Response.json({ ok: true, status: "unchanged", sourceId: SOURCE_ID, fileId: String(matchingHash.id) });
    }

    const now = new Date().toISOString();
    const fileId = `contact:${SOURCE_ID}:${contentHash.slice(0, 24)}`;
    const objectKey = `contact-lists/${SOURCE_ID}/${contentHash}.json`;
    await context.env.ROSTER_FILES.put(objectKey, bytes, {
      httpMetadata: { contentType: "application/json; charset=utf-8" },
    });
    await db.prepare(`
      INSERT INTO contact_list_files (
        id, source_id, name, size, last_modified, object_key, content_type,
        content_hash, provider_version, provider_modified_at, received_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      fileId, SOURCE_ID, FILE_NAME, bytes.byteLength,
      Date.parse(extract.providerModifiedAt) || Date.now(), objectKey,
      "application/json; charset=utf-8", contentHash, providerVersion,
      extract.providerModifiedAt, now,
    ).run();

    for (const entry of existing.results) {
      await context.env.ROSTER_FILES.delete(String(entry.object_key));
      await db.prepare("DELETE FROM contact_list_files WHERE id = ?").bind(String(entry.id)).run();
    }
    return Response.json({ ok: true, status: "stored", sourceId: SOURCE_ID, fileId, receivedAt: now });
  } catch (error) {
    console.error("Contact-list extract ingestion failed", error);
    return Response.json({ error: "Contact-list extract could not be stored." }, { status: 422 });
  }
}

function normaliseExtract(payload) {
  if (String(payload?.sourceId || "").trim() !== SOURCE_ID || !Array.isArray(payload?.contacts)) return null;
  if (payload.contacts.length > MAX_CONTACTS) return null;
  const contacts = payload.contacts.map((entry) => ({
    area: String(entry?.area || "").trim(),
    shift: String(entry?.shift || "").trim(),
    role: String(entry?.role || "").trim(),
    name: String(entry?.name || "").trim(),
    phone: String(entry?.phone || "").trim(),
    isPopulated: Boolean(entry?.isPopulated),
  }));
  if (contacts.some((entry) => !["Adult Emergency", "Paediatric Emergency"].includes(entry.area)
    || !["AM", "PM", "Night"].includes(entry.shift)
    || !entry.role
    || /\bnic\b|nurs|(^|\W)(rn|en)(\W|$)/i.test(entry.role))) return null;
  return {
    sourceId: SOURCE_ID,
    fileName: FILE_NAME,
    sourceDate: String(payload?.sourceDate || "").trim(),
    providerModifiedAt: String(payload?.providerModifiedAt || "").trim(),
    contacts,
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
