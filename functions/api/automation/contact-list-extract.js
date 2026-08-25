import { ensureCalendarSchema, hasCalendarDb } from "../../_lib/d1-calendar.js";
import { sha256Hex } from "../../_lib/automation-import.js";
import { normaliseContactListExtract } from "../../../public/static/contact-allocations.js";

const MAX_BODY_BYTES = 512 * 1024;

// Receives the small, doctors-only result from the Excel Office Script. This
// intentionally has a separate endpoint from the legacy workbook uploads.
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

  try {
    const contentLength = Number(context.request.headers.get("content-length") || "0");
    if (Number.isFinite(contentLength) && contentLength > MAX_BODY_BYTES) {
      return Response.json({ error: "Contact-list extract is too large." }, { status: 413 });
    }
    const payload = await context.request.json();
    const extract = normaliseContactListExtract(payload);
    if (!extract) return Response.json({ error: "Invalid doctor contact extract." }, { status: 400 });
    const sourceId = extract.sourceId;
    const fileName = extract.fileName;

    const bytes = new TextEncoder().encode(JSON.stringify(extract));
    if (bytes.byteLength > MAX_BODY_BYTES) {
      return Response.json({ error: "Contact-list extract is too large." }, { status: 413 });
    }

    const db = context.env.ROSTER_DB;
    await ensureCalendarSchema(db);
    const providerVersion = String(payload?.providerVersion || "").trim();
    const contentHash = await sha256Hex(bytes);
    const existing = await db.prepare(`
      SELECT id, object_key, content_hash FROM contact_list_files
      WHERE source_id = ?
      ORDER BY received_at DESC
    `).bind(sourceId).all();
    const matchingHash = existing.results.find((entry) => String(entry.content_hash || "") === contentHash);
    if (matchingHash?.id) {
      for (const entry of existing.results) {
        if (String(entry.id) === String(matchingHash.id)) continue;
        await context.env.ROSTER_FILES.delete(String(entry.object_key));
        await db.prepare("DELETE FROM contact_list_files WHERE id = ?").bind(String(entry.id)).run();
      }
      return Response.json({
        ok: true,
        status: "unchanged",
        sourceId,
        sourceDate: extract.sourceDate,
        contactCount: extract.contacts.filter((contact) => contact.isPopulated).length,
        fileId: String(matchingHash.id),
      });
    }

    const now = new Date().toISOString();
    const fileId = `contact:${sourceId}:${contentHash.slice(0, 24)}`;
    const objectKey = `contact-lists/${sourceId}/${contentHash}.json`;
    await context.env.ROSTER_FILES.put(objectKey, bytes, {
      httpMetadata: { contentType: "application/json; charset=utf-8" },
    });
    await db.prepare(`
      INSERT INTO contact_list_files (
        id, source_id, name, size, last_modified, object_key, content_type,
        content_hash, provider_version, provider_modified_at, received_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      fileId, sourceId, fileName, bytes.byteLength,
      Date.parse(extract.providerModifiedAt) || Date.now(), objectKey,
      "application/json; charset=utf-8", contentHash, providerVersion,
      extract.providerModifiedAt, now,
    ).run();

    for (const entry of existing.results) {
      await context.env.ROSTER_FILES.delete(String(entry.object_key));
      await db.prepare("DELETE FROM contact_list_files WHERE id = ?").bind(String(entry.id)).run();
    }
    return Response.json({
      ok: true,
      status: "stored",
      sourceId,
      sourceDate: extract.sourceDate,
      contactCount: extract.contacts.filter((contact) => contact.isPopulated).length,
      fileId,
      receivedAt: now,
    });
  } catch (error) {
    console.error("Contact-list extract ingestion failed", error);
    return Response.json({ error: "Contact-list extract could not be stored." }, { status: 422 });
  }
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
