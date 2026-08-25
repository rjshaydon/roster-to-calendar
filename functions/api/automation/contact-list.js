import { ensureCalendarSchema, hasCalendarDb } from "../../_lib/d1-calendar.js";
import { sha256Hex } from "../../_lib/automation-import.js";
import { extractDdhClinicianContactsFromWorkbook, extractMmcDoctorContactsFromWorkbook } from "../../_lib/contact-list-workbook.js";
import { normaliseContactListExtract } from "../../../public/static/contact-allocations.js";

const SOURCES = new Map([
  ["mmc-shift-allocations", {
    expectedFileName: "shift allocations.xlsx",
    outputFileName: "SHIFT ALLOCATIONS doctors.json",
    extract: extractMmcDoctorContactsFromWorkbook,
  }],
  ["ddh-daily-contact-sheet", {
    expectedFileName: "daily contact sheet.xlsx",
    outputFileName: "Daily Contact Sheet clinicians.json",
    extract: extractDdhClinicianContactsFromWorkbook,
  }],
]);
const MAX_CONTACT_LIST_BYTES = 50 * 1024 * 1024;

// Accept the existing Power Automate full-workbook delivery, extract only the
// doctor rows from SHIFT ALLOCATIONS, then retain only the small JSON result.
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
    const source = SOURCES.get(upload.sourceId);
    if (!source) return Response.json({ error: "Unknown contact-list source." }, { status: 400 });
    if (normaliseFileName(upload.file.name) !== source.expectedFileName) {
      return Response.json({ error: "The uploaded filename does not match this contact-list source." }, { status: 400 });
    }
    if (!upload.file.size || upload.file.size > MAX_CONTACT_LIST_BYTES) {
      return Response.json({ error: "Contact-list workbook is missing or too large." }, { status: 413 });
    }

    const workbookBytes = new Uint8Array(await upload.file.arrayBuffer());
    const parsed = await source.extract(workbookBytes, {
      providerModifiedAt: upload.providerModifiedAt,
    });
    const extract = normaliseContactListExtract(parsed);
    if (!extract) return Response.json({ error: "Doctor contacts could not be extracted from the workbook." }, { status: 422 });
    const bytes = new TextEncoder().encode(JSON.stringify(extract));
    const contentHash = await sha256Hex(bytes);

    const db = context.env.ROSTER_DB;
    await ensureCalendarSchema(db);
    const existing = await db.prepare(`
      SELECT id, object_key, content_hash, provider_version, name
      FROM contact_list_files
      WHERE source_id = ?
      ORDER BY received_at DESC
    `).bind(upload.sourceId).all();
    const match = existing.results.find((entry) => String(entry.content_hash || "") === contentHash
      || (upload.providerVersion && String(entry.provider_version || "") === upload.providerVersion
        && String(entry.name || "").toLowerCase() === source.outputFileName.toLowerCase()));
    if (match?.id) {
      await deleteOtherVersions(context, existing.results, String(match.id));
      return Response.json({
        ok: true,
        status: "unchanged",
        sourceId: upload.sourceId,
        sourceDate: extract.sourceDate,
        contactCount: extract.contacts.filter((contact) => contact.isPopulated).length,
        fileId: String(match.id),
      });
    }

    const now = new Date().toISOString();
    const fileId = `contact:${upload.sourceId}:${contentHash.slice(0, 24)}`;
    const objectKey = `contact-lists/${upload.sourceId}/${contentHash}.json`;
    await context.env.ROSTER_FILES.put(objectKey, bytes, {
      httpMetadata: { contentType: "application/json; charset=utf-8" },
    });
    await db.prepare(`
      INSERT INTO contact_list_files (
        id, source_id, name, size, last_modified, object_key, content_type,
        content_hash, provider_version, provider_modified_at, received_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      fileId, upload.sourceId, source.outputFileName, bytes.byteLength,
      Date.parse(upload.providerModifiedAt) || upload.file.lastModified || Date.now(),
      objectKey, "application/json; charset=utf-8", contentHash,
      upload.providerVersion, upload.providerModifiedAt, now,
    ).run();
    await deleteOtherVersions(context, existing.results, "");
    return Response.json({
      ok: true,
      status: "stored",
      sourceId: upload.sourceId,
      sourceDate: extract.sourceDate,
      contactCount: extract.contacts.filter((contact) => contact.isPopulated).length,
      fileId,
      receivedAt: now,
    }, { status: 202 });
  } catch (error) {
    console.error("Contact-list ingestion failed", error);
    return Response.json({ error: "Contact-list workbook could not be processed." }, { status: 422 });
  }
}

async function deleteOtherVersions(context, entries, keepId) {
  for (const entry of entries) {
    if (keepId && String(entry.id) === keepId) continue;
    if (entry.object_key) await context.env.ROSTER_FILES.delete(String(entry.object_key));
    await context.env.ROSTER_DB.prepare("DELETE FROM contact_list_files WHERE id = ?").bind(String(entry.id)).run();
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
