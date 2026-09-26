import { ingestContactExtract } from "./contact-list-extract.js";
import {
  extractDdhClinicianContactsFromWorkbook,
  extractMmcDoctorContactsFromWorkbook,
} from "../../_lib/contact-list-workbook.js";
import { contactAutomationPausedResponse, contactAutomationSourceEnabled } from "../../_lib/contact-automation-guard.js";

const XLSX_CONTENT_TYPES = new Set([
  "application/octet-stream",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
]);
const POWER_AUTOMATE_CONTENT_TYPES = new Set([
  "application/json",
  "text/plain",
]);
const SOURCE_RULES = new Map([
  ["ddh-daily-contact-sheet", {
    fileName: "Daily Contact Sheet.xlsx",
    maximumBytes: 5 * 1024 * 1024,
    extract: extractDdhClinicianContactsFromWorkbook,
  }],
  ["mmc-shift-allocations", {
    fileName: "SHIFT ALLOCATIONS.xlsx",
    maximumBytes: 5 * 1024 * 1024,
    extract: extractMmcDoctorContactsFromWorkbook,
  }],
]);

export async function onRequestPost(context) {
  if (!hasValidAutomationToken(
    context.request,
    context.env.ROSTER_AUTOMATION_TOKEN,
    context.env.DDH_CONTACT_AUTOMATION_TOKEN,
  )) {
    return Response.json({ error: "Unauthorized." }, { status: 401 });
  }

  const sourceId = header(context.request, "x-contact-source-id");
  const rule = SOURCE_RULES.get(sourceId);
  if (!rule) return Response.json({ error: "Unknown contact workbook source." }, { status: 400 });
  const fileName = header(context.request, "x-contact-file-name");
  if (fileName !== rule.fileName) return Response.json({ error: "Unexpected contact workbook filename." }, { status: 400 });
  const contentType = String(context.request.headers.get("content-type") || "").split(";", 1)[0].trim().toLowerCase();
  if (!XLSX_CONTENT_TYPES.has(contentType) && !POWER_AUTOMATE_CONTENT_TYPES.has(contentType)) {
    return Response.json({ error: "Expected an Excel workbook." }, { status: 415 });
  }
  if (!contactAutomationSourceEnabled(context.env, sourceId)) return contactAutomationPausedResponse();

  try {
    const workbookBytes = await readWorkbookBody(context.request, contentType, rule.maximumBytes);
    const providerModifiedAt = header(context.request, "x-provider-modified-at");
    const providerVersion = header(context.request, "x-provider-version");
    const extract = await rule.extract(workbookBytes, { providerModifiedAt });
    return ingestContactExtract(context, {
      ...extract,
      fileName,
      providerModifiedAt,
      providerVersion,
    });
  } catch (error) {
    if (error?.code === "contact-workbook-too-large") {
      return Response.json({ error: "Contact workbook is too large." }, { status: 413 });
    }
    console.error("Contact workbook extraction failed", error);
    return Response.json({ error: "Contact workbook could not be read." }, { status: 422 });
  }
}

async function readWorkbookBody(request, contentType, maximumBytes) {
  // Power Automate serializes connector file content as either the standard
  // { "$content-type", "$content" } JSON envelope or, in some tenants, the
  // base64 value alone. It can retain the explicitly configured XLSX HTTP
  // content type even when the body is one of those encoded representations,
  // so inspect the bounded body rather than trusting the MIME type alone.
  // Arbitrary JSON or text never reaches the workbook parser.
  const encodedLimit = Math.ceil(maximumBytes * 4 / 3) + 4096;
  const encodedBytes = await readBodyWithLimit(request, encodedLimit);
  if (hasZipSignature(encodedBytes)) {
    if (encodedBytes.byteLength > maximumBytes) throw tooLarge();
    return encodedBytes;
  }
  const text = new TextDecoder().decode(encodedBytes).trim();
  let base64 = text;
  if (contentType === "application/json" || text.startsWith("{") || text.startsWith('"')) {
    let envelope;
    try {
      envelope = JSON.parse(text);
    } catch {
      throw invalidWorkbook();
    }
    if (typeof envelope === "string") base64 = envelope;
    else {
      const envelopeType = String(envelope?.["$content-type"] || "").split(";", 1)[0].trim().toLowerCase();
      if (!envelope || typeof envelope !== "object" || Array.isArray(envelope)
        || typeof envelope.$content !== "string"
        || (envelopeType && !XLSX_CONTENT_TYPES.has(envelopeType))) throw invalidWorkbook();
      base64 = envelope.$content;
    }
  }
  const bytes = decodeBase64Workbook(base64, maximumBytes);
  if (!hasZipSignature(bytes)) throw invalidWorkbook();
  return bytes;
}

async function readBodyWithLimit(request, maximumBytes) {
  const declared = Number(request.headers.get("content-length") || "0");
  if (Number.isFinite(declared) && declared > maximumBytes) throw tooLarge();
  if (!request.body?.getReader) {
    const bytes = new Uint8Array(await request.arrayBuffer());
    if (bytes.byteLength > maximumBytes) throw tooLarge();
    return bytes;
  }
  const reader = request.body.getReader();
  const chunks = [];
  let total = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > maximumBytes) {
      await reader.cancel();
      throw tooLarge();
    }
    chunks.push(value);
  }
  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes;
}

function tooLarge() {
  const error = new Error("Contact workbook is too large.");
  error.code = "contact-workbook-too-large";
  return error;
}

function invalidWorkbook() {
  const error = new Error("Expected an Excel workbook.");
  error.code = "contact-workbook-invalid";
  return error;
}

function decodeBase64Workbook(value, maximumBytes) {
  const compact = String(value || "").replace(/\s+/g, "");
  if (!compact || compact.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(compact)) throw invalidWorkbook();
  const padding = compact.endsWith("==") ? 2 : compact.endsWith("=") ? 1 : 0;
  const decodedLength = (compact.length / 4) * 3 - padding;
  if (decodedLength > maximumBytes) throw tooLarge();
  let binary;
  try {
    binary = atob(compact);
  } catch {
    throw invalidWorkbook();
  }
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

function hasZipSignature(bytes) {
  return bytes?.byteLength >= 4
    && bytes[0] === 0x50
    && bytes[1] === 0x4b
    && ((bytes[2] === 0x03 && bytes[3] === 0x04)
      || (bytes[2] === 0x05 && bytes[3] === 0x06)
      || (bytes[2] === 0x07 && bytes[3] === 0x08));
}

function header(request, name) {
  return String(request.headers.get(name) || "").trim();
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
