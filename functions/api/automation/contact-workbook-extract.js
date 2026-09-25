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
  if (!XLSX_CONTENT_TYPES.has(contentType)) return Response.json({ error: "Expected an Excel workbook." }, { status: 415 });
  if (!contactAutomationSourceEnabled(context.env, sourceId)) return contactAutomationPausedResponse();

  try {
    const workbookBytes = await readBodyWithLimit(context.request, rule.maximumBytes);
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
