import { buildAutomatedDerivedRosterPayload } from "../functions/_lib/automation-import.js";
import { buildVhhDerivedRosterPayload, VHH_ROSTER_SOURCE_ID } from "../functions/_lib/vhh-roster.js";
import { extractVhhRosterWorkbook } from "./vhh-roster-workbook.mjs";

const baseUrl = String(process.env.ROSTER_AUTOMATION_BASE_URL || "https://roster-to-calendar.pages.dev").replace(/\/$/, "");
const token = String(process.env.ROSTER_AUTOMATION_TOKEN || "");
const doctorChunkSize = 18;

if (!token) throw new Error("ROSTER_AUTOMATION_TOKEN is required.");

const pending = await automationRequest("/api/automation/pending?limit=4");
const runs = Array.isArray(pending.runs) ? pending.runs : [];
const parserConfig = await automationRequest("/api/automation/parser-config");
const parserExtensions = parserConfig?.parserExtensions && typeof parserConfig.parserExtensions === "object"
  ? parserConfig.parserExtensions
  : {};
console.log(`Found ${runs.length} queued roster file(s).`);
const failures = [];

for (const run of runs) {
  try {
    await processRun(run);
  } catch (error) {
    const message = `Failed to process ${run.fileName || run.id}: ${error?.message || error}`;
    console.error(message);
    failures.push(message);
    try {
      await automationRequest("/api/automation/derived", {
        method: "POST",
        body: {
          runId: run.id,
          sourceId: run.sourceId,
          phase: "failed",
          file: { id: run.fileId, name: run.fileName, sourceId: run.sourceId, sourceType: run.sourceType },
          message: String(error?.message || "Background processor could not parse or save this roster.").slice(0, 300),
        },
      });
    } catch (reportError) {
      const reportingMessage = `Could not mark ${run.fileName || run.id} failed: ${reportError?.message || reportError}`;
      console.error(reportingMessage);
      failures.push(reportingMessage);
    }
  }
}

if (failures.length) {
  throw new Error(`${failures.length} roster file${failures.length === 1 ? "" : "s"} failed during background processing.`);
}

async function processRun(run) {
  const response = await fetch(`${baseUrl}/api/automation/raw?runId=${encodeURIComponent(run.id)}`, {
    headers: authorizationHeaders(),
  });
  if (!response.ok) throw new Error(`Roster download returned HTTP ${response.status}.`);
  let payload;
  let processedFileName = run.fileName || "roster.xlsx";
  if (run.sourceId === VHH_ROSTER_SOURCE_ID) {
    const isLegacyJson = /json/i.test(run.contentType || "") || /\.json$/i.test(processedFileName);
    const bytes = new Uint8Array(await response.arrayBuffer());
    const file = new File([bytes], processedFileName, {
      type: run.contentType || (isLegacyJson ? "application/json" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
      lastModified: Number(run.lastModified || Date.now()),
    });
    console.log(`Parsing ${file.name} (${file.size} bytes).`);
    const extract = isLegacyJson
      ? JSON.parse(new TextDecoder().decode(bytes))
      : await extractVhhRosterWorkbook(file, {
          providerModifiedAt: new Date(file.lastModified).toISOString(),
          providerVersion: run.providerVersion,
        });
    payload = buildVhhDerivedRosterPayload({
      extract,
      contentHash: run.contentHash,
      fileId: run.fileId,
      providerVersion: run.providerVersion,
      fileSize: file.size,
      lastModified: file.lastModified,
    });
  } else {
    const bytes = new Uint8Array(await response.arrayBuffer());
    const file = new File([bytes], run.fileName || "roster.xlsx", {
      type: run.contentType || "application/octet-stream",
      lastModified: Number(run.lastModified || Date.now()),
    });
    console.log(`Parsing ${file.name} (${file.size} bytes).`);
    payload = await buildAutomatedDerivedRosterPayload({
      file,
      sourceId: run.sourceId,
      contentHash: run.contentHash,
      fileId: run.fileId,
      providerVersion: run.providerVersion,
      parserExtensions,
    });
  }
  console.log(`Parsed ${payload.file.name}: ${payload.doctors.length} doctors, ${payload.eventCount} calendar events.`);
  payload.file = {
    ...payload.file,
    lastModified: Number(run.lastModified || payload.file.lastModified || Date.now()),
  };
  await postDerived(run, payload, "start", payload.doctors, {}, {});
  console.log("Created the derived roster record.");
  const doctorKeys = payload.doctors.map((doctor) => doctor.key).filter(Boolean);
  for (let index = 0; index < doctorKeys.length; index += doctorChunkSize) {
    const keys = doctorKeys.slice(index, index + doctorChunkSize);
    await postDerived(
      run,
      payload,
      "events",
      payload.doctors.filter((doctor) => keys.includes(doctor.key)),
      Object.fromEntries(keys.map((key) => [key, payload.eventsByDoctor[key] || []])),
      Object.fromEntries(keys.map((key) => [key, payload.issuesByDoctor[key] || []])),
    );
    console.log(`Saved calendar event batch ${Math.floor(index / doctorChunkSize) + 1} of ${Math.ceil(doctorKeys.length / doctorChunkSize)}.`);
  }
  const finished = await postDerived(run, payload, "finish", payload.doctors, {}, {});
  console.log(`Indexed ${processedFileName}: ${finished.doctorCount} doctors, ${finished.eventCount} shifts.`);
}

async function postDerived(run, payload, phase, doctors, eventsByDoctor, issuesByDoctor) {
  return automationRequest("/api/automation/derived", {
    method: "POST",
    body: {
      runId: run.id,
      sourceId: run.sourceId,
      phase,
      file: payload.file,
      doctors,
      eventsByDoctor,
      issuesByDoctor,
    },
  });
}

async function automationRequest(path, options = {}) {
  const headers = { ...authorizationHeaders(), ...(options.body ? { "Content-Type": "application/json" } : {}) };
  let lastError = null;
  for (let attempt = 1; attempt <= 6; attempt += 1) {
    try {
      const response = await fetch(`${baseUrl}${path}`, {
        method: options.method || "GET",
        headers,
        body: options.body ? JSON.stringify(options.body) : undefined,
      });
      const text = await response.text();
      let result = {};
      try {
        result = text ? JSON.parse(text) : {};
      } catch {
        // A Pages/Cloudflare error page is transient infrastructure output,
        // not a roster parser response. Treat it like a retryable 5xx rather
        // than failing the retained-file reparse immediately.
        lastError = new Error(`Automation endpoint returned non-JSON HTTP ${response.status || 502}.`);
        if (attempt < 6) {
          await new Promise((resolve) => setTimeout(resolve, attempt * 2000));
          continue;
        }
        break;
      }
      if (response.ok) return result;
      const diagnostic = String(result.code || result.phase || "").trim();
      lastError = new Error(`${result.error || `HTTP ${response.status}`}${diagnostic ? ` (${diagnostic})` : ""}`);
      if (![408, 429, 500, 502, 503, 504].includes(response.status)) break;
    } catch (error) {
      lastError = error;
    }
    await new Promise((resolve) => setTimeout(resolve, attempt * 2000));
  }
  throw lastError || new Error("Automation request failed.");
}

function authorizationHeaders() {
  return { Authorization: `Bearer ${token}` };
}
