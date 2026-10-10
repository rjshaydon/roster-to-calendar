import { buildAutomatedDerivedRosterPayload } from "../functions/_lib/automation-import.js";
import { buildVhhDerivedRosterPayload, VHH_ROSTER_SOURCE_ID } from "../functions/_lib/vhh-roster.js";
import { extractVhhRosterWorkbook } from "./vhh-roster-workbook.mjs";
import { planRosterImportBatches } from "../functions/_lib/roster-import-batches.js";
import { executeBoundedRosterImport } from "./roster-import-driver.mjs";

import { guardedFetch } from "../functions/_lib/outbound-network.js";
import {appendFile} from 'node:fs/promises';
let processorDeferred=false;
async function processorOutcome(value) {if(process.env.GITHUB_OUTPUT) await appendFile(process.env.GITHUB_OUTPUT,`result=${value}\n`);}

const baseUrl = String(process.env.ROSTER_AUTOMATION_BASE_URL || "https://roster-to-calendar.pages.dev").replace(/\/$/, "");
const token = String(process.env.ROSTER_AUTOMATION_TOKEN || "");
const maintenanceBudget = process.env.ROSTER_MAINTENANCE_SESSION ? JSON.parse(process.env.ROSTER_MAINTENANCE_SESSION) : null;
const sourceId = String(process.env.ROSTER_AUTOMATION_SOURCE_ID || "").trim();

if (!token) throw new Error("ROSTER_AUTOMATION_TOKEN is required.");
if (!sourceId) throw new Error("ROSTER_AUTOMATION_SOURCE_ID is required.");

const admission = await automationRequest("/api/automation/account-budget", { method: "POST", body: {} });
if (admission.deferred || admission.paused) {
  console.log("Account budget deferred; queued work retained.");
  await processorOutcome('deferred');
  process.exit(0);
}
let lastAccountAdmission = Date.now();
const pending = await automationRequest(`/api/automation/pending?limit=1&sourceId=${encodeURIComponent(sourceId)}`);
const runs = Array.isArray(pending.runs) ? pending.runs : [];
if (runs.some((run) => run.sourceId !== sourceId)) throw new Error("Roster queue returned work for a different source.");
console.log(`Found ${runs.length} queued roster file(s).`);
if ((!runs.length && !pending.publicationPending) || pending.maintenanceDeferred || (process.env.ROSTER_AUTOMATION_RESUME_ONLY === "true" && pending.boundedImportEnabled !== true)) {
  console.log("No admitted import work; progress retained.");
  await processorOutcome(pending.maintenanceDeferred || runs.length || pending.publicationPending?'deferred':'completed');
  process.exit(0);
}
const parserConfig = runs.length ? await automationRequest(`/api/automation/parser-config?sourceId=${encodeURIComponent(sourceId)}`) : {};
if (parserConfig.deferred || parserConfig.paused) {await processorOutcome('deferred');process.exit(0);}
const parserExtensions = parserConfig?.parserExtensions && typeof parserConfig.parserExtensions === "object" ? parserConfig.parserExtensions : {};
const failures = [];

for (const run of runs) {
  try {
    await processRun(run);
  } catch (error) {
    const message = `Failed to process ${run.fileName || run.id}: ${error?.message || error}`;
    console.error(message);
    failures.push(message);
    if (run.boundedImportStarted) {
      console.error("Inactive staged data and receipts retained for the next bounded retry.");
      continue;
    }
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

if (pending.boundedImportEnabled === true) {
  for (let step = 0; step < 24; step += 1) {
    const refresh = await automationRequest("/api/automation/facility-refresh", { method: "POST", body: { sourceId, maintenanceBudget } });
    if(refresh.deferred) processorDeferred=true;
    if (refresh.idle || refresh.deferred || refresh.completed) break;
    if(step===23) processorDeferred=true;
  }
}

if (failures.length) {
  throw new Error(`${failures.length} roster file${failures.length === 1 ? "" : "s"} failed during background processing.`);
}
if(pending.boundedImportEnabled===true) {
  const remaining=await automationRequest(`/api/automation/pending?limit=1&sourceId=${encodeURIComponent(sourceId)}`);
  if(remaining.publicationPending || remaining.runs?.length || remaining.deferred || remaining.paused) processorDeferred=true;
}
await processorOutcome(processorDeferred?'deferred':'completed');

async function processRun(run) {
  const response = await guardedFetch(process.env, `${baseUrl}/api/automation/raw?runId=${encodeURIComponent(run.id)}&sourceId=${encodeURIComponent(sourceId)}`, {
    headers: authorizationHeaders(),
  }, { label: "Roster processor download" });
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
  let finished;
  if (pending.boundedImportEnabled === true) {
    const plan = await planRosterImportBatches(payload);
    run.boundedImportStarted = true;
    const importStep = forceStaging => step => automationRequest("/api/automation/derived", {
      method: "POST", body: { ...step, forceStaging, maintenanceBudget, runId: run.id, sourceId: run.sourceId, file: payload.file },
    });
    finished = await executeBoundedRosterImport(plan, importStep(false));
    if (finished.mode === "complete") {
      run.boundedImportStarted = false;
      try {
        finished = await postDerived(run, payload, "complete", payload.doctors, payload.eventsByDoctor, payload.issuesByDoctor);
      } catch (error) {
        if (error.code !== "ROSTER_INCREMENTAL_BUDGET" || !error.boundedReplacementRequired) throw error;
        run.boundedImportStarted = true;
        finished = await executeBoundedRosterImport(plan, importStep(true));
      }
    }
    if (finished.deferred) {
      processorDeferred=true;
      console.log("Import write allowance exhausted; queued progress retained for automatic continuation.");
      return;
    }
  }
  finished ||= await postDerived(run, payload, "complete", payload.doctors, payload.eventsByDoctor, payload.issuesByDoctor);
  if (finished.deferred) {
    processorDeferred=true;
    console.log("Roster correction deferred by the shared maintenance budget; existing calendars retained.");
    return;
  }
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
  if (path !== "/api/automation/account-budget" && Date.now() - lastAccountAdmission > 5 * 60 * 1000) {
    const admission = await automationRequest("/api/automation/account-budget", { method: "POST", body: {} });
    lastAccountAdmission = Date.now();
    if (admission.deferred || admission.paused) return { ok: true, deferred: true, maintenanceDeferred: true };
  }
  const headers = { ...authorizationHeaders(), ...(options.body ? { "Content-Type": "application/json" } : {}) };
  let lastError = null;
  for (let attempt = 1; attempt <= 6; attempt += 1) {
    try {
      const response = await guardedFetch(process.env, `${baseUrl}${path}`, {
        method: options.method || "GET",
        headers,
        body: options.body ? JSON.stringify({ ...options.body, maintenanceBudget }) : undefined,
      }, { label: "Roster processor callback" });
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
      lastError.code = String(result.code || '');
      lastError.boundedReplacementRequired = result.boundedReplacementRequired === true;
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
