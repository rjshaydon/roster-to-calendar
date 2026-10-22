import {createMaintenanceAdmission,isReadOnlyPatchPage,retryablePatchReadFailure,retryPatchRead} from './automation-checkpoint-client.mjs';
import {preserveVhhHistory} from './process-vhh-history.mjs';
import {drainRosterPublication} from './roster-publication-driver.mjs';
import {publishDoctorSource} from './process-doctor-source.mjs';
import {publishDoctorMonth,doctorCalendarRequest} from './process-doctor-calendars.mjs';
import { buildAutomatedDerivedRosterPayload } from "../functions/_lib/automation-import.js";
import { buildVhhDerivedRosterPayload, VHH_ROSTER_SOURCE_ID } from "../functions/_lib/vhh-roster.js";
import { extractVhhRosterWorkbook } from "./vhh-roster-workbook.mjs";
import { planRosterImportBatches } from "../functions/_lib/roster-import-batches.js";
import { executeRosterPatch } from "./roster-patch-plan.mjs";
import { executeRosterRebase } from './roster-staging-reuse.mjs';
import { attachRosterCompactPlan } from "./roster-compact-plan.mjs";
import { executeBoundedRosterImport } from "./roster-import-driver.mjs";

import { guardedFetch } from "../functions/_lib/outbound-network.js";
import {appendFile} from 'node:fs/promises';
let processorDeferred=false;
let processorReason='';
async function processorOutcome(value) {if(process.env.GITHUB_OUTPUT) await appendFile(process.env.GITHUB_OUTPUT,`result=${value}\nreason=${processorReason|| (value==='deferred'?'account-budget':'')}\n`);}

const baseUrl = String(process.env.ROSTER_AUTOMATION_BASE_URL || "https://roster-to-calendar.pages.dev").replace(/\/$/, "");
const token = String(process.env.ROSTER_AUTOMATION_TOKEN || "");
const maintenanceBudget = process.env.ROSTER_MAINTENANCE_SESSION ? JSON.parse(process.env.ROSTER_MAINTENANCE_SESSION) : null;
const sourceId = String(process.env.ROSTER_AUTOMATION_SOURCE_ID || "").trim();
const publicationOnly=process.env.ROSTER_PUBLICATION_ONLY==='true';
const recoverVhhHistory=process.env.ROSTER_RECOVER_VHH_HISTORY==='true';

if (!token) throw new Error("ROSTER_AUTOMATION_TOKEN is required.");
if (!sourceId) throw new Error("ROSTER_AUTOMATION_SOURCE_ID is required.");

const accountAdmission=createMaintenanceAdmission({request:body=>automationRequest("/api/automation/account-budget",{method:"POST",body})});
const admission = await accountAdmission();
if (admission.deferred || admission.paused) {
  console.log("Account budget deferred; queued work retained.");
  await processorOutcome('deferred');
  process.exit(0);
}
const pending = await automationRequest(`/api/automation/pending?limit=1&sourceId=${encodeURIComponent(sourceId)}`);
if(pending.maintenanceDeferred&&!publicationOnly){console.log('Insufficient usable maintenance headroom; retained work left untouched before parsing.');await processorOutcome('deferred');process.exit(0);}
let runs = Array.isArray(pending.runs) ? pending.runs : [];
if(recoverVhhHistory) {
  if(sourceId!==VHH_ROSTER_SOURCE_ID) throw Error('Historical recovery is restricted to VHH.');
  const retained=await automationRequest('/api/automation/derived',{method:'POST',body:{phase:'bounded-history-inspect',sourceId,maintenanceBudget}});
  if(retained.deferred){await processorOutcome('deferred');process.exit(0);}
  runs=retained.needsRecovery?[retained.run]:[];
}
if (runs.some((run) => run.sourceId !== sourceId)) throw new Error("Roster queue returned work for a different source.");
console.log(`Found ${runs.length} queued roster file(s).`);
if ((!runs.length && !pending.publicationPending && !publicationOnly) || (pending.maintenanceDeferred&&!publicationOnly) || (process.env.ROSTER_AUTOMATION_RESUME_ONLY === "true" && pending.boundedImportEnabled !== true)) {
  console.log("No admitted import work; progress retained.");
  await processorOutcome(pending.maintenanceDeferred || runs.length || pending.publicationPending?'deferred':'completed');
  process.exit(0);
}
const publicationBudget={remaining:24};
if(pending.boundedImportEnabled&&(pending.publicationPending||publicationOnly)){
  const delivery=await drainPublication();
  const after=await automationRequest(`/api/automation/pending?limit=1&sourceId=${encodeURIComponent(sourceId)}`);
  const blocked=after.publicationPending||(delivery.deferred&&delivery.reason!=='run-limit');
  if(blocked||publicationOnly){
    processorReason=blocked?(delivery.reason||'publication-pending'):'';
    await processorOutcome(blocked?'deferred':'completed');process.exit(0);
  }
  if(!recoverVhhHistory) runs=Array.isArray(after.runs)?after.runs:[];
  if(!runs.length){await processorOutcome('completed');process.exit(0);}
  if(!publicationBudget.remaining){processorReason='run-limit';await processorOutcome('deferred');process.exit(0);}
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
 const delivery=await drainPublication();
 if(delivery.deferred){processorDeferred=true;processorReason=delivery.reason;}
}

async function drainPublication(){
 const sourceType={"monash-adults":"mmc","monash-paeds":"mch","dandenong-findmyshift":"ddh","vhh-active-medical-roster":"vhh"}[sourceId];
 const request=doctorCalendarRequest({base:baseUrl,token});
 return drainRosterPublication({budget:publicationBudget,
  step:()=>automationRequest('/api/automation/facility-refresh',{method:'POST',body:{sourceId,maintenanceBudget}}),
  publishMonth:r=>publishDoctorMonth({sourceType:r.sourceType||sourceType,month:r.month,revision:r.revision,request}),
  publishSource:r=>publishDoctorSource({sourceType:r.sourceType||sourceType,request,...r.doctorSource}),
 });
}

if (failures.length) {
  throw new Error(`${failures.length} roster file${failures.length === 1 ? "" : "s"} failed during background processing.`);
}
if(pending.boundedImportEnabled===true) {
  const remaining=await automationRequest(`/api/automation/pending?limit=1&sourceId=${encodeURIComponent(sourceId)}`);
  if(remaining.publicationPending || remaining.runs?.length || remaining.deferred || remaining.paused) processorDeferred=true;
  else if(processorReason==='run-limit'){processorDeferred=false;processorReason='';}
}
await processorOutcome(processorDeferred?'deferred':'completed');

async function processRun(run) {
  const response = await guardedFetch(process.env, `${baseUrl}/api/automation/raw?runId=${encodeURIComponent(run.id)}&sourceId=${encodeURIComponent(sourceId)}`, {
    headers: authorizationHeaders(),
  }, { label: "Roster processor download" });
  if (!response.ok) throw new Error(`Roster download returned HTTP ${response.status}.`);
  let payload;
  let historyWindows=[];
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
    run.boundedImportStarted=true;
    const preservation=await preserveVhhHistory({run,extract,recoveryOnly:recoverVhhHistory,
      request:body=>automationRequest('/api/automation/derived',{method:'POST',body:{...body,sourceId:run.sourceId,runId:run.id,maintenanceBudget}})});
    if(preservation.deferred){processorDeferred=true;return;}
    historyWindows=preservation.historyWindows;
    if(recoverVhhHistory) return;
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
    const plan = await attachRosterCompactPlan(await planRosterImportBatches(payload, {maximumBatchBytes:65536,maximumOccurrences:64}), payload);
    run.boundedImportStarted = true;
    const importStep = forceStaging => step => automationRequest("/api/automation/derived", {
      method: "POST", body: { ...step, historyWindows, supportsPatch: pending.patchImportEnabled === true, forceStaging, maintenanceBudget, runId: run.id, sourceId: run.sourceId, file: payload.file },
    });
    finished = await executeBoundedRosterImport(plan, importStep(false));
    if(finished.mode==='rebase'){
      try{finished=await executeRosterRebase(payload,finished,{startDate:plan.manifest.startDate,rosterEndDate:plan.manifest.rosterEndDate},importStep(false));}
      catch(error){
        if(error.code!=='ROSTER_INCREMENTAL_BUDGET'||!error.boundedReplacementRequired)throw error;
        // A genuinely large rewrite still uses the established, budgeted full
        // staging path. Ordinary autosaves reuse the invisible baseline.
        finished=await executeBoundedRosterImport(plan,importStep(true));
      }
    }
    if (finished.mode === "patch") {
      try { finished = await executeRosterPatch(payload,finished,importStep(false)); }
      catch(error) {
        if (error.code !== "ROSTER_INCREMENTAL_BUDGET" || !error.boundedReplacementRequired) throw error;
        finished = await executeBoundedRosterImport(plan,importStep(true));
      }
    }
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
      console.log("Roster import deferred by maintenance safeguards; queued progress retained for automatic continuation.");
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
  if(path!=="/api/automation/account-budget") {
    const admission=await accountAdmission();
    if(admission.deferred||admission.paused)return {ok:true,deferred:true,maintenanceDeferred:true};
  }
  if(!isReadOnlyPatchPage(path,options.body))return automationRequestAttempt(path,options);
  const read=()=>retryPatchRead(()=>automationRequestAttempt(path,options));
  const result=await read();
  if(!result.deferred)return result;
  // Expiry can race a callback even after the local deadline check. Renew
  // once, then retry only this read. A real budget deferral still stops work.
  const admission=await accountAdmission({force:true});
  if(admission.deferred||admission.paused)return result;
  return read();
}

async function automationRequestAttempt(path,options) {
  const readOnly=isReadOnlyPatchPage(path,options.body);
  const headers = { ...authorizationHeaders(), ...(options.body ? { "Content-Type": "application/json" } : {}) };
  let lastError = null;
  for (let attempt = 1; attempt <= (readOnly?1:6); attempt += 1) {
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
        // An uncertain derived callback is not retried in this pass. Keep
        // its durable cursor for a later invocation; other reads may retry.
        lastError = new Error(`Automation endpoint returned non-JSON HTTP ${response.status || 502}.`);
        if (path === "/api/automation/derived") {lastError.retryablePatchRead=readOnly&&(response.ok||retryablePatchReadFailure(response.status));break;}
        if (attempt < 6) {
          await new Promise((resolve) => setTimeout(resolve, attempt * 2000));
          continue;
        }
        break;
      }
      if (response.ok) return result;
      if(path==='/api/automation/facility-refresh'&&response.status===409&&result.stalePlan)return {...result,recoverableConflict:true};
      const diagnostic = String(result.reason || result.code || result.phase || "").trim();
      lastError = new Error(`${path} source=${sourceId} phase=${String(result.mode||options.body?.phase||'publication')} ${result.error || `HTTP ${response.status}`}${diagnostic ? ` (${diagnostic})` : ""}`);
      lastError.code = String(result.code || '');
      lastError.boundedReplacementRequired = result.boundedReplacementRequired === true;
      lastError.retryablePatchRead=readOnly&&retryablePatchReadFailure(response.status,result);
      if (path === "/api/automation/derived" || ![408, 429, 500, 502, 503, 504].includes(response.status)) break;
    } catch (error) {
      lastError = error;
      if(path==="/api/automation/derived"){lastError.retryablePatchRead=readOnly;break;}
    }
    await new Promise((resolve) => setTimeout(resolve, attempt * 2000));
  }
  throw lastError || new Error("Automation request failed.");
}

function authorizationHeaders() {
  return { Authorization: `Bearer ${token}` };
}
