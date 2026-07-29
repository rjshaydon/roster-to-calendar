import { buildAutomatedDerivedRosterPayload } from "../functions/_lib/automation-import.js";

const baseUrl = String(process.env.ROSTER_AUTOMATION_BASE_URL || "https://roster-to-calendar.pages.dev").replace(/\/$/, "");
const token = String(process.env.ROSTER_AUTOMATION_TOKEN || "");
const doctorChunkSize = 18;

if (!token) throw new Error("ROSTER_AUTOMATION_TOKEN is required.");

const pending = await automationRequest("/api/automation/pending?limit=4");
const runs = Array.isArray(pending.runs) ? pending.runs : [];
console.log(`Found ${runs.length} queued roster file(s).`);
const failures = [];

for (const run of runs) {
  try {
    await processRun(run);
  } catch (error) {
    const message = `Failed to process ${run.fileName || run.id}: ${error?.message || error}`;
    console.error(message);
    failures.push(message);
    await automationRequest("/api/automation/derived", {
      method: "POST",
      body: {
        runId: run.id,
        sourceId: run.sourceId,
        phase: "failed",
        file: { id: run.fileId, name: run.fileName, sourceId: run.sourceId, sourceType: run.sourceType },
        message: "Background processor could not parse or save this roster.",
      },
    }).catch(() => null);
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
  const bytes = new Uint8Array(await response.arrayBuffer());
  const file = new File([bytes], run.fileName || "roster.xlsx", {
    type: run.contentType || "application/octet-stream",
    lastModified: Number(run.lastModified || Date.now()),
  });
  console.log(`Parsing ${file.name} (${file.size} bytes).`);
  const payload = await buildAutomatedDerivedRosterPayload({
    file,
    sourceId: run.sourceId,
    contentHash: run.contentHash,
    providerVersion: run.providerVersion,
  });
  payload.file = {
    ...payload.file,
    lastModified: Number(run.lastModified || payload.file.lastModified || Date.now()),
  };
  await postDerived(run, payload, "start", payload.doctors, {}, {});
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
  }
  const finished = await postDerived(run, payload, "finish", payload.doctors, {}, {});
  console.log(`Indexed ${file.name}: ${finished.doctorCount} doctors, ${finished.eventCount} shifts.`);
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
  for (let attempt = 1; attempt <= 4; attempt += 1) {
    try {
      const response = await fetch(`${baseUrl}${path}`, {
        method: options.method || "GET",
        headers,
        body: options.body ? JSON.stringify(options.body) : undefined,
      });
      const text = await response.text();
      const result = text ? JSON.parse(text) : {};
      if (response.ok) return result;
      lastError = new Error(result.error || `HTTP ${response.status}`);
      if (![408, 429, 500, 502, 503, 504].includes(response.status)) break;
    } catch (error) {
      lastError = error;
    }
    await new Promise((resolve) => setTimeout(resolve, attempt * 1000));
  }
  throw lastError || new Error("Automation request failed.");
}

function authorizationHeaders() {
  return { Authorization: `Bearer ${token}` };
}
