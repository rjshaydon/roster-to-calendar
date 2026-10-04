import { spawnSync } from "node:child_process";
import { guardedFetch } from "../functions/_lib/outbound-network.js";

const base = String(process.env.ROSTER_AUTOMATION_BASE_URL || "https://roster-to-calendar.pages.dev").replace(/\/$/, "");
const token = String(process.env.ROSTER_AUTOMATION_TOKEN || "");
if (!token) throw new Error("ROSTER_AUTOMATION_TOKEN required.");
const allSources = ["monash-adults", "monash-paeds", "dandenong-findmyshift", "vhh-active-medical-roster"];
const requestedSource = String(process.env.ROSTER_AUTOMATION_SOURCE_ID || '');
if (requestedSource && !allSources.includes(requestedSource)) throw new Error('Unknown maintenance source.');
const sources = requestedSource ? [requestedSource] : allSources;
const offset = Math.floor(Date.now() / 86400000) % sources.length;
const failures = [];
let maintenanceBudget = null;
let lastAdmission = 0;
async function request(path, body) {
  if (path !== "/api/automation/account-budget" && Date.now() - lastAdmission > 5 * 60 * 1000) {
    const admission = await request("/api/automation/account-budget", {});
    lastAdmission = Date.now();
    if (admission.deferred || admission.paused) { console.log("Account budget deferred; durable progress retained."); return { paused: true, deferred: true }; }
  }
  const response = await guardedFetch(process.env, `${base}${path}`, { method: body ? "POST" : "GET", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" }, ...(body ? { body: JSON.stringify({ ...body, maintenanceBudget }) } : {}) }, { label: "Bounded roster maintenance" });
  if (response.status === 503) return { paused: true };
  const result = await response.json();
  if (!response.ok) throw new Error(`${path}: ${result.reason || result.error || response.status}`);
  return result;
}
if (process.env.ROSTER_SEED_CURRENT_VIEWS === "true") {
  const admission = await request(`/api/automation/pending?limit=1&sourceId=${sources[0]}`);
  maintenanceBudget = admission.maintenanceBudget;
  if (admission.boundedImportEnabled && !admission.maintenanceDeferred) {
    for (const sourceId of sources) await request("/api/automation/facility-refresh", { sourceId, mode: "queue-seed", seedCurrent: true });
  }
}
for (const sourceId of [...sources.slice(offset), ...sources.slice(0, offset)]) {
  try {
    const pending = await request(`/api/automation/pending?limit=1&sourceId=${encodeURIComponent(sourceId)}`);
    if (!maintenanceBudget && pending.maintenanceBudget) maintenanceBudget = pending.maintenanceBudget;
    if (pending.paused || !pending.boundedImportEnabled || pending.maintenanceDeferred) continue;
    let preparationDeferred = false;
    for (let step = 0; step < 32; step += 1) {
      const prepared = await request("/api/automation/facility-refresh", { sourceId, mode: "prepare-coverage" });
      if (prepared.deferred || prepared.paused) { preparationDeferred = true; break; }
      if (prepared.idle) break;
    }
    if (preparationDeferred) continue;
    if (sourceId === "dandenong-findmyshift") await request("/api/automation/findmyshift-check", {});
    const parsed = spawnSync(process.execPath, ["scripts/process-roster-queue.mjs"], { stdio: "inherit", env: { ...process.env, ROSTER_AUTOMATION_SOURCE_ID: sourceId, ROSTER_AUTOMATION_RESUME_ONLY: "true", ROSTER_MAINTENANCE_SESSION: JSON.stringify(maintenanceBudget) }, timeout: 480000 });
    if (parsed.status !== 0) throw new Error(`${sourceId}: processor failed.`);
    // At most one term (120 dates / seven per batch), plus plan/month/finalize.
    for (let step = 0; step < 24; step += 1) {
      const result = await request("/api/automation/facility-refresh", { sourceId });
      if (result.paused || result.idle || result.deferred || result.completed) break;
    }
  } catch (error) { failures.push(String(error.message || error)); }
}
if (failures.length) throw new Error(failures.join("\n"));
