import { spawnSync } from "node:child_process";
import { guardedFetch } from "../functions/_lib/outbound-network.js";

const base = String(process.env.ROSTER_AUTOMATION_BASE_URL || "https://roster-to-calendar.pages.dev").replace(/\/$/, "");
const token = String(process.env.ROSTER_AUTOMATION_TOKEN || "");
if (!token) throw new Error("ROSTER_AUTOMATION_TOKEN required.");
const sources = ["monash-adults", "monash-paeds", "dandenong-findmyshift", "vhh-active-medical-roster"];
const offset = Math.floor(Date.now() / 86400000) % sources.length;
const failures = [];
async function request(path, body) {
  const response = await guardedFetch(process.env, `${base}${path}`, { method: body ? "POST" : "GET", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" }, ...(body ? { body: JSON.stringify(body) } : {}) }, { label: "Bounded roster maintenance" });
  if (response.status === 503) return { paused: true };
  const result = await response.json();
  if (!response.ok) throw new Error(`${path}: ${result.reason || result.error || response.status}`);
  return result;
}
for (const sourceId of [...sources.slice(offset), ...sources.slice(0, offset)]) {
  try {
    const pending = await request(`/api/automation/pending?limit=1&sourceId=${encodeURIComponent(sourceId)}`);
    if (pending.paused || !pending.boundedImportEnabled || pending.maintenanceDeferred) continue;
    if (process.env.ROSTER_SEED_CURRENT_VIEWS === "true") await request("/api/automation/facility-refresh", { sourceId, seedCurrent: true });
    if (sourceId === "dandenong-findmyshift") await request("/api/automation/findmyshift-check", {});
    const parsed = spawnSync(process.execPath, ["scripts/process-roster-queue.mjs"], { stdio: "inherit", env: { ...process.env, ROSTER_AUTOMATION_SOURCE_ID: sourceId, ROSTER_AUTOMATION_RESUME_ONLY: "true" }, timeout: 480000 });
    if (parsed.status !== 0) throw new Error(`${sourceId}: processor failed.`);
    // At most one term (120 dates / seven per batch), plus plan/month/finalize.
    for (let step = 0; step < 24; step += 1) {
      const result = await request("/api/automation/facility-refresh", { sourceId });
      if (result.paused || result.idle || result.deferred || result.completed) break;
    }
  } catch (error) { failures.push(String(error.message || error)); }
}
if (failures.length) throw new Error(failures.join("\n"));
