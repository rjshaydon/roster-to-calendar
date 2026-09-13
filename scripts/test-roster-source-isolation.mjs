import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

import { requestQueuedRosterProcessing } from "../functions/_lib/automation-dispatch.js";
import { claimRosterDispatch, listQueuedRosterSyncRuns } from "../functions/_lib/d1-calendar.js";
import { advancedRosterMaintenanceEnabled, rosterWritesExplicitlyPaused } from "../functions/_lib/roster-automation-guard.js";
import { onRequestPost as derived } from "../functions/api/automation/derived.js";
import { onRequestPost as dispatch } from "../functions/api/automation/dispatch.js";
import { onRequestPost as findmyshift } from "../functions/api/automation/findmyshift-check.js";
import { onRequestPost as ingest } from "../functions/api/automation/ingest.js";
import { onRequestGet as parserConfig } from "../functions/api/automation/parser-config.js";
import { onRequestGet as pending } from "../functions/api/automation/pending.js";
import { onRequestGet as raw } from "../functions/api/automation/raw.js";
import { onRequestPost as vhh } from "../functions/api/automation/vhh-roster-extract.js";

const token = "source-isolation-token";
let d1Touches = 0;
let r2Touches = 0;
let outboundTouches = 0;
const forbiddenDb = {
  prepare() {
    d1Touches += 1;
    throw new Error("A disallowed source touched D1.");
  },
};
const forbiddenR2 = new Proxy({}, {
  get() {
    r2Touches += 1;
    throw new Error("A disallowed source touched R2.");
  },
});
const env = {
  ROSTER_AUTOMATION_TOKEN: token,
  VHH_AUTOMATION_TOKEN: token,
  ROSTER_AUTOMATION_WRITES_ENABLED: "true",
  ROSTER_AUTOMATION_QUEUE_ENABLED: "true",
  ROSTER_AUTOMATION_SOURCE_ALLOWLIST: "monash-adults",
  ROSTER_DB: forbiddenDb,
  ROSTER_FILES: forbiddenR2,
};

function post(path, body) {
  return new Request(`https://example.test${path}`, {
    method: "POST",
    headers: { authorization: `Bearer ${token}`, "content-type": "application/json" },
    body: JSON.stringify(body),
  });
}

const cases = [
  ["ingest", ingest, post("/api/automation/ingest", { sourceId: "monash-paeds", fileName: "paeds.xlsx", contentBase64: "AQ==" })],
  ["derived failure", derived, post("/api/automation/derived", { sourceId: "monash-paeds", runId: "other-run", phase: "failed", file: { id: "other-file" } })],
  ["dispatch", dispatch, post("/api/automation/dispatch", { action: "kick", sourceId: "monash-paeds" })],
  ["FindMyShift", findmyshift, post("/api/automation/findmyshift-check", {})],
  ["VHH", vhh, post("/api/automation/vhh-roster-extract", {})],
  ["pending", pending, new Request("https://example.test/api/automation/pending?sourceId=monash-paeds", { headers: { authorization: `Bearer ${token}` } })],
  ["raw", raw, new Request("https://example.test/api/automation/raw?sourceId=monash-paeds&runId=other-run", { headers: { authorization: `Bearer ${token}` } })],
  ["parser config", parserConfig, new Request("https://example.test/api/automation/parser-config?sourceId=monash-paeds", { headers: { authorization: `Bearer ${token}` } })],
];

for (const [label, handler, request] of cases) {
  const response = await handler({ request, env });
  assert.equal(response.status, 503, `${label} must reject a source outside the exact allowlist`);
  assert.equal((await response.json()).status, "paused", `${label} must return the controlled pause response`);
}
assert.equal(d1Touches, 0, "disallowed sources must perform zero D1 operations");
assert.equal(r2Touches, 0, "disallowed sources must perform zero R2 operations");

const wrongLifecycle = await dispatch({
  request: post("/api/automation/dispatch", { action: "lifecycle", event: "started", sourceId: "monash-adults", dispatchId: "dispatch:monash-paeds:wrong" }),
  env,
});
assert.equal(wrongLifecycle.status, 400, "a lifecycle callback cannot update another source's dispatch");
assert.equal(d1Touches, 0, "a mismatched lifecycle callback must stop before D1");

const originalFetch = globalThis.fetch;
globalThis.fetch = async () => {
  outboundTouches += 1;
  return new Response(null, { status: 204 });
};
try {
  const rejectedDispatch = await requestQueuedRosterProcessing(env, { sourceId: "monash-paeds", reason: "wrong-source" });
  assert.equal(rejectedDispatch.reason, "source-disabled");
  assert.equal(d1Touches, 0, "a disallowed dispatch must stop before D1");
  assert.equal(outboundTouches, 0, "a disallowed dispatch must stop before GitHub");
} finally {
  globalThis.fetch = originalFetch;
}

assert.equal(rosterWritesExplicitlyPaused({ ROSTER_AUTOMATION_WRITES_ENABLED: "true" }), true, "automatic authority must not open manual roster writes");
assert.equal(rosterWritesExplicitlyPaused({ MANUAL_ROSTER_WRITES_ENABLED: "true" }), false, "manual roster writes require their own explicit authority");
assert.equal(advancedRosterMaintenanceEnabled({ ROSTER_AUTOMATION_WRITES_ENABLED: "true", ROSTER_ADVANCED_MAINTENANCE_ENABLED: "true" }), false, "automatic authority must not open advanced maintenance");
assert.equal(advancedRosterMaintenanceEnabled({ MANUAL_ROSTER_WRITES_ENABLED: "true", ROSTER_ADVANCED_MAINTENANCE_ENABLED: "true" }), true);

const scopedStatements = [];
const scopedDb = {
  prepare(sql) {
    const statement = {
      sql: String(sql).replace(/\s+/g, " ").trim(),
      args: [],
      bind(...args) { this.args = args; scopedStatements.push(this); return this; },
      async first() {
        if (this.sql.startsWith("SELECT id FROM roster_sync_runs")) return { id: "adult-run" };
        if (this.sql.startsWith("SELECT * FROM roster_dispatches WHERE status IN")) return null;
        return null;
      },
      async all() { return { results: [] }; },
      async run() { return { success: true, meta: { changes: 1 } }; },
    };
    return statement;
  },
};
await listQueuedRosterSyncRuns(scopedDb, "monash-adults", 1);
const queueStatement = scopedStatements.find((statement) => statement.sql.includes("raw_roster_files.name AS file_name"));
assert.match(queueStatement.sql, /WHERE roster_sync_runs\.source_id = \?/);
assert.deepEqual(queueStatement.args, ["monash-adults", 1], "queue SQL must bind the exact source before its limit");
await claimRosterDispatch(scopedDb, { sourceId: "monash-adults", reason: "test", now: "2026-09-13T00:00:00Z", retryAfter: "2026-09-13T00:02:00Z" });
const pendingStatement = scopedStatements.find((statement) => statement.sql.startsWith("SELECT id FROM roster_sync_runs"));
assert.match(pendingStatement.sql, /WHERE source_id = \?/);
assert.deepEqual(pendingStatement.args, ["monash-adults"], "dispatch claim must bind the exact source");

const workflowSource = await readFile(new URL("../.github/workflows/monash-roster-sync.yml", import.meta.url), "utf8");
const processorSource = await readFile(new URL("./process-roster-queue.mjs", import.meta.url), "utf8");
assert.match(workflowSource, /source_id:[\s\S]*required: true[\s\S]*ROSTER_AUTOMATION_SOURCE_ID: \$\{\{ inputs\.source_id \}\}/, "workflow dispatch must require an exact source");
assert.match(workflowSource, /pending\?limit=1&sourceId=\$\{ROSTER_AUTOMATION_SOURCE_ID\}/, "workflow must inspect one source-scoped queue item");
assert.match(workflowSource, /sourceId\\\":\\\"\$\{SOURCE_ID\}/, "workflow lifecycle callbacks must carry the exact source");
assert.match(processorSource, /ROSTER_AUTOMATION_SOURCE_ID is required/, "processor must reject an omitted source");
assert.match(processorSource, /pending\?limit=1&sourceId=/, "processor must retrieve one source-scoped queue item");
assert.match(processorSource, /raw\?runId=.*&sourceId=/, "processor raw reads must carry the same source");

console.log("Roster automation source isolation and manual-write separation checks passed.");
