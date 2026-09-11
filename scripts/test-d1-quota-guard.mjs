import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

import { onRequestPost as ingest } from "../functions/api/automation/ingest.js";
import { onRequestPost as derived } from "../functions/api/automation/derived.js";
import { onRequestPost as findmyshiftCheck } from "../functions/api/automation/findmyshift-check.js";
import { onRequestPost as dispatch } from "../functions/api/automation/dispatch.js";
import { onRequestPost as vhhExtract } from "../functions/api/automation/vhh-roster-extract.js";
import { onRequestGet as pending } from "../functions/api/automation/pending.js";
import { onRequestPost as contactExtract } from "../functions/api/automation/contact-list-extract.js";
import { onRequestPost as facilityBootstrap } from "../functions/api/automation/facility-bootstrap.js";
import { ensureCalendarSchema } from "../functions/_lib/d1-calendar.js";
import watchdog from "../worker/roster-queue-watchdog.js";

let databaseTouches = 0;
let objectStoreTouches = 0;
const blockedDb = {
  prepare() {
    databaseTouches += 1;
    throw new Error("Paused automation must not touch D1.");
  },
};
const blockedObjectStore = {
  put() {
    objectStoreTouches += 1;
    throw new Error("Paused automation must not touch R2.");
  },
};
const env = {
  ROSTER_AUTOMATION_TOKEN: "test-token",
  ROSTER_WATCHDOG_TOKEN: "test-token",
  VHH_AUTOMATION_TOKEN: "test-token",
  ROSTER_AUTOMATION_WRITES_ENABLED: "false",
  ROSTER_AUTOMATION_ENABLED: "false",
  ROSTER_DB: blockedDb,
  ROSTER_FILES: blockedObjectStore,
};

function request(path, body = {}) {
  return new Request(`https://example.test${path}`, {
    method: "POST",
    headers: { Authorization: "Bearer test-token", "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
}

for (const [path, handler] of [
  ["/api/automation/ingest", ingest],
  ["/api/automation/derived", derived],
  ["/api/automation/findmyshift-check", findmyshiftCheck],
  ["/api/automation/dispatch", dispatch],
  ["/api/automation/vhh-roster-extract", vhhExtract],
]) {
  const response = await handler({ request: request(path), env });
  const payload = await response.json();
  assert.equal(response.status, 503, `${path} should fail closed while automation is paused`);
  assert.equal(payload.status, "paused", `${path} should identify the quota pause`);
}

const pendingResponse = await pending({
  request: new Request("https://example.test/api/automation/pending", {
    headers: { Authorization: "Bearer test-token" },
  }),
  env,
});
assert.equal(pendingResponse.status, 503, "queue polling should fail closed while automation is paused");
assert.equal((await pendingResponse.json()).status, "paused");

const contactResponse = await contactExtract({
  request: request("/api/automation/contact-list-extract", { sourceId: "mmc-shift-allocations", sourceDate: "2026-09-06", providerModifiedAt: "2026-09-06T00:00:00Z", contacts: [] }),
  env,
});
assert.equal(contactResponse.status, 503, "contact ingestion should fail closed independently of roster automation");
assert.equal((await contactResponse.json()).status, "paused");

assert.equal(databaseTouches, 0, "paused roster automation must perform zero D1 operations");
assert.equal(objectStoreTouches, 0, "paused roster automation must perform zero R2 operations");

for (const [label, body, bootstrapEnv, expectedStatus = 503] of [
  ["disabled inspection", { sourceType: "mmc", fileId: "exact-file" }, {}],
  ["wrong inspection file", { sourceType: "mmc", fileId: "wrong-file" }, {
    FACILITY_BOOTSTRAP_INSPECTION_ENABLED: "true",
    FACILITY_BOOTSTRAP_FILE_ALLOWLIST: "exact-file",
    FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST: "mmc",
  }],
  ["multiple-file allowlist", { sourceType: "mmc", fileId: "exact-file" }, {
    FACILITY_BOOTSTRAP_INSPECTION_ENABLED: "true",
    FACILITY_BOOTSTRAP_FILE_ALLOWLIST: "exact-file,other-file",
    FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST: "mmc",
  }],
  ["malformed limit", { sourceType: "mmc", fileId: "exact-file", maximumEventRows: "not-a-number" }, {
    FACILITY_BOOTSTRAP_INSPECTION_ENABLED: "true",
    FACILITY_BOOTSTRAP_FILE_ALLOWLIST: "exact-file",
    FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST: "mmc",
  }],
  ["inspection cannot execute", { sourceType: "mmc", fileId: "exact-file", execute: true }, {
    FACILITY_BOOTSTRAP_INSPECTION_ENABLED: "true",
    FACILITY_BOOTSTRAP_FILE_ALLOWLIST: "exact-file",
    FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST: "mmc",
  }],
  ["expired execution plan", { sourceType: "mmc", fileId: "exact-file", execute: true, planRevision: "any", planGeneratedAt: "2026-01-01T00:00:00.000Z" }, {
    FACILITY_BOOTSTRAP_EXECUTION_ENABLED: "true",
    FACILITY_BOOTSTRAP_FILE_ALLOWLIST: "exact-file",
    FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST: "mmc",
    ROSTER_ADVANCED_MAINTENANCE_ENABLED: "true",
    FACILITY_SHARED_EMERGENCY_PAUSED: "false",
  }, 409],
]) {
  const response = await facilityBootstrap({
    request: request("/api/automation/facility-bootstrap", body),
    env: { ...env, ...bootstrapEnv },
  });
  assert.equal(response.status, expectedStatus, `${label} must fail before D1`);
}
assert.equal(databaseTouches, 0, "disabled and mismatched bootstrap capabilities must perform zero D1 operations");

assert.equal(await ensureCalendarSchema(blockedDb), true, "deployed databases are managed by migrations");
assert.equal(databaseTouches, 0, "ordinary requests must not inspect or modify the D1 schema");

let scheduledWork = 0;
await watchdog.scheduled({}, env, { waitUntil() { scheduledWork += 1; } });
assert.equal(scheduledWork, 0, "the paused watchdog must not call production endpoints");

const health = await watchdog.fetch(new Request("https://watchdog.test/health"), env);
assert.equal((await health.json()).paused, true, "watchdog health should expose its paused state");

const stateSource = await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8");
const appSource = await readFile(new URL("../public/static/app.js", import.meta.url), "utf8");
const ensureInviteBody = stateSource.match(/async function ensureInviteSchema[\s\S]*?\n}/)?.[0] || "";
assert.match(ensureInviteBody, /ensureCalendarSchema\(db\)/, "invite setup should use the shared schema check");
assert.doesNotMatch(ensureInviteBody, /CREATE\s+(?:TABLE|INDEX)/i, "ordinary API requests must not issue invite DDL directly");

for (const action of [
  "syncRosterRepository",
  "removeRosterImports",
  "saveDerivedCalendarFile",
  "uploadRawRosterFile",
  "resetDerivedCalendarFile",
  "replaceActiveRosterFiles",
  "repairRosterDailyPresence",
]) {
  const actionBody = stateSource.match(new RegExp(`if \\(action === "${action}"\\)[\\s\\S]*?(?=\\n    if \\(action === |$)`))?.[0] || "";
  assert.match(actionBody, /rosterWritesExplicitlyPaused\(context\.env\)/, `${action} must stop while roster writes are paused`);
}
const calendarStoreStatusAction = stateSource.match(/if \(action === "calendarStoreStatus"\)[\s\S]*?(?=\n    if \(action === |$)/)?.[0] || "";
assert.ok(
  calendarStoreStatusAction.indexOf("rosterStatusSummaryEnabled(context.env)")
    < calendarStoreStatusAction.indexOf("calendarStoreStatus(null, context.env.ROSTER_DB"),
  "disabled compact calendar status must stop before roster repository reads",
);
assert.doesNotMatch(
  calendarStoreStatusAction,
  /rosterWritesExplicitlyPaused\(context\.env\)/,
  "the compact read-only status must remain available while roster writes are paused",
);
const calendarStoreStatusBody = stateSource.match(/async function calendarStoreStatus[\s\S]*?function rosterSourceStatuses/)?.[0] || "";
assert.match(calendarStoreStatusBody, /queryRosterFileStatusSummaries/);
assert.doesNotMatch(calendarStoreStatusBody, /queryRosterFiles|queryRawRosterFiles|countDerivedEventsByFile|countDerivedDoctorsByFile|roster_events|roster_file_doctors/,
  "normal calendar status must not contain a historical roster scan");
assert.match(
  stateSource,
  /removedImportIds\.length && rosterWritesExplicitlyPaused\(context\.env\)/,
  "save must not remove roster files while roster writes are paused",
);
const statusBody = appSource.match(/function setStatus\([\s\S]*?(?=\nfunction removeSupersededStatusMessages)/)?.[0] || "";
assert.doesNotMatch(statusBody, /persistConsoleMessage|appendConsoleMessage/, "ordinary UI status messages must not create D1 console-history writes");
const postLoginRefreshBody = appSource.match(/function queuePostLoginSnapshotRefresh[\s\S]*?(?=\nfunction markLoginPhase)/)?.[0] || "";
assert.match(postLoginRefreshBody, /for \(const delayMs of \[1500\]\)/, "post-login snapshot refresh must have one bounded retry");
assert.match(postLoginRefreshBody, /if \(document\.hidden\) return/, "hidden tabs must not retry a post-login snapshot refresh");
const byStreamOpenBody = appSource.match(/async function openFacilityOverviewByStream[\s\S]*?(?=\nfunction closeFacilityOverview)/)?.[0] || "";
assert.equal((byStreamOpenBody.match(/loadFacilityOverviewMetadata\(\)/g) || []).length, 1, "opening By stream must not immediately repeat a failed metadata request");

console.log("D1 quota emergency guards passed.");
