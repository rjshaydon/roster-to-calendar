import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

import { onRequestPost as ingest } from "../functions/api/automation/ingest.js";
import { onRequestPost as derived } from "../functions/api/automation/derived.js";
import { onRequestPost as findmyshiftCheck } from "../functions/api/automation/findmyshift-check.js";
import { onRequestPost as dispatch } from "../functions/api/automation/dispatch.js";
import { onRequestPost as vhhExtract } from "../functions/api/automation/vhh-roster-extract.js";
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

assert.equal(databaseTouches, 0, "paused roster automation must perform zero D1 operations");
assert.equal(objectStoreTouches, 0, "paused roster automation must perform zero R2 operations");

let scheduledWork = 0;
await watchdog.scheduled({}, env, { waitUntil() { scheduledWork += 1; } });
assert.equal(scheduledWork, 0, "the paused watchdog must not call production endpoints");

const health = await watchdog.fetch(new Request("https://watchdog.test/health"), env);
assert.equal((await health.json()).paused, true, "watchdog health should expose its paused state");

const stateSource = await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8");
const ensureInviteBody = stateSource.match(/async function ensureInviteSchema[\s\S]*?\n}/)?.[0] || "";
assert.match(ensureInviteBody, /ensureCalendarSchema\(db\)/, "invite setup should use the shared schema check");
assert.doesNotMatch(ensureInviteBody, /CREATE\s+(?:TABLE|INDEX)/i, "ordinary API requests must not issue invite DDL directly");

console.log("D1 quota emergency guards passed.");
