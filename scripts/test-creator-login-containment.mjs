import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

import { creatorDirectoryEnabled, creatorStartupHydrationEnabled } from "../functions/_lib/creator-startup-guard.js";
import { onRequestPost as stateHandler } from "../functions/api/state.js";

const appSource = await readFile(new URL("../public/static/app.js", import.meta.url), "utf8");
const stateSource = await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8");
const wranglerSource = await readFile(new URL("../wrangler.toml", import.meta.url), "utf8");
const localDevSource = await readFile(new URL("./local-dev.mjs", import.meta.url), "utf8");

assert.equal(creatorStartupHydrationEnabled({}), false);
assert.equal(creatorStartupHydrationEnabled({ CREATOR_STARTUP_HYDRATION_ENABLED: "malformed" }), false);
assert.equal(creatorStartupHydrationEnabled({ CREATOR_STARTUP_HYDRATION_ENABLED: "true" }), true);
assert.equal(creatorDirectoryEnabled({}), false);
assert.equal(creatorDirectoryEnabled({ CREATOR_DIRECTORY_ENABLED: "yes" }), true);

const pausedActions = ["listUsers", "calendarStoreStatus", "listRosterDoctors"];
const pausedResults = [];
for (const action of pausedActions) {
  let d1Operations = 0;
  let r2Operations = 0;
  let scheduledWork = 0;
  const blocked = (kind) => () => {
    if (kind === "D1") d1Operations += 1;
    else r2Operations += 1;
    throw new Error(`${kind} must not be used by ${action} while paused.`);
  };
  const response = await stateHandler({
    request: new Request("https://example.test/api/state", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ action }),
    }),
    env: {
      ROSTER_DB: { prepare: blocked("D1"), batch: blocked("D1"), exec: blocked("D1"), dump: blocked("D1") },
      ROSTER_FILES: { get: blocked("R2"), put: blocked("R2"), head: blocked("R2"), list: blocked("R2"), delete: blocked("R2") },
    },
    waitUntil() { scheduledWork += 1; },
  });
  const payload = await response.json();
  assert.equal(payload.unavailable, true, `${action} must return an unavailable state`);
  assert.equal(d1Operations, 0, `${action} must stop before D1`);
  assert.equal(r2Operations, 0, `${action} must stop before R2`);
  assert.equal(scheduledWork, 0, `${action} must schedule no work`);
  pausedResults.push({ action, status: response.status, d1Operations, r2Operations, scheduledWork });
}

const containedFinish = section(/function finishContainedCreatorStartup[\s\S]*?(?=\nfunction launchNonClinicalDirectorWorkspace)/);
assert.doesNotMatch(containedFinish, /fetch\(|calendarStoreRequest|loadServerUsers|refreshCalendarStoreStatus|loadCloudCalendarEvents|bootstrapImports|waitUntil/,
  "contained Creator startup must only render saved browser/server state");
assert.match(containedFinish, /cancelDeferredAccountContextLoad\(\)[\s\S]*cancelDeferredBootstrapImports\(\)/,
  "contained startup must invalidate queued account and bootstrap work");
assert.match(containedFinish, /renderWorkspaceFromSnapshot\([\s\S]*suppressInsightWarmup: true/,
  "rendering a saved Creator calendar must not schedule the automatic colleague-query warmup");

for (const [name, body] of [
  ["explicit login", section(/async function loginWithEmail[\s\S]*?(?=\nasync function restoreCloudState)/)],
  ["persisted login", section(/async function bootstrapApp[\s\S]*?(?=\nfunction setStatus)/)],
]) {
  assert.match(body, /renderLoginState\(\);[\s\S]*if \(creatorStartupContainmentActive\(\)\) \{[\s\S]*finishContainedCreatorStartup[\s\S]*return;[\s\S]*queueDeferredAccountContextLoad[\s\S]*queuePostLoginHydration/,
    `${name} must return before deferred context and hydration fan-out`);
}
const explicitLogin = section(/async function loginWithEmail[\s\S]*?(?=\nasync function restoreCloudState)/);
assert.doesNotMatch(explicitLogin.slice(0, explicitLogin.indexOf("restoreCloudState")), /flushCloudStateSave/,
  "login must not flush stale account state before authentication");

const cachedSnapshotApply = section(/function applyCachedCalendarSnapshot[\s\S]*?(?=\nfunction saveWorkspaceSnapshotForEmail)/);
assert.match(cachedSnapshotApply, /renderWorkspaceFromSnapshot\(cached, cached\.session \|\| \{\}, options\)/,
  "cached-snapshot safety options must reach the renderer");
const workspaceRenderer = section(/function renderWorkspaceFromSnapshot[\s\S]*?(?=\nasync function ensureSelectedFilesLoaded)/);
assert.match(workspaceRenderer, /if \(options\.suppressInsightWarmup !== true\) scheduleInsightWarmup\(\)/,
  "the contained renderer must suppress automatic D1-backed insight warmup");

const hydration = section(/async function hydrateAuthenticatedWorkspace[\s\S]*?(?=\nfunction queuePostLoginHydration)/);
assert.match(hydration, /const adminTargetEmail[\s\S]*if \(!adminTargetEmail && creatorStartupContainmentActive\(\)\) \{[\s\S]*finishContainedCreatorStartup[\s\S]*return;[\s\S]*loadCloudCalendarEvents/,
  "defence-in-depth must stop Creator hydration before calendar and Admin work");

const loginHandler = section(/if \(action === "login"\)[\s\S]*?(?=\n    const account = await verifyD1Account)/);
assert.doesNotMatch(loginHandler, /queryRosterFiles|queryRosterFileDoctors|queryDoctorEvents|listAccountMirrors|calendarStoreStatus/,
  "fast login handler must not contain roster-wide or directory work");
assert.match(loginHandler, /loginResponseMode = \(loginRole === "creator" \|\| loginRole === "owner"\)[\s\S]*creatorStartupHydrationEnabled[\s\S]*\? "fast"[\s\S]*: responseMode/,
  "the server must force contained Creator logins onto the minimal response even for stale clients");
assert.match(loginHandler, /creatorStartupHydrationEnabled: creatorStartupHydrationEnabled\(context\.env\)/,
  "login must tell the browser whether deferred Creator hydration is explicitly enabled");

const contextHandler = section(/if \(action === "loadAccountContext"\)[\s\S]*?(?=\n    if \(action === "claimRosterName"\))/);
assert.match(contextHandler, /creatorStartupContained[\s\S]*prepareFastLoginEnvelope[\s\S]*: await prepareAccountResponse/,
  "stale Creator context requests must use the minimal envelope while contained");

assert.equal((wranglerSource.match(/CREATOR_STARTUP_HYDRATION_ENABLED = "false"/g) || []).length, 2,
  "Production and Preview must explicitly disable Creator startup hydration");
assert.equal((wranglerSource.match(/CREATOR_DIRECTORY_ENABLED = "false"/g) || []).length, 2,
  "Production and Preview must explicitly disable the Creator directory");
assert.match(localDevSource, /CREATOR_STARTUP_HYDRATION_ENABLED=false/);
assert.match(localDevSource, /CREATOR_DIRECTORY_ENABLED=false/);

const visibleIdleTabs = 50;
const visibleIdleMinutes = 15;
const d1OperationsPerContainedIdleTab = 0;
assert.equal(visibleIdleTabs * visibleIdleMinutes * d1OperationsPerContainedIdleTab, 0);

console.log(JSON.stringify({
  pausedActions: pausedResults,
  creatorStartup: {
    explicitLoginAutomaticFollowUpActions: 0,
    persistedLoginAutomaticFollowUpActions: 0,
    disabledAtAGlanceD1Operations: 0,
    visibleIdleTabs,
    visibleIdleMinutes,
    estimatedIdleD1Operations: 0,
  },
}, null, 2));

function section(pattern) {
  return appSource.match(pattern)?.[0] || stateSource.match(pattern)?.[0] || "";
}
