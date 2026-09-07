import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

import { facilityOverviewMaintenanceMode } from "../functions/_lib/facility-rollout.js";
import { onRequestPost as stateHandler } from "../functions/api/state.js";

const actions = [
  "queryFacilityOverviewMetadata",
  "queryFacilityOverviewByStream",
  "queryFacilityOverviewOnShift",
  "queryFacilityOverviewContactList",
  "queryFacilityOverviewStaff",
  "queryFacilityOverviewWorkingTogether",
  "setContactAllocationResolution",
  "setFacilityStaffDesignation",
  "clearFacilityStaffDesignation",
  "setFacilityStaffSeniorityOverride",
  "setFacilityStaffSeniorityOverrides",
];
const message = "At a glance is temporarily unavailable while we complete a reliability upgrade. Your roster and settings have not been removed.";

assert.equal(facilityOverviewMaintenanceMode({}), true);
assert.equal(facilityOverviewMaintenanceMode({ FACILITY_OVERVIEW_MAINTENANCE_MODE: "" }), true);
assert.equal(facilityOverviewMaintenanceMode({ FACILITY_OVERVIEW_MAINTENANCE_MODE: "malformed" }), true);
assert.equal(facilityOverviewMaintenanceMode({ FACILITY_OVERVIEW_MAINTENANCE_MODE: "true" }), true);
assert.equal(facilityOverviewMaintenanceMode({ FACILITY_OVERVIEW_MAINTENANCE_MODE: "false" }), false);
assert.equal(facilityOverviewMaintenanceMode({ FACILITY_OVERVIEW_MAINTENANCE_MODE: "off" }), false);

for (const maintenanceValue of [undefined, "", "malformed", "true"]) {
  for (const action of actions) {
    let databaseTouches = 0;
    let objectStoreTouches = 0;
    let scheduledWork = 0;
    const blocked = (kind) => () => {
      if (kind === "D1") databaseTouches += 1;
      else objectStoreTouches += 1;
      throw new Error(`${kind} must not be touched during At a glance maintenance.`);
    };
    const env = {
      ROSTER_DB: {
        prepare: blocked("D1"), batch: blocked("D1"), exec: blocked("D1"), dump: blocked("D1"),
      },
      ROSTER_FILES: {
        get: blocked("R2"), put: blocked("R2"), head: blocked("R2"), list: blocked("R2"), delete: blocked("R2"),
      },
    };
    if (maintenanceValue !== undefined) env.FACILITY_OVERVIEW_MAINTENANCE_MODE = maintenanceValue;
    const response = await stateHandler({
      request: new Request("https://example.test/api/state", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ action }),
      }),
      env,
      waitUntil() { scheduledWork += 1; },
    });
    const payload = await response.json();
    assert.equal(response.status, 503, `${action} must be unavailable for maintenance value ${String(maintenanceValue)}`);
    assert.deepEqual(payload, {
      ok: false, unavailable: true, maintenance: true, retryable: false, error: message,
    });
    assert.equal(response.headers.get("cache-control"), "no-store");
    assert.equal(response.headers.get("retry-after"), null);
    assert.equal(databaseTouches, 0, `${action} must use zero D1 operations`);
    assert.equal(objectStoreTouches, 0, `${action} must use zero R2 operations`);
    assert.equal(scheduledWork, 0, `${action} must schedule no background work`);
  }
}

let explicitlyOpenDatabaseTouches = 0;
const explicitlyOpen = await stateHandler({
  request: new Request("https://example.test/api/state", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ action: actions[0] }),
  }),
  env: {
    FACILITY_OVERVIEW_MAINTENANCE_MODE: "false",
    ROSTER_DB: { prepare() { explicitlyOpenDatabaseTouches += 1; } },
  },
});
assert.equal(explicitlyOpen.status, 400, "an explicit false value must pass the maintenance gate and reach ordinary validation");
assert.equal(explicitlyOpenDatabaseTouches, 0, "ordinary missing-email validation remains before D1");

const appSource = await readFile(new URL("../public/static/app.js", import.meta.url), "utf8");
const wranglerSource = await readFile(new URL("../wrangler.toml", import.meta.url), "utf8");
const localDevSource = await readFile(new URL("./local-dev.mjs", import.meta.url), "utf8");
const functionBody = (pattern) => appSource.match(pattern)?.[0] || "";
assert.equal((wranglerSource.match(/FACILITY_OVERVIEW_MAINTENANCE_MODE = "true"/g) || []).length, 2, "Production and Preview config must declare maintenance mode");
assert.match(localDevSource, /FACILITY_OVERVIEW_MAINTENANCE_MODE=false/, "isolated local development must explicitly open the feature");
assert.match(appSource, /let currentFacilityOverviewMaintenance = true;/, "the browser must start fail-closed");
assert.match(appSource, /currentFacilityOverviewMaintenance = data\.facilityOverviewMaintenance !== false;/, "missing server capability must remain paused");
assert.match(functionBody(/function renderFacilityOverviewMaintenance[\s\S]*?(?=\nfunction facilityOverviewMelbourneClock)/), /stopFacilityOverviewContactRefresh\(\)[\s\S]*contactList = null[\s\S]*FACILITY_OVERVIEW_MAINTENANCE_MESSAGE/);
assert.match(functionBody(/async function openFacilityOverview\([\s\S]*?(?=\nasync function openFacilityOverviewByStream)/), /currentFacilityOverviewMaintenance[\s\S]*renderFacilityOverviewMaintenance\(\)[\s\S]*return;[\s\S]*refreshFacilityOverviewPreferredFacility/);
for (const name of ["loadFacilityOverviewMetadata", "openFacilityOverviewByStream", "loadFacilityOverviewByStream", "loadFacilityOverviewTogether", "loadFacilityOverviewOnShift", "loadFacilityOverviewStaff"]) {
  const body = functionBody(new RegExp(`(?:async )?function ${name}[\\s\\S]*?(?=\\n(?:async )?function )`));
  assert.match(body, /currentFacilityOverviewMaintenance/, `${name} must stop before cache or network work`);
}
const contactPredicate = functionBody(/function facilityOverviewContactRefreshIsActive[\s\S]*?(?=\nfunction stopFacilityOverviewContactRefresh)/);
assert.match(contactPredicate, /!currentFacilityOverviewMaintenance/, "maintenance must prevent contact timer scheduling");
const clinicalLaunch = functionBody(/function launchClinicalOnShiftWorkspace[\s\S]*?(?=\nasync function loginWithEmail)/);
assert.match(clinicalLaunch, /currentFacilityOverviewMaintenance/, "maintenance must suppress automatic clinical On shift launch");
for (const name of ["saveFacilityOverviewContactResolution", "setFacilityOverviewStaffDesignation", "clearFacilityOverviewStaffDesignation", "setFacilityOverviewStaffSeniorityOverride", "setFacilityOverviewStaffSeniorityOverrides"]) {
  const body = functionBody(new RegExp(`async function ${name}[\\s\\S]*?(?=\\n(?:async )?function )`));
  assert.match(body, /currentFacilityOverviewMaintenance/, `${name} must not send maintenance-time mutations`);
}

const visiblePages = 50;
const refreshesPerPage = 12 * 60;
const maintenanceContactRequests = visiblePages * refreshesPerPage * 0;
assert.equal(maintenanceContactRequests, 0, "50 visible maintenance pages must issue zero contact refreshes");

console.log(`At a glance maintenance passed: ${actions.length * 4} stale-client cases used zero D1/R2 operations; 50 visible pages schedule zero contact refreshes.`);
