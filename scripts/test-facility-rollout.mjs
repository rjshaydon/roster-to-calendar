import assert from "node:assert/strict";
import { facilityBuildSources, facilityOverviewMaintenanceMode, facilityReadRoute, facilityRolloutPaused, facilitySharedReaderAllowed } from "../functions/_lib/facility-rollout.js";
import { automatedRosterQueueEnabled, automatedRosterSourceEnabled, advancedRosterMaintenanceEnabled, facilityMaterializationMaintenanceEnabled } from "../functions/_lib/roster-automation-guard.js";
import { onRequestPost as materialize } from "../functions/api/automation/facility-materialize.js";

const env = {
  FACILITY_OVERVIEW_MAINTENANCE_MODE: "false",
  FACILITY_SHARED_ROLLOUT_ACTIVE: "true",
  FACILITY_SHARED_EMERGENCY_PAUSED: "false",
  FACILITY_LEGACY_READS_PAUSED: "false",
  FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST: "mmc",
  FACILITY_SHARED_READER_SOURCE_ALLOWLIST: "mmc",
  FACILITY_SHARED_READER_COHORT: "creator",
  ROSTER_AUTOMATION_WRITES_ENABLED: "true",
  ROSTER_AUTOMATION_SOURCE_ALLOWLIST: "monash-adults",
  ROSTER_AUTOMATION_QUEUE_ENABLED: "true",
  ROSTER_ADVANCED_MAINTENANCE_ENABLED: "true",
};
assert.equal(facilityOverviewMaintenanceMode({}), true, "missing maintenance configuration must fail closed");
assert.equal(facilityOverviewMaintenanceMode({ FACILITY_OVERVIEW_MAINTENANCE_MODE: "malformed" }), true, "malformed maintenance configuration must fail closed");
assert.equal(facilityOverviewMaintenanceMode(env), false, "only an explicit false value opens the maintenance gate");
assert.deepEqual(facilityBuildSources(env, ["MMC", "DDH"]), ["mmc"]);
assert.equal(facilitySharedReaderAllowed(env, { actorRole: "creator", actorEmail: "creator@example.com", subjectEmail: "creator@example.com", sources: ["mmc"] }), true);
assert.equal(facilitySharedReaderAllowed(env, { actorRole: "creator", actorEmail: "creator@example.com", subjectEmail: "other@example.com", sources: ["mmc"] }), false, "Creator impersonation must not enter the Creator-only cohort");
assert.equal(facilitySharedReaderAllowed(env, { actorRole: "user", actorEmail: "user@example.com", subjectEmail: "user@example.com", sources: ["mmc"] }), false);
assert.equal(facilitySharedReaderAllowed(env, { actorRole: "creator", actorEmail: "creator@example.com", subjectEmail: "creator@example.com", sources: ["ddh"] }), false);
assert.equal(facilityReadRoute(env, { actorRole: "creator", actorEmail: "creator@example.com", subjectEmail: "creator@example.com", sources: ["ddh"] }), "blocked");
assert.equal(facilityReadRoute(env, { actorRole: "user", actorEmail: "user@example.com", subjectEmail: "user@example.com", sources: ["mmc"] }), "legacy");
assert.equal(facilityReadRoute({ ...env, FACILITY_LEGACY_READS_PAUSED: "true" }, { actorRole: "user", actorEmail: "user@example.com", subjectEmail: "user@example.com", sources: ["mmc"] }), "blocked");
assert.equal(facilityRolloutPaused({ ...env, FACILITY_SHARED_EMERGENCY_PAUSED: "true" }), true);
assert.deepEqual(facilityBuildSources({ ...env, FACILITY_SHARED_EMERGENCY_PAUSED: "true" }, ["mmc"]), []);
assert.equal(facilityReadRoute({}, { actorRole: "creator", actorEmail: "creator@example.com", subjectEmail: "creator@example.com", sources: ["mmc"] }), "blocked", "missing rollout settings must not restore legacy reads");
assert.equal(facilityReadRoute({ FACILITY_LEGACY_READS_PAUSED: "malformed" }, { actorRole: "creator", actorEmail: "creator@example.com", subjectEmail: "creator@example.com", sources: ["mmc"] }), "blocked", "malformed legacy-read setting must fail closed");
assert.equal(facilityRolloutPaused({}), true, "missing emergency-pause setting must remain paused");
assert.equal(facilityRolloutPaused({ FACILITY_SHARED_EMERGENCY_PAUSED: "malformed" }), true, "malformed emergency-pause setting must remain paused");
assert.deepEqual(facilityBuildSources({ FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST: "mmc" }, ["mmc"]), [], "missing emergency-pause setting must block builders");
assert.equal(automatedRosterSourceEnabled(env, "monash-adults"), true);
assert.equal(automatedRosterSourceEnabled(env, "monash-paeds"), false);
assert.equal(automatedRosterQueueEnabled(env), true);
assert.equal(advancedRosterMaintenanceEnabled(env), true);
assert.equal(facilityMaterializationMaintenanceEnabled({ ROSTER_ADVANCED_MAINTENANCE_ENABLED: "true" }), true, "facility bootstrap may be enabled without general roster writes");
assert.equal(automatedRosterSourceEnabled({ ROSTER_AUTOMATION_WRITES_ENABLED: "true" }, "monash-adults"), false, "missing source allowlist must fail closed");

let databaseCalls = 0;
const pausedResponse = await materialize({
  request: new Request("http://local/api/automation/facility-materialize", { method: "POST", headers: { authorization: "Bearer token", "content-type": "application/json" }, body: JSON.stringify({ sourceType: "mmc" }) }),
  env: { ROSTER_AUTOMATION_TOKEN: "token", ROSTER_DB: { prepare() { databaseCalls += 1; throw new Error("must not run"); } } },
});
assert.equal(pausedResponse.status, 503);
assert.equal(databaseCalls, 0, "disabled materialisation must stop before D1");

const invalidTermResponse = await materialize({
  request: new Request("http://local/api/automation/facility-materialize", { method: "POST", headers: { authorization: "Bearer token", "content-type": "application/json" }, body: JSON.stringify({ sourceType: "mmc" }) }),
  env: { ...env, ROSTER_AUTOMATION_TOKEN: "token", ROSTER_DB: { prepare() { databaseCalls += 1; throw new Error("must not run"); } }, ROSTER_FILES: { put() {} } },
});
assert.equal(invalidTermResponse.status, 400);
assert.equal(databaseCalls, 0, "invalid or missing term must stop before D1");
console.log("Facility rollout cohort, source allowlist, emergency pause and pre-D1 guards passed.");
