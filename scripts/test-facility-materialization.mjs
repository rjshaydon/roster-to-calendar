import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile, readdir } from "node:fs/promises";
import { DatabaseSync } from "node:sqlite";
import {
  deleteDerivedRosterFile,
  australianTermEndForStart,
  queryMaterializedFacilityCoverage,
  queryMaterializedFacilityTermStaff,
  queryRosterFileStatusSummaries,
  refreshFacilityOverviewMaterializationForFile,
  replaceDerivedRosterFile,
  startDerivedRosterFileSave,
  appendDerivedRosterFileEvents,
  activateDerivedRosterFile,
  FACILITY_BOOTSTRAP_EVENT_SQL,
} from "../functions/_lib/d1-calendar.js";
import { onRequestPost as saveAutomatedDerivedRoster } from "../functions/api/automation/derived.js";
import { onRequestPost as bootstrapFacility } from "../functions/api/automation/facility-bootstrap.js";
import { onRequestPost as materializeFacility } from "../functions/api/automation/facility-materialize.js";
import { onRequestPost as ingestContacts } from "../functions/api/automation/contact-list-extract.js";
import { onRequestPost as stateHandler } from "../functions/api/state.js";
import { initializeFacilityMaterialization, loadPublishedFacilityDays, loadPublishedFacilityMetadata, loadPublishedFacilityRange, loadPublishedFacilityStaff, publishFacilityDays, publishFacilityStaffMetadata } from "../functions/_lib/facility-overview-cache.js";

class LocalD1 {
  constructor(sqlite) { this.sqlite = sqlite; this.rowsWritten = 0; this.sql = []; this.failRunIncludes = ""; }
  prepare(sql) {
    const owner = this;
    owner.sql.push(sql.replace(/\s+/g, " ").trim());
    return {
      args: [],
      bind(...args) { this.args = args; return this; },
      async run() {
        if (owner.failRunIncludes && sql.includes(owner.failRunIncludes)) {
          owner.failRunIncludes = "";
          throw new Error("Injected D1 statement failure");
        }
        const result = owner.sqlite.prepare(sql).run(...this.args);
        owner.rowsWritten += Number(result.changes || 0);
        return { success: true, meta: { changes: Number(result.changes || 0) } };
      },
      async all() { return { success: true, results: owner.sqlite.prepare(sql).all(...this.args) }; },
      async first() { return owner.sqlite.prepare(sql).get(...this.args) || null; },
    };
  }
  async batch(statements) {
    const results = [];
    const ownsTransaction = !this.sqlite.isTransaction;
    if (ownsTransaction) this.sqlite.exec("BEGIN");
    try {
      for (const statement of statements) results.push(await statement.run());
      if (ownsTransaction) this.sqlite.exec("COMMIT");
      return results;
    } catch (error) {
      if (ownsTransaction) this.sqlite.exec("ROLLBACK");
      throw error;
    }
  }
}

class LocalR2 {
  constructor() { this.objects = new Map(); this.puts = 0; this.gets = 0; this.version = 0; this.failPointerOnce = false; }
  async put(key, value, options = {}) {
    if (this.failPointerOnce && key.endsWith("/manifest.json")) { this.failPointerOnce = false; throw new Error("Injected manifest failure"); }
    const current = this.objects.get(key);
    if (options.onlyIf?.etagMatches && current?.etag !== options.onlyIf.etagMatches) throw new Error("Precondition failed");
    this.puts += 1;
    const bytes = value instanceof Uint8Array ? value : new Uint8Array(value);
    this.version += 1;
    this.objects.set(key, { bytes, options, etag: `local-${this.version}` });
  }
  async get(key) {
    this.gets += 1;
    const item = this.objects.get(key);
    if (!item) return null;
    return { etag: item.etag, arrayBuffer: async () => item.bytes.buffer.slice(item.bytes.byteOffset, item.bytes.byteOffset + item.bytes.byteLength) };
  }
}

const sqlite = new DatabaseSync(":memory:");
assert.equal(australianTermEndForStart("2026-02-02"), "2026-05-03");
assert.equal(australianTermEndForStart("2026-08-03"), "2026-11-01");
for (const name of (await readdir(new URL("../migrations", import.meta.url))).filter((name) => name.endsWith(".sql")).sort()) {
  sqlite.exec(await readFile(new URL(`../migrations/${name}`, import.meta.url), "utf8"));
}
const db = new LocalD1(sqlite);
const bootstrapEventPlan = sqlite.prepare(`EXPLAIN QUERY PLAN ${FACILITY_BOOTSTRAP_EVENT_SQL}`).all("fixture-mmc", 25001).map((row) => String(row.detail));
assert.ok(bootstrapEventPlan.some((line) => /SEARCH roster_events USING INDEX idx_roster_events_file \(file_id=\?\)/.test(line)), "bootstrap must use the exact file index");
assert.equal(bootstrapEventPlan.some((line) => /SCAN roster_events|TEMP B-TREE/i.test(line)), false, "bootstrap must not scan or sort roster history before applying its limit");
const bootstrapStaffPlan = sqlite.prepare("EXPLAIN QUERY PLAN SELECT source_type, term_start, doctor_key, file_id, fact_digest FROM facility_term_staff_contributions WHERE file_id = ? LIMIT ?").all("fixture-mmc", 751).map((row) => String(row.detail));
const bootstrapCatalogPlan = sqlite.prepare("EXPLAIN QUERY PLAN SELECT source_type, term_start, file_id, catalog_key, fact_digest FROM facility_stream_catalog_contributions WHERE file_id = ? LIMIT ?").all("fixture-mmc", 751).map((row) => String(row.detail));
assert.ok(bootstrapStaffPlan.some((line) => /idx_facility_term_staff_file/.test(line)), "bootstrap Staff lookup must use its leading file index");
assert.ok(bootstrapCatalogPlan.some((line) => /idx_facility_stream_catalog_file/.test(line)), "bootstrap catalogue lookup must use its leading file index");
const publicationStaffPlan = sqlite.prepare("EXPLAIN QUERY PLAN SELECT * FROM facility_term_staff_contributions WHERE source_type = ? AND term_start = ? LIMIT ?").all("mmc", "2026-08-03", 24001).map((row) => String(row.detail));
const publicationCatalogPlan = sqlite.prepare("EXPLAIN QUERY PLAN SELECT * FROM facility_stream_catalog_contributions WHERE source_type = ? AND term_start = ? LIMIT ?").all("mmc", "2026-08-03", 24001).map((row) => String(row.detail));
const publicationFilesPlan = sqlite.prepare("EXPLAIN QUERY PLAN SELECT f.id FROM roster_files AS f INDEXED BY idx_roster_files_source_active WHERE f.source_type = ? AND f.active = 1 LIMIT ?").all("mmc", 33).map((row) => String(row.detail));
const statusPlan = sqlite.prepare("EXPLAIN QUERY PLAN SELECT file_id FROM roster_file_status_summaries WHERE active = 1 ORDER BY source_type, updated_at, file_id LIMIT ?").all(100).map((row) => String(row.detail));
assert.ok(publicationStaffPlan.some((line) => /idx_facility_term_staff_lookup/.test(line)), "publication Staff inputs must use the exact ED/term index");
assert.ok(publicationCatalogPlan.some((line) => /idx_facility_stream_catalog_lookup/.test(line)), "publication catalogue inputs must use the exact ED/term index");
assert.ok(publicationFilesPlan.some((line) => /idx_roster_files_source_active/.test(line)), "publication coverage must begin with the capped active-file index");
assert.ok(statusPlan.some((line) => /idx_roster_file_status_active_source_updated/.test(line)), "roster status must use the compact active-summary index");
const file = { id: "incremental-a", name: "Synthetic.xlsx", sourceType: "mmc", sourceId: "test", active: true, parserVersion: "test-v1" };
const doctors = [
  { key: "PERMANENT SMS", displayName: "Permanent SMS", seniority: "SMS", membershipSource: "roster" },
  { key: "TERM TRAINEE", displayName: "Term Trainee", seniority: "Registrar", membershipSource: "roster" },
];
const event = (id, date, title = "Day") => ({ id, source: "MMC", start: `${date}T08:00:00`, end: `${date}T16:00:00`, title, rawValue: title, seniority: id.startsWith("sms") ? "SMS" : "Registrar" });
const initialEvents = {
  "PERMANENT SMS": [event("sms-1", "2026-08-03")],
  "TERM TRAINEE": [event("trainee-1", "2026-08-03")],
};

const chunkCrashFile = { ...file, id: "chunk-crash", active: false };
await startDerivedRosterFileSave(db, chunkCrashFile, doctors);
db.failRunIncludes = "UPDATE roster_file_status_summaries";
await assert.rejects(appendDerivedRosterFileEvents(db, chunkCrashFile, doctors, initialEvents), /Injected D1 statement failure/);
assert.equal(sqlite.prepare("SELECT COUNT(*) AS count FROM roster_events WHERE file_id = ?").get(chunkCrashFile.id).count, 0, "a failed chunk must roll back its event inserts");
assert.equal(sqlite.prepare("SELECT derived_state FROM roster_file_status_summaries WHERE file_id = ?").get(chunkCrashFile.id).derived_state, "building", "a failed chunk must not publish ready status");
await deleteDerivedRosterFile(db, chunkCrashFile.id);

const chunkedFile = { ...file, id: "chunked-ready" };
await startDerivedRosterFileSave(db, chunkedFile, doctors);
await appendDerivedRosterFileEvents(db, chunkedFile, doctors, initialEvents, {}, { indexedEventCount: 2 });
const firstChunkRevision = sqlite.prepare("SELECT status_revision FROM roster_file_status_summaries WHERE file_id = ?").get(chunkedFile.id).status_revision;
await appendDerivedRosterFileEvents(db, chunkedFile, doctors, initialEvents, {}, { indexedEventCount: 2 });
const repeatedChunk = sqlite.prepare("SELECT event_count, status_revision FROM roster_file_status_summaries WHERE file_id = ?").get(chunkedFile.id);
assert.equal(repeatedChunk.event_count, 2, "a retried chunk must not double-count events");
assert.equal(repeatedChunk.status_revision, firstChunkRevision, "a retried chunk must not change the status revision");
await activateDerivedRosterFile(db, chunkedFile.id);
const chunkedReady = sqlite.prepare("SELECT active, derived_state FROM roster_file_status_summaries WHERE file_id = ?").get(chunkedFile.id);
assert.equal(chunkedReady.active, 1, "a complete chunked save must activate its summary");
assert.equal(chunkedReady.derived_state, "ready", "a complete chunked save must publish ready only after activation");
await deleteDerivedRosterFile(db, chunkedFile.id);

const publicationCrashFile = { ...file, id: "publication-crash" };
await startDerivedRosterFileSave(db, publicationCrashFile, doctors);
await appendDerivedRosterFileEvents(db, publicationCrashFile, doctors, initialEvents, {}, { indexedEventCount: 2 });
db.failRunIncludes = "SET active = 1, derived_state = 'ready'";
await assert.rejects(activateDerivedRosterFile(db, publicationCrashFile.id), /Injected D1 statement failure/);
assert.equal(sqlite.prepare("SELECT derived_state FROM roster_file_status_summaries WHERE file_id = ?").get(publicationCrashFile.id).derived_state, "building", "a failed final publication must not expose ready status");
await deleteDerivedRosterFile(db, publicationCrashFile.id);

const materializationCrashFile = { ...file, id: "materialization-crash" };
db.failRunIncludes = "INSERT INTO roster_file_coverage";
await assert.rejects(replaceDerivedRosterFile(db, materializationCrashFile, doctors, initialEvents), /Injected D1 statement failure/);
assert.equal(sqlite.prepare("SELECT derived_state FROM roster_file_status_summaries WHERE file_id = ?").get(materializationCrashFile.id).derived_state, "building", "a failed materialisation must not leave a ready summary");
await deleteDerivedRosterFile(db, materializationCrashFile.id);

const contactR2 = new LocalR2();
const contactPayload = { sourceId: "mmc-shift-allocations", sourceDate: "2026-09-06", providerModifiedAt: "2026-09-06T01:00:00Z", contacts: [{ area: "Adult Emergency", shift: "AM", role: "Consultant", name: "Alex Example", phone: "555-0100", isPopulated: true }] };
async function callContactExtract() {
  return ingestContacts({
    request: new Request("http://local/api/automation/contact-list-extract", { method: "POST", headers: { authorization: "Bearer contact-token", "content-type": "application/json" }, body: JSON.stringify(contactPayload) }),
    env: { ROSTER_DB: db, ROSTER_FILES: contactR2, ROSTER_AUTOMATION_TOKEN: "contact-token", CONTACT_AUTOMATION_WRITES_ENABLED: "true", CONTACT_AUTOMATION_SOURCE_ALLOWLIST: "mmc-shift-allocations" },
  });
}
const firstContact = await callContactExtract();
assert.equal(firstContact.status, 200);
assert.equal((await firstContact.json()).status, "stored");
db.rowsWritten = 0;
const contactPuts = contactR2.puts;
const repeatedContact = await callContactExtract();
assert.equal((await repeatedContact.json()).status, "unchanged");
assert.equal(db.rowsWritten, 0, "an unchanged allowed contact extract must write no D1 rows");
assert.equal(contactR2.puts, contactPuts, "an unchanged allowed contact extract must write no R2 objects");

async function callBootstrap(body) {
  const response = await bootstrapFacility({
    request: new Request("http://local/api/automation/facility-bootstrap", { method: "POST", headers: { authorization: "Bearer bootstrap-token", "content-type": "application/json" }, body: JSON.stringify(body) }),
    env: { ROSTER_DB: db, ROSTER_AUTOMATION_TOKEN: "bootstrap-token", ROSTER_ADVANCED_MAINTENANCE_ENABLED: "true", FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST: "mmc" },
  });
  return { response, payload: await response.json() };
}

const legacyFile = { ...file, id: "bootstrap-file", name: "Legacy.xlsx" };
await startDerivedRosterFileSave(db, legacyFile, doctors);
await appendDerivedRosterFileEvents(db, legacyFile, doctors, initialEvents);
sqlite.prepare("UPDATE roster_files SET active = 1 WHERE id = ?").run(legacyFile.id);
assert.equal(sqlite.prepare("SELECT COUNT(*) AS count FROM roster_file_coverage WHERE file_id = 'bootstrap-file'").get().count, 0);
const bootstrapPlan = await callBootstrap({ sourceType: "mmc", fileId: legacyFile.id, maximumEventRows: 10, maximumWrites: 100 });
assert.equal(bootstrapPlan.response.status, 200);
assert.equal(bootstrapPlan.payload.dryRun, true);
const staleBootstrap = await callBootstrap({ sourceType: "mmc", fileId: legacyFile.id, maximumEventRows: 10, maximumWrites: 100, execute: true, planRevision: "stale" });
assert.equal(staleBootstrap.response.status, 409);
db.rowsWritten = 0;
const writeLimitedBootstrap = await callBootstrap({ sourceType: "mmc", fileId: legacyFile.id, maximumEventRows: 10, maximumWrites: 1, execute: true, planRevision: bootstrapPlan.payload.planRevision });
assert.equal(writeLimitedBootstrap.response.status, 409);
assert.equal(writeLimitedBootstrap.payload.result.reason, "compact-write-limit");
assert.equal(db.rowsWritten, 0, "a compact-write overage must stop before writes");
db.rowsWritten = 0;
const bootstrapExecution = await callBootstrap({ sourceType: "mmc", fileId: legacyFile.id, maximumEventRows: 10, maximumWrites: 100, execute: true, planRevision: bootstrapPlan.payload.planRevision });
assert.equal(bootstrapExecution.response.status, 200);
assert.equal(bootstrapExecution.payload.result.overBudget, undefined);
assert.ok(db.rowsWritten <= 100);
for (const [label, limits, reason] of [
  ["doctor", { maximumDoctorRows: 1 }, "doctor-read-limit"],
  ["existing Staff", { maximumExistingStaffRows: 1 }, "existing-staff-read-limit"],
  ["existing catalogue", { maximumExistingCatalogRows: 1 }, "existing-catalog-read-limit"],
]) {
  db.rowsWritten = 0;
  const bounded = await refreshFacilityOverviewMaterializationForFile(db, legacyFile.id, { sourceType: "mmc", maximumEventRows: 10, maximumWrites: 100, ...limits });
  assert.equal(bounded.reason, reason, `${label} sentinel must reject the file`);
  assert.equal(db.rowsWritten, 0, `${label} sentinel must stop before writes`);
}
db.rowsWritten = 0;
const repeatedBootstrap = await callBootstrap({ sourceType: "mmc", fileId: legacyFile.id, maximumEventRows: 10, maximumWrites: 100, execute: true, planRevision: bootstrapExecution.payload.planRevision });
assert.equal(repeatedBootstrap.payload.unchanged, true);
assert.equal(db.rowsWritten, 0, "a repeated compact bootstrap must write nothing");

const oversizedFile = { ...file, id: "bootstrap-oversized", name: "Oversized.xlsx" };
await startDerivedRosterFileSave(db, oversizedFile, doctors);
await appendDerivedRosterFileEvents(db, oversizedFile, doctors, initialEvents);
sqlite.prepare("UPDATE roster_files SET active = 1 WHERE id = ?").run(oversizedFile.id);
const oversizedPlan = await callBootstrap({ sourceType: "mmc", fileId: oversizedFile.id, maximumEventRows: 1, maximumWrites: 100 });
db.rowsWritten = 0;
const oversizedExecution = await callBootstrap({ sourceType: "mmc", fileId: oversizedFile.id, maximumEventRows: 1, maximumWrites: 100, execute: true, planRevision: oversizedPlan.payload.planRevision });
assert.equal(oversizedExecution.response.status, 409);
assert.equal(oversizedExecution.payload.result.reason, "event-read-limit");
assert.equal(db.rowsWritten, 0, "an over-budget bootstrap must stop before compact writes");
await deleteDerivedRosterFile(db, legacyFile.id);
await deleteDerivedRosterFile(db, oversizedFile.id);

const first = await replaceDerivedRosterFile(db, file, doctors, initialEvents);
assert.equal(first.unchanged, false);
assert.equal((await queryMaterializedFacilityCoverage(db, { sourceType: "mmc" }))[0].startDate, "2026-08-03");
assert.equal((await queryMaterializedFacilityTermStaff(db, { sourceType: "mmc", termStart: "2026-08-03" })).length, 2);
assert.equal(sqlite.prepare("SELECT visible_from FROM facility_term_visibility WHERE source_type='mmc' AND term_start='2026-08-03'").get().visible_from, "2026-07-20");

db.rowsWritten = 0;
const unchanged = await replaceDerivedRosterFile(db, file, doctors, initialEvents);
assert.equal(unchanged.unchanged, true);
assert.equal(db.rowsWritten, 0, "identical import must perform zero writes");

db.rowsWritten = 0;
const correctedEvents = { ...initialEvents, "TERM TRAINEE": [event("trainee-1", "2026-08-03", "Sick leave")] };
const corrected = await replaceDerivedRosterFile(db, file, doctors, correctedEvents);
assert.equal(corrected.changes.events, 1, "one correction must change one event fact");
assert.ok(db.rowsWritten <= 9, `one crash-safe correction wrote ${db.rowsWritten} rows`);
const correctedStatus = (await queryRosterFileStatusSummaries(db, { expectedFileIds: [file.id] }))[0];
assert.equal(correctedStatus.eventCount, 2);
assert.equal(correctedStatus.indexedDoctors, 2);
assert.equal(correctedStatus.derivedState, "ready");
assert.ok(correctedStatus.statusRevision.length <= 36, "status revisions must remain fixed-size");

db.rowsWritten = 0;
await assert.rejects(
  replaceDerivedRosterFile(db, file, doctors, {
    "PERMANENT SMS": [event("sms-1", "2026-08-03", "Changed SMS shift")],
    "TERM TRAINEE": [event("trainee-1", "2026-08-03", "Changed trainee shift")],
  }, {}, { maximumIncrementalFacts: 1 }),
  (error) => error?.code === "ROSTER_INCREMENTAL_BUDGET" && error.changedFactCount === 2,
);
assert.equal(db.rowsWritten, 0, "an over-budget automatic revision must stop before writes");

sqlite.prepare("INSERT INTO facility_term_visibility (source_type, term_start, visible_from, revision, updated_at) VALUES ('mmc', '2026-02-02', '2026-01-19', '', '')").run();
const plannedR2 = new LocalR2();
db.sql = [];
const materializeEnv = { ROSTER_DB: db, ROSTER_FILES: plannedR2, ROSTER_AUTOMATION_TOKEN: "bootstrap-token", ROSTER_ADVANCED_MAINTENANCE_ENABLED: "true", FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST: "mmc" };
const materializationPlanResponse = await materializeFacility({ request: new Request("http://local/api/automation/facility-materialize", { method: "POST", headers: { authorization: "Bearer bootstrap-token", "content-type": "application/json" }, body: JSON.stringify({ sourceType: "mmc", termStart: "2026-08-03" }) }), env: materializeEnv });
const materializationPlan = await materializationPlanResponse.json();
assert.equal(materializationPlan.dryRun, true);
assert.ok(materializationPlan.planRevision);
const stalePlan = await initializeFacilityMaterialization({ env: { ROSTER_DB: db, ROSTER_FILES: plannedR2 } }, "mmc", { termStart: "2026-08-03", dryRun: false, planRevision: "stale" });
assert.equal(stalePlan.stalePlan, true);
assert.equal(plannedR2.puts, 0, "a stale publication plan must write nothing");
db.sql = [];
const plannedPutsBefore = plannedR2.puts;
const plannedGetsBefore = plannedR2.gets;
const materializationExecution = await initializeFacilityMaterialization({ env: { ROSTER_DB: db, ROSTER_FILES: plannedR2 } }, "mmc", { termStart: "2026-08-03", dryRun: false, planRevision: materializationPlan.planRevision });
assert.equal(materializationExecution.ok, true);
assert.ok(plannedR2.puts - plannedPutsBefore <= materializationPlan.estimate.maximumR2Puts);
assert.ok(plannedR2.gets - plannedGetsBefore <= materializationPlan.estimate.maximumR2Gets);
assert.ok(db.sql.length <= materializationPlan.estimate.maximumD1ReadStatements + materializationPlan.estimate.maximumPublicationStateStatements);
const plannedStaffQueries = db.sql.filter((sql) => /FROM facility_term_staff_contributions/.test(sql));
assert.ok(plannedStaffQueries.length >= 1 && plannedStaffQueries.every((sql) => /(?:s\.)?term_start = \?/.test(sql)), "every planned Staff query must be constrained to the requested term");
const plannedCatalogQueries = db.sql.filter((sql) => /FROM facility_stream_catalog_contributions/.test(sql));
assert.ok(plannedCatalogQueries.length >= 1 && plannedCatalogQueries.every((sql) => /(?:c\.)?term_start = \?/.test(sql)), "every planned catalogue query must be constrained to the requested term");

async function assertPublicationPlanInvalidated(label, mutate, restore) {
  const dryRun = await initializeFacilityMaterialization({ env: { ROSTER_DB: db, ROSTER_FILES: plannedR2 } }, "mmc", { termStart: "2026-08-03" });
  await mutate();
  db.rowsWritten = 0;
  const putsBefore = plannedR2.puts;
  const result = await initializeFacilityMaterialization({ env: { ROSTER_DB: db, ROSTER_FILES: plannedR2 } }, "mmc", {
    termStart: "2026-08-03", dryRun: false, planRevision: dryRun.planRevision,
  });
  assert.equal(result.stalePlan, true, `${label} must invalidate the approved publication plan`);
  assert.equal(db.rowsWritten, 0, `${label} must be rejected before D1 writes`);
  assert.equal(plannedR2.puts, putsBefore, `${label} must be rejected before R2 writes`);
  await restore();
}

const baselineCoverage = sqlite.prepare("SELECT content_revision, daily_digest FROM roster_file_coverage WHERE file_id=?").get(file.id);
const baselineEventJson = sqlite.prepare("SELECT event_json FROM roster_events WHERE file_id=? AND doctor_key='TERM TRAINEE'").get(file.id).event_json;
await assertPublicationPlanInvalidated("daily roster content", async () => {
  sqlite.prepare("UPDATE roster_events SET event_json=? WHERE file_id=? AND doctor_key='TERM TRAINEE'").run(JSON.stringify({ ...JSON.parse(baselineEventJson), rawValue: "Changed without changing the catalogue signature" }), file.id);
  sqlite.prepare("UPDATE roster_file_coverage SET daily_digest='changed-daily' WHERE file_id=?").run(file.id);
}, async () => {
  sqlite.prepare("UPDATE roster_events SET event_json=? WHERE file_id=? AND doctor_key='TERM TRAINEE'").run(baselineEventJson, file.id);
  sqlite.prepare("UPDATE roster_file_coverage SET daily_digest=? WHERE file_id=?").run(baselineCoverage.daily_digest, file.id);
});
await assertPublicationPlanInvalidated("Staff grade", async () => {
  sqlite.prepare("UPDATE facility_term_staff_contributions SET seniority='HMO' WHERE file_id=? AND doctor_key='TERM TRAINEE'").run(file.id);
}, async () => {
  sqlite.prepare("UPDATE facility_term_staff_contributions SET seniority='Registrar' WHERE file_id=? AND doctor_key='TERM TRAINEE'").run(file.id);
});
await assertPublicationPlanInvalidated("SMS continuity", async () => {
  sqlite.prepare(`INSERT INTO facility_sms_memberships
    (source_type, doctor_key, display_name, first_seen_date, last_seen_date)
    VALUES ('mmc', 'SMS ON LEAVE', 'SMS On Leave', '2026-01-01', '2026-05-01')`).run();
}, async () => {
  sqlite.prepare("DELETE FROM facility_sms_memberships WHERE source_type='mmc' AND doctor_key='SMS ON LEAVE'").run();
});
await assertPublicationPlanInvalidated("Staff designation", async () => {
  sqlite.prepare(`INSERT INTO facility_staff_designations
    (id, source_type, doctor_key, display_name, seniority, designation, term_start, term_end, active)
    VALUES ('test-designation', 'mmc', 'PERMANENT SMS', 'Permanent SMS', 'SMS', 'sabbatical_leave', '2026-08-03', '2026-11-01', 1)`).run();
}, async () => {
  sqlite.prepare("DELETE FROM facility_staff_designations WHERE id='test-designation'").run();
});
await assertPublicationPlanInvalidated("seniority override", async () => {
  sqlite.prepare(`INSERT INTO facility_staff_seniority_overrides
    (id, source_type, doctor_key, display_name, seniority, term_start, active)
    VALUES ('test-override', 'mmc', 'TERM TRAINEE', 'Term Trainee', 'HMO', '2026-08-03', 1)`).run();
}, async () => {
  sqlite.prepare("DELETE FROM facility_staff_seniority_overrides WHERE id='test-override'").run();
});
await assertPublicationPlanInvalidated("term visibility", async () => {
  sqlite.prepare("UPDATE facility_term_visibility SET visible_from='2026-07-21' WHERE source_type='mmc' AND term_start='2026-08-03'").run();
}, async () => {
  sqlite.prepare("UPDATE facility_term_visibility SET visible_from='2026-07-20' WHERE source_type='mmc' AND term_start='2026-08-03'").run();
});
await assertPublicationPlanInvalidated("compact coverage", async () => {
  sqlite.prepare("UPDATE roster_file_coverage SET content_revision='changed-content' WHERE file_id=?").run(file.id);
}, async () => {
  sqlite.prepare("UPDATE roster_file_coverage SET content_revision=? WHERE file_id=?").run(baselineCoverage.content_revision, file.id);
});
const manifestKey = "facility-overview/v1/mmc/manifest.json";
const originalManifestObject = plannedR2.objects.get(manifestKey);
await assertPublicationPlanInvalidated("manifest ETag", async () => {
  plannedR2.objects.set(manifestKey, { ...originalManifestObject, etag: "changed-etag" });
}, async () => {
  plannedR2.objects.set(manifestKey, originalManifestObject);
});

sqlite.prepare(`INSERT INTO facility_staff_designations
  (id, source_type, doctor_key, display_name, seniority, designation, term_start, term_end, active)
  VALUES ('concurrent-designation', 'mmc', 'PERMANENT SMS', 'Permanent SMS', 'SMS', 'sabbatical_leave', '2026-08-03', '2026-11-01', 1)`).run();
const concurrentPlan = await initializeFacilityMaterialization({ env: { ROSTER_DB: db, ROSTER_FILES: plannedR2 } }, "mmc", { termStart: "2026-08-03" });
const fixedManifestBeforeConcurrentBuild = plannedR2.objects.get(manifestKey);
const concurrentResult = await initializeFacilityMaterialization({ env: { ROSTER_DB: db, ROSTER_FILES: plannedR2 } }, "mmc", {
  termStart: "2026-08-03", dryRun: false, planRevision: concurrentPlan.planRevision,
  beforeFinalValidation: async () => sqlite.prepare("UPDATE roster_file_coverage SET daily_digest='changed-during-build' WHERE file_id=?").run(file.id),
});
assert.equal(concurrentResult.stalePlan, true, "an input change during publication must abort the fixed pointer update");
assert.equal(plannedR2.objects.get(manifestKey).etag, fixedManifestBeforeConcurrentBuild.etag, "a stale concurrent build must preserve the fixed manifest");
sqlite.prepare("UPDATE roster_file_coverage SET daily_digest=? WHERE file_id=?").run(baselineCoverage.daily_digest, file.id);
sqlite.prepare("DELETE FROM facility_staff_designations WHERE id='concurrent-designation'").run();

const r2 = new LocalR2();
db.sql = [];
const publication = await publishFacilityStaffMetadata({ env: { ROSTER_DB: db, ROSTER_FILES: r2 } }, ["mmc"]);
assert.equal(publication.ok, true);
assert.ok(r2.puts >= 2, "first publication should write immutable staff and a manifest");
assert.equal(db.sql.some((sql) => /\broster_events\b/i.test(sql)), false, "Staff/metadata publication must use compact facts only");
const firstPutCount = r2.puts;
await publishFacilityStaffMetadata({ env: { ROSTER_DB: db, ROSTER_FILES: r2 } }, ["mmc"]);
assert.equal(r2.puts, firstPutCount, "unchanged Staff/metadata publication must write no R2 objects");
db.sql = [];
const publishedMetadata = await loadPublishedFacilityMetadata(r2, ["mmc"], "2026-08-03");
const publishedStaff = await loadPublishedFacilityStaff(r2, ["mmc"], "2026-08-03", "2026-08-03");
assert.equal(publishedMetadata.preparing, false);
assert.ok(publishedMetadata.catalogEvents.length > 0, "published metadata must retain the stream catalogue");
assert.equal(publishedStaff.preparing, false);
assert.equal(publishedStaff.members.length, 2);
assert.equal(db.sql.length, 0, "shared Staff/metadata readers must perform zero D1 queries");
assert.equal((await loadPublishedFacilityStaff(new LocalR2(), ["mmc"], "2026-08-03", "2026-08-03")).preparing, true, "a missing object must return preparing without a fallback");
const firstDayPublication = await publishFacilityDays({ env: { ROSTER_DB: db, ROSTER_FILES: r2 } }, "mmc", ["2026-08-03"]);
assert.equal(firstDayPublication.ok, true);
db.sql = [];
const publishedDay = await loadPublishedFacilityDays(r2, ["mmc"], "2026-08-03");
assert.equal(publishedDay.preparing, false);
assert.equal(publishedDay.rows.length, 2);
assert.equal((await loadPublishedFacilityDays(r2, ["mmc"], "2026-08-03", "2026-07-19")).preparing, true, "a known future day must remain unavailable 15 days before term start");
assert.equal((await loadPublishedFacilityDays(r2, ["mmc"], "2026-08-03", "2026-07-20")).preparing, false, "a known future day must become available at the 14-day boundary");
assert.equal(db.sql.length, 0, "shared day readers must perform zero D1 queries");
const publishedRange = await loadPublishedFacilityRange(r2, ["mmc"], "2026-08-01", "2026-08-31", "2026-08-03");
assert.equal(publishedRange.preparing, false);
assert.equal(publishedRange.events.length, 2, "the monthly range snapshot must contain the same published day rows");
assert.equal(db.sql.length, 0, "shared range readers must perform zero D1 queries");
const rangeReadsBeforeRevalidation = r2.gets;
const unchangedPublishedRange = await loadPublishedFacilityRange(r2, ["mmc"], "2026-08-01", "2026-08-31", "2026-08-03", { cachedRevision: publishedRange.revision });
assert.equal(unchangedPublishedRange.unchanged, true);
assert.equal(r2.gets - rangeReadsBeforeRevalidation, 1, "unchanged range revalidation must read only the ED manifest");
const dayPutCount = r2.puts;
const repeatedDayPublication = await publishFacilityDays({ env: { ROSTER_DB: db, ROSTER_FILES: r2 } }, "mmc", ["2026-08-03"]);
assert.equal(repeatedDayPublication.unchanged, true);
assert.equal(r2.puts, dayPutCount, "unchanged day publication must write no R2 objects");

const password = "local-password";
const salt = "local-salt";
const passwordHash = createHash("sha256").update(`${salt}:${password}`).digest("hex");
sqlite.prepare(`INSERT INTO account_profiles
  (email, real_name, role, facility_overview_enabled, password_salt, password_hash, created_at, updated_at)
  VALUES (?, 'Term Trainee', 'user', 1, ?, ?, ?, ?)`)
  .run("doctor@example.com", salt, passwordHash, new Date().toISOString(), new Date().toISOString());
sqlite.prepare(`INSERT INTO account_claims (email, source_type, doctor_key, display_name, matched_at, updated_at)
  VALUES ('doctor@example.com', 'mmc', 'TERM TRAINEE', 'Term Trainee', '', '')`).run();
async function callSharedAction(body, options = {}) {
  db.sql = [];
  const response = await stateHandler({
    request: new Request("http://127.0.0.1/api/state", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ email: "doctor@example.com", password, ...body }) }),
    env: { ROSTER_DB: db, ROSTER_FILES: options.r2 || r2, FACILITY_SHARED_ROLLOUT_ACTIVE: "true", FACILITY_ACCESS_MATERIALIZATION_ENABLED: "true", FACILITY_SHARED_METADATA_ENABLED: "true", FACILITY_SHARED_DAYS_ENABLED: "true", FACILITY_SHARED_READER_SOURCE_ALLOWLIST: "mmc", FACILITY_SHARED_READER_COHORT: "all" },
    waitUntil() {},
  });
  const payload = await response.json();
  assert.equal(response.status, options.status || 200, JSON.stringify(payload));
  assert.equal(db.sql.some((sql) => /\broster_events\b/i.test(sql)), false, `${body.action} must not query roster_events`);
  return payload;
}
const handlerMetadata = await callSharedAction({ action: "queryFacilityOverviewMetadata", sourceTypes: ["mmc"] });
assert.ok(handlerMetadata.catalogEvents.length > 0);
const handlerStaff = await callSharedAction({ action: "queryFacilityOverviewStaff", facilityKey: "mmc", termStart: "2026-08-03", termEnd: "2026-11-02" });
assert.equal(handlerStaff.members.length, 2);
const handlerDay = await callSharedAction({ action: "queryFacilityOverviewOnShift", facilityKey: "mmc", date: "2026-08-03", includeClinicalSupport: true });
assert.equal(handlerDay.events.length, 1, "On shift handler must filter the shared day object using existing working-shift rules");
const unchangedHandlerDay = await callSharedAction({ action: "queryFacilityOverviewOnShift", facilityKey: "mmc", date: "2026-08-03", includeClinicalSupport: true, cachedRevision: handlerDay.revision });
assert.equal(unchangedHandlerDay.rosterUnchanged, true, "an unchanged On shift revision must not retransmit roster events");
assert.equal(unchangedHandlerDay.events, undefined);
const handlerRange = await callSharedAction({ action: "queryFacilityOverviewByStream", startDate: "2026-08-01", endDate: "2026-08-31", selections: [{ id: "day", facilityKey: "mmc", streamKey: "day", seniority: "ALL" }] });
assert.equal(handlerRange.events.length, 1, "By stream must use the shared monthly object and existing working-shift filtering");
assert.equal(db.sql.some((sql) => /roster_events|roster_daily_presence/i.test(sql)), false, "By stream shared reads must not query roster history");
assert.equal((await callSharedAction({ action: "queryFacilityOverviewByStream", startDate: "2026-08-01", endDate: "2026-08-31", selections: [{ id: "day", facilityKey: "mmc", streamKey: "day", seniority: "ALL" }], cachedRevision: handlerRange.revision })).unchanged, true);
const handlerTogether = await callSharedAction({ action: "queryFacilityOverviewWorkingTogether", startDate: "2026-08-01", endDate: "2026-08-31", sourceTypes: ["mmc"], doctorKeys: ["PERMANENT SMS"] });
assert.equal(handlerTogether.events.length, 1, "Working together must filter the shared monthly object by doctor");
assert.equal(db.sql.some((sql) => /roster_events|roster_daily_presence/i.test(sql)), false, "Working together shared reads must not query roster history");
assert.equal((await callSharedAction({ action: "queryFacilityOverviewWorkingTogether", startDate: "2026-08-01", endDate: "2026-08-31", sourceTypes: ["mmc"], doctorKeys: ["PERMANENT SMS"], cachedRevision: handlerTogether.revision })).unchanged, true);
const missingHandlerDay = await callSharedAction(
  { action: "queryFacilityOverviewOnShift", facilityKey: "mmc", date: "2026-08-03", includeClinicalSupport: true },
  { r2: new LocalR2(), status: 503 },
);
assert.equal(missingHandlerDay.preparing, true, "an On shift miss must not build or query roster events");
const blockedDisallowedDay = await callSharedAction(
  { action: "queryFacilityOverviewOnShift", facilityKey: "ddh", date: "2026-08-03", includeClinicalSupport: true },
  { status: 403 },
);
assert.match(blockedDisallowedDay.error, /not available/i, "a disallowed canary ED must be blocked without a legacy query");
const forbiddenAllStaff = await callSharedAction(
  { action: "queryFacilityOverviewStaff", facilityKey: "all", termStart: "2026-08-03", termEnd: "2026-11-02" },
  { status: 403 },
);
assert.match(forbiddenAllStaff.error, /not available/i, "a site-scoped account must not read All EDs Staff data");
const missingHandlerStaff = await callSharedAction(
  { action: "queryFacilityOverviewStaff", facilityKey: "mmc", termStart: "2026-08-03", termEnd: "2026-11-02" },
  { r2: new LocalR2(), status: 503 },
);
assert.equal(missingHandlerStaff.preparing, true, "a handler cache miss must not run the legacy Staff query");

const lateCorrectionEvents = { ...initialEvents, "TERM TRAINEE": [event("trainee-1", "2026-08-03", "Evening shift")] };
const lateCorrection = await replaceDerivedRosterFile(db, file, doctors, lateCorrectionEvents);
assert.deepEqual(lateCorrection.affectedDates, ["2026-08-03"]);
await publishFacilityStaffMetadata({ env: { ROSTER_DB: db, ROSTER_FILES: r2 } }, ["mmc"]);
r2.failPointerOnce = true;
await assert.rejects(
  publishFacilityDays({ env: { ROSTER_DB: db, ROSTER_FILES: r2 } }, "mmc", lateCorrection.affectedDates),
  /Injected manifest failure/,
);
let retainedDay = await loadPublishedFacilityDays(r2, ["mmc"], "2026-08-03");
assert.ok(retainedDay.rows.some((row) => row.event?.title === "Sick leave"), "failed publication must retain the previous complete day");
await publishFacilityDays({ env: { ROSTER_DB: db, ROSTER_FILES: r2 } }, "mmc", lateCorrection.affectedDates);
retainedDay = await loadPublishedFacilityDays(r2, ["mmc"], "2026-08-03");
assert.ok(retainedDay.rows.some((row) => row.event?.title === "Evening shift"), "retry must publish the corrected day");

const concurrencyCorrection = { ...initialEvents, "TERM TRAINEE": [event("trainee-1", "2026-08-03", "Night shift")] };
const concurrencyChange = await replaceDerivedRosterFile(db, file, doctors, concurrencyCorrection);
await publishFacilityStaffMetadata({ env: { ROSTER_DB: db, ROSTER_FILES: r2 } }, ["mmc"]);
let releasePublisher;
let publisherReady;
const ready = new Promise((resolve) => { publisherReady = resolve; });
const release = new Promise((resolve) => { releasePublisher = resolve; });
const firstPublisher = publishFacilityDays({ env: { ROSTER_DB: db, ROSTER_FILES: r2 } }, "mmc", concurrencyChange.affectedDates, {
  beforePointer: async () => { publisherReady(); await release; },
});
await ready;
const competingPublisher = await publishFacilityDays({ env: { ROSTER_DB: db, ROSTER_FILES: r2 } }, "mmc", concurrencyChange.affectedDates);
assert.equal(competingPublisher.busy, true, "a second Worker instance must not publish while the ED lease is held");
releasePublisher();
await firstPublisher;

const expiryCorrection = { ...initialEvents, "TERM TRAINEE": [event("trainee-1", "2026-08-03", "Recovered shift")] };
const expiryChange = await replaceDerivedRosterFile(db, file, doctors, expiryCorrection);
await publishFacilityStaffMetadata({ env: { ROSTER_DB: db, ROSTER_FILES: r2 } }, ["mmc"]);
let releaseExpiredPublisher;
let expiredPublisherReady;
const expiredReady = new Promise((resolve) => { expiredPublisherReady = resolve; });
const expiredRelease = new Promise((resolve) => { releaseExpiredPublisher = resolve; });
const expiredPublisher = publishFacilityDays({ env: { ROSTER_DB: db, ROSTER_FILES: r2 } }, "mmc", expiryChange.affectedDates, {
  beforePointer: async () => { expiredPublisherReady(); await expiredRelease; },
});
await expiredReady;
sqlite.prepare("UPDATE facility_day_publications SET lease_expires_at = '2000-01-01T00:00:00Z' WHERE source_type = 'mmc'").run();
const replacementPublisher = await publishFacilityDays({ env: { ROSTER_DB: db, ROSTER_FILES: r2 } }, "mmc", expiryChange.affectedDates);
assert.equal(replacementPublisher.ok, true, "an expired publication lease must be recoverable by a new owner");
releaseExpiredPublisher();
await assert.rejects(expiredPublisher, /lease was superseded/, "an expired old worker must not overwrite the replacement generation");
const recoveredDay = await loadPublishedFacilityDays(r2, ["mmc"], "2026-08-03");
assert.ok(recoveredDay.rows.some((row) => row.event?.title === "Recovered shift"));

const postPointerCorrection = { ...initialEvents, "TERM TRAINEE": [event("trainee-1", "2026-08-03", "Committed shift")] };
const postPointerChange = await replaceDerivedRosterFile(db, file, doctors, postPointerCorrection);
await publishFacilityStaffMetadata({ env: { ROSTER_DB: db, ROSTER_FILES: r2 } }, ["mmc"]);
await assert.rejects(
  publishFacilityDays({ env: { ROSTER_DB: db, ROSTER_FILES: r2 } }, "mmc", postPointerChange.affectedDates, {
    afterPointer: async () => { throw new Error("Injected completion failure"); },
  }),
  /Injected completion failure/,
);
const committedDespiteMarkerFailure = await loadPublishedFacilityDays(r2, ["mmc"], "2026-08-03");
assert.ok(committedDespiteMarkerFailure.rows.some((row) => row.event?.title === "Committed shift"), "a committed pointer must remain readable if completion marking fails");
const recoveredCompletion = await publishFacilityDays({ env: { ROSTER_DB: db, ROSTER_FILES: r2 } }, "mmc", postPointerChange.affectedDates);
assert.equal(recoveredCompletion.unchanged, true, "retry must recognise an already committed pointer without rewriting R2");

const secondFile = { ...file, id: "incremental-b", name: "Overlapping.xlsx" };
await replaceDerivedRosterFile(db, secondFile, [doctors[1]], { "TERM TRAINEE": [event("trainee-2", "2026-08-04")] });
let staff = await queryMaterializedFacilityTermStaff(db, { sourceType: "mmc", termStart: "2026-08-03" });
assert.equal(staff.find((row) => row.doctorKey === "TERM TRAINEE").contributionCount, 2);
await deleteDerivedRosterFile(db, file.id);
staff = await queryMaterializedFacilityTermStaff(db, { sourceType: "mmc", termStart: "2026-08-03" });
assert.equal(staff.find((row) => row.doctorKey === "TERM TRAINEE").contributionCount, 1, "deleting one file must preserve another contribution");
assert.equal(sqlite.prepare("SELECT COUNT(*) AS count FROM facility_sms_memberships WHERE doctor_key='PERMANENT SMS'").get().count, 1, "SMS continuity must survive a missing term/file");

const routeSqlite = new DatabaseSync(":memory:");
for (const name of (await readdir(new URL("../migrations", import.meta.url))).filter((name) => name.endsWith(".sql")).sort()) {
  routeSqlite.exec(await readFile(new URL(`../migrations/${name}`, import.meta.url), "utf8"));
}
const routeDb = new LocalD1(routeSqlite);
const token = "local-automation-test";
routeSqlite.prepare(`INSERT INTO roster_sources (id, provider, source_type, label, enabled) VALUES (?, ?, ?, ?, 1)`)
  .run("monash-adults", "sharepoint", "mmc", "Monash Adults");

async function runCompleteRoute(runId, incomingFileId, events) {
  routeSqlite.prepare(`INSERT INTO roster_sync_runs (id, source_id, trigger_type, file_id, source_file_id, status, started_at) VALUES (?, ?, ?, ?, ?, 'queued', ?)`)
    .run(runId, "monash-adults", "automatic", incomingFileId, incomingFileId, new Date().toISOString());
  const response = await saveAutomatedDerivedRoster({
    request: new Request("http://127.0.0.1/api/automation/derived", {
      method: "POST",
      headers: { authorization: `Bearer ${token}`, "content-type": "application/json" },
      body: JSON.stringify({
        runId,
        sourceId: "monash-adults",
        phase: "complete",
        file: { ...file, id: incomingFileId, sourceId: "monash-adults", name: "Automated.xlsx" },
        doctors,
        eventsByDoctor: events,
        issuesByDoctor: {},
      }),
    }),
    env: { ROSTER_DB: routeDb, ROSTER_AUTOMATION_TOKEN: token, ROSTER_AUTOMATION_WRITES_ENABLED: "true", ROSTER_AUTOMATION_SOURCE_ALLOWLIST: "monash-adults" },
    waitUntil() {},
  });
  const payload = await response.json();
  assert.equal(response.status, 200, JSON.stringify(payload));
  return payload;
}

const routeFirst = await runCompleteRoute("route-run-1", "route-file-1", initialEvents);
assert.equal(routeFirst.fileId, "route-file-1");
const routeEventCount = routeSqlite.prepare("SELECT COUNT(*) AS count FROM roster_events WHERE file_id = 'route-file-1'").get().count;
const routeRepeat = await runCompleteRoute("route-run-2", "route-file-2", initialEvents);
assert.equal(routeRepeat.fileId, "route-file-1", "repeat automation must target the stable active file");
assert.equal(routeRepeat.unchanged, true, "repeat automation must be a semantic no-op");
assert.equal(routeSqlite.prepare("SELECT COUNT(*) AS count FROM roster_events WHERE file_id = 'route-file-1'").get().count, routeEventCount);
assert.equal(routeSqlite.prepare("SELECT COUNT(*) AS count FROM roster_files WHERE id = 'route-file-2'").get().count, 0, "repeat automation must not create an inactive copy");
const routeCorrection = await runCompleteRoute("route-run-3", "route-file-3", correctedEvents);
assert.equal(routeCorrection.fileId, "route-file-1");
assert.equal(routeCorrection.changes.events, 1, "automated correction must update only its changed event fact");
assert.equal(routeSqlite.prepare("SELECT COUNT(*) AS count FROM roster_files WHERE id = 'route-file-3'").get().count, 0, "correction must not create an inactive copy");

console.log("Facility materialisation and automated-handler checks passed unchanged, correction, overlap, SMS continuity, and 14-day visibility checks.");
