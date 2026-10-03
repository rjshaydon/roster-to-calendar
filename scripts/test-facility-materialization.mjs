import { configureRosterMaintenanceBudget, reserveRosterMaintenanceBudget } from "../functions/_lib/roster-maintenance-budget.js";
import { onRequest as middleware } from "../functions/_middleware.js";
import { onRequestPost as refreshFacility } from "../functions/api/automation/facility-refresh.js";
import { facilityRefreshStatements } from "../functions/_lib/facility-refresh-queue.js";
import { executeBoundedRosterImport } from "./roster-import-driver.mjs";
import { ROSTER_IMPORT_WRITE_COST } from "../functions/_lib/roster-import-batches.js";
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
  queryCoworkerEvents,
  queryOverlapDoctors,
  listAccountDirectoryPage,
} from "../functions/_lib/d1-calendar.js";
import { onRequestPost as saveAutomatedDerivedRoster } from "../functions/api/automation/derived.js";
import { onRequestPost as bootstrapFacility } from "../functions/api/automation/facility-bootstrap.js";
import { onRequestPost as materializeFacility } from "../functions/api/automation/facility-materialize.js";
import { onRequestPost as ingestContacts } from "../functions/api/automation/contact-list-extract.js";
import { onRequestPost as stateHandler } from "../functions/api/state.js";
import { loadPublishedRosterDoctors, facilityPublicationBatches, initializeFacilityMaterialization, loadPublishedFacilityDays, loadPublishedFacilityMetadata, loadPublishedFacilityRange, loadPublishedFacilityStaff, publishFacilityDays, publishFacilityStaffMetadata, runFacilityPublicationStep } from "../functions/_lib/facility-overview-cache.js";
import { contactOperationalDate } from "../public/static/contact-allocations.js";
import { planRosterImportBatches } from "../functions/_lib/roster-import-batches.js";
import { beginBoundedRosterImport, stageBoundedRosterBatch, prepareBoundedRosterPresence, prepareBoundedRosterMetadata, activateBoundedRosterTerm } from "../functions/_lib/roster-import-staging.js";

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
        return { success: true, meta: { changes: Number(result.changes || 0), rows_read: Number(result.changes || 0) + 1, rows_written: Number(result.changes || 0) } };
      },
      // Metadata is simulated here to test reservation/refund logic, not to
      // claim SQLite result counts measure Cloudflare index billing.
      async all() { const results = owner.sqlite.prepare(sql).all(...this.args); return { success: true, results, meta: { rows_read: results.length + 1, rows_written: 0 } }; },
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
    if (options.onlyIf?.etagDoesNotMatch === "*" && current) throw new Error("Precondition failed");
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
for (const [table, kind] of [["roster_events", "event"], ["roster_issues", "issue"], ["roster_daily_presence", "presence"], ["roster_doctors", "doctor"], ["roster_file_doctors", "fileDoctor"], ["facility_sms_memberships", "sms"], ["facility_term_staff_contributions", "compact"], ["facility_stream_catalog_contributions", "compact"]]) {
  assert.ok(ROSTER_IMPORT_WRITE_COST[kind] >= 1 + sqlite.prepare(`PRAGMA index_list(${table})`).all().length, `${table}: reservation must include every index`);
}

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
const contactPayload = { sourceId: "mmc-shift-allocations", sourceDate: contactOperationalDate(), providerModifiedAt: "2026-09-06T01:00:00Z", contacts: [{ area: "Adult Emergency", shift: "AM", role: "Consultant", name: "Alex Example", phone: "555-0100", isPopulated: true }] };
async function callContactExtract(overrides = {}) {
  return ingestContacts({
    request: new Request("http://local/api/automation/contact-list-extract", { method: "POST", headers: { authorization: "Bearer contact-token", "content-type": "application/json" }, body: JSON.stringify({ ...contactPayload, ...overrides }) }),
    env: {
      ROSTER_DB: db,
      ROSTER_FILES: contactR2,
      ROSTER_AUTOMATION_TOKEN: "contact-token",
      CONTACT_AUTOMATION_WRITES_ENABLED: "true",
      CONTACT_AUTOMATION_SOURCE_ALLOWLIST: "mmc-shift-allocations,vhh-shift-phone-allocations",
      FACILITY_SHARED_CONTACTS_BUILD_ENABLED: "true",
      FACILITY_SHARED_EMERGENCY_PAUSED: "false",
      FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST: "mmc,mch,vhh",
    },
  });
}
const vhhHttpPayload = { sourceId: "vhh-shift-phone-allocations", sourceDate: new Intl.DateTimeFormat("en-CA", { timeZone: "Australia/Melbourne", year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date()),
  doctors: [{ role: "ED Doctor", phone: "12017", name: "Alex Example" }], contacts: undefined };
assert.equal((await callContactExtract({ ...vhhHttpPayload, sourceDate: "" })).status, 400, "undated VHH discovery data is rejected before storage");
const vhhHttpResult = await callContactExtract(vhhHttpPayload);
assert.equal(vhhHttpResult.status, 200);
assert.equal((await vhhHttpResult.json()).status, "stored");
assert.ok(contactR2.objects.has("facility-overview/v1/contacts/vhh-shift-phone-allocations/manifest.json"), "VHH HTTP ingress publishes its own bounded R2 overlay");
db.rowsWritten = 0;
const vhhHttpWrites = contactR2.puts;
assert.equal((await (await callContactExtract(vhhHttpPayload)).json()).status, "unchanged");
assert.equal(db.rowsWritten, 0);
assert.equal(contactR2.puts, vhhHttpWrites, "an unchanged VHH submission writes no D1 rows or R2 objects");
const firstContact = await callContactExtract();
assert.equal(firstContact.status, 200);
assert.equal((await firstContact.json()).status, "stored");
db.rowsWritten = 0;
const contactPuts = contactR2.puts;
const repeatedContact = await callContactExtract({
  providerModifiedAt: "2026-09-06T02:00:00Z",
  providerVersion: "metadata-only-change",
});
assert.equal((await repeatedContact.json()).status, "unchanged");
assert.equal(db.rowsWritten, 0, "an allocation with changed provider metadata must write no D1 rows");
assert.equal(contactR2.puts, contactPuts, "an allocation with changed provider metadata must write no R2 objects");

async function callBootstrap(body, envOverrides = {}) {
  const response = await bootstrapFacility({
    request: new Request("http://local/api/automation/facility-bootstrap", { method: "POST", headers: { authorization: "Bearer bootstrap-token", "content-type": "application/json" }, body: JSON.stringify(body) }),
    env: {
      ROSTER_DB: db,
      ROSTER_AUTOMATION_TOKEN: "bootstrap-token",
      FACILITY_BOOTSTRAP_INSPECTION_ENABLED: "true",
      FACILITY_BOOTSTRAP_EXECUTION_ENABLED: "true",
      FACILITY_BOOTSTRAP_FILE_ALLOWLIST: String(body?.fileId || ""),
      ROSTER_ADVANCED_MAINTENANCE_ENABLED: "true",
      FACILITY_SHARED_EMERGENCY_PAUSED: "false",
      FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST: "mmc",
      ...envOverrides,
    },
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
const staleBootstrap = await callBootstrap({ sourceType: "mmc", fileId: legacyFile.id, maximumEventRows: 10, maximumWrites: 100, execute: true, planRevision: "stale", planGeneratedAt: bootstrapPlan.payload.planGeneratedAt });
assert.equal(staleBootstrap.response.status, 409);
db.rowsWritten = 0;
const writeLimitedBootstrap = await callBootstrap({ sourceType: "mmc", fileId: legacyFile.id, maximumEventRows: 10, maximumWrites: 1, execute: true, planRevision: bootstrapPlan.payload.planRevision, planGeneratedAt: bootstrapPlan.payload.planGeneratedAt });
assert.equal(writeLimitedBootstrap.response.status, 409);
assert.equal(writeLimitedBootstrap.payload.result.reason, "compact-write-limit");
assert.equal(db.rowsWritten, 0, "a compact-write overage must stop before writes");
db.rowsWritten = 0;
const bootstrapExecution = await callBootstrap({ sourceType: "mmc", fileId: legacyFile.id, maximumEventRows: 10, maximumWrites: 100, execute: true, planRevision: bootstrapPlan.payload.planRevision, planGeneratedAt: bootstrapPlan.payload.planGeneratedAt });
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
const repeatedBootstrap = await callBootstrap({ sourceType: "mmc", fileId: legacyFile.id, maximumEventRows: 10, maximumWrites: 100, execute: true, planRevision: bootstrapExecution.payload.planRevision, planGeneratedAt: bootstrapExecution.payload.planGeneratedAt });
assert.equal(repeatedBootstrap.payload.unchanged, true);
assert.equal(db.rowsWritten, 0, "a repeated compact bootstrap must write nothing");

const oversizedFile = { ...file, id: "bootstrap-oversized", name: "Oversized.xlsx" };
await startDerivedRosterFileSave(db, oversizedFile, doctors);
await appendDerivedRosterFileEvents(db, oversizedFile, doctors, initialEvents);
sqlite.prepare("UPDATE roster_files SET active = 1 WHERE id = ?").run(oversizedFile.id);
const oversizedPlan = await callBootstrap({ sourceType: "mmc", fileId: oversizedFile.id, maximumEventRows: 1, maximumWrites: 100 });
db.rowsWritten = 0;
const oversizedExecution = await callBootstrap({ sourceType: "mmc", fileId: oversizedFile.id, maximumEventRows: 1, maximumWrites: 100, execute: true, planRevision: oversizedPlan.payload.planRevision, planGeneratedAt: oversizedPlan.payload.planGeneratedAt });
assert.equal(oversizedExecution.response.status, 409);
assert.equal(oversizedExecution.payload.result.reason, "event-read-limit");
assert.equal(db.rowsWritten, 0, "an over-budget bootstrap must stop before compact writes");
await deleteDerivedRosterFile(db, legacyFile.id);
await deleteDerivedRosterFile(db, oversizedFile.id);

const first = await replaceDerivedRosterFile(db, file, doctors, initialEvents);
assert.equal(first.unchanged, false);
assert.equal((await queryMaterializedFacilityCoverage(db, { sourceType: "mmc" }))[0].startDate, "2026-08-03");
sqlite.prepare("INSERT INTO roster_files (id, name, source_type, active) VALUES (?, ?, ?, 1)").run("unprepared-active", "Unprepared.xlsx", "mmc");
await assert.rejects(
  queryMaterializedFacilityCoverage(db, { sourceType: "mmc", maximumRows: 32 }),
  (error) => error?.reason === "active-file-not-prepared" && error?.fileIds?.includes("unprepared-active"),
  "bounded publication coverage must identify active files whose compact facts are missing",
);
sqlite.prepare("DELETE FROM roster_files WHERE id = ?").run("unprepared-active");
sqlite.prepare("INSERT INTO roster_files (id, name, source_type, active) VALUES (?, ?, ?, 1)").run(
  "historical-unprepared",
  "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx",
  "mmc",
);
assert.equal(
  (await queryMaterializedFacilityCoverage(db, { sourceType: "mmc", startDate: "2026-08-03", endDate: "2026-11-01", maximumRows: 32 }))[0].startDate,
  "2026-08-03",
  "a parseable non-overlapping historical file must not block current-term publication",
);
sqlite.prepare("DELETE FROM roster_files WHERE id = ?").run("historical-unprepared");
assert.equal((await queryMaterializedFacilityTermStaff(db, { sourceType: "mmc", termStart: "2026-08-03" })).length, 2);
assert.equal(sqlite.prepare("SELECT visible_from FROM facility_term_visibility WHERE source_type='mmc' AND term_start='2026-08-03'").get().visible_from, "2026-07-20");

db.rowsWritten = 0;
const unchanged = await replaceDerivedRosterFile(db, file, doctors, initialEvents);
assert.equal(unchanged.unchanged, true);
assert.equal(db.rowsWritten, 0, "identical import must perform zero writes");

const correctedEvents = { ...initialEvents, "TERM TRAINEE": [event("trainee-1", "2026-08-03", "Sick leave")] };
db.failRunIncludes = "INSERT INTO roster_events";
await assert.rejects(replaceDerivedRosterFile(db, file, doctors, correctedEvents), /Injected D1 statement failure/);
assert.equal(JSON.parse(sqlite.prepare("SELECT event_json FROM roster_events WHERE id = ?").get(`${file.id}:TERM TRAINEE:trainee-1`).event_json).title, "Day", "a failed correction batch must preserve the previous active event");

db.rowsWritten = 0;
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

const bulkFile = { ...file, id: "incremental-bulk", name: "Bulk.xlsx" };
const bulkInitial = {
  "PERMANENT SMS": Array.from({ length: 300 }, (_value, index) => event(
    `bulk-${index}`,
    `2026-08-${String(3 + (index % 20)).padStart(2, "0")}`,
    "Day",
  )),
};
await replaceDerivedRosterFile(db, bulkFile, [doctors[0]], bulkInitial);
db.sql = [];
const bulkCorrected = {
  "PERMANENT SMS": bulkInitial["PERMANENT SMS"].map((item) => ({ ...item, title: "Corrected", rawValue: "Corrected" })),
};
const bulkResult = await replaceDerivedRosterFile(db, bulkFile, [doctors[0]], bulkCorrected, {}, { maximumIncrementalFacts: 400 });
assert.equal(bulkResult.changes.events, 300);
assert.ok(db.sql.filter((sql) => sql.startsWith("DELETE FROM roster_daily_presence WHERE event_id IN")).length <= 3,
  "large incremental corrections must batch daily-presence deletes by the D1 bind limit");
assert.equal(db.sql.some((sql) => sql === "DELETE FROM roster_daily_presence WHERE event_id = ?"), false,
  "large incremental corrections must not emit one daily-presence delete per event");
await deleteDerivedRosterFile(db, bulkFile.id);

sqlite.prepare("INSERT INTO facility_term_visibility (source_type, term_start, visible_from, revision, updated_at) VALUES ('mmc', '2026-02-02', '2026-01-19', '', '')").run();
const plannedR2 = new LocalR2();
db.sql = [];
const materializeEnv = { ROSTER_DB: db, ROSTER_FILES: plannedR2, ROSTER_AUTOMATION_TOKEN: "bootstrap-token", ROSTER_ADVANCED_MAINTENANCE_ENABLED: "true", FACILITY_SHARED_EMERGENCY_PAUSED: "false", FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST: "mmc" };
const materializationPlanResponse = await materializeFacility({ request: new Request("http://local/api/automation/facility-materialize", { method: "POST", headers: { authorization: "Bearer bootstrap-token", "content-type": "application/json" }, body: JSON.stringify({ sourceType: "mmc", termStart: "2026-08-03" }) }), env: materializeEnv });
const materializationPlan = await materializationPlanResponse.json();
assert.ok(materializationPlan.operationRevision);
assert.equal(materializationPlan.batchSize, 7);
const legacyPlan = await initializeFacilityMaterialization({ env: { ROSTER_DB: db, ROSTER_FILES: plannedR2 } }, "mmc", { termStart: "2026-08-03" });
const stalePlan = await initializeFacilityMaterialization({ env: { ROSTER_DB: db, ROSTER_FILES: plannedR2 } }, "mmc", { termStart: "2026-08-03", dryRun: false, planRevision: "stale" });
assert.equal(stalePlan.stalePlan, true);
assert.equal(plannedR2.puts, 0, "a stale publication plan must write nothing");
db.sql = [];
const plannedPutsBefore = plannedR2.puts;
const plannedGetsBefore = plannedR2.gets;
const materializationExecution = await initializeFacilityMaterialization({ env: { ROSTER_DB: db, ROSTER_FILES: plannedR2 } }, "mmc", { termStart: "2026-08-03", dryRun: false, planRevision: legacyPlan.planRevision });
assert.equal(materializationExecution.ok, true);
assert.ok(plannedR2.puts - plannedPutsBefore <= legacyPlan.estimate.maximumR2Puts);
assert.ok(plannedR2.gets - plannedGetsBefore <= legacyPlan.estimate.maximumR2Gets);
assert.ok(db.sql.length <= legacyPlan.estimate.maximumD1ReadStatements + legacyPlan.estimate.maximumPublicationStateStatements);
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

const ninetyOneDates = Array.from({ length: 91 }, (_, index) => {
  const date = new Date("2026-08-03T12:00:00Z");
  date.setUTCDate(date.getUTCDate() + index);
  return date.toISOString().slice(0, 10);
});
const deterministicBatches = facilityPublicationBatches(ninetyOneDates);
assert.equal(deterministicBatches.length, 13, "a 91-day term must produce 13 explicit batches");
assert.ok(deterministicBatches.every((batch) => batch.length <= 7), "no publication batch may exceed seven dates");
assert.deepEqual(deterministicBatches.flat(), ninetyOneDates, "batching must preserve the complete ordered date plan");

const chunkedR2 = new LocalR2();
const chunkContext = { env: { ROSTER_DB: db, ROSTER_FILES: chunkedR2 } };
const originalChunkCoverageEnd = sqlite.prepare("SELECT coverage_end FROM roster_file_coverage WHERE file_id=?").get(file.id).coverage_end;
sqlite.prepare("UPDATE roster_file_coverage SET coverage_end='2026-11-01' WHERE file_id=?").run(file.id);
const chunkPlan = await runFacilityPublicationStep(chunkContext, "mmc", { mode: "plan", termStart: "2026-08-03" });
assert.equal(chunkPlan.ok, true);
assert.equal(chunkPlan.batchSize, 7);
assert.equal(chunkPlan.batchCount, 13, "the real 91-day operation plan must require 13 explicit requests");
assert.equal(chunkPlan.batches.flat().length, chunkPlan.plannedDates.length);
assert.equal(chunkPlan.estimate.eachBatch.maximumRowsExamined,
  chunkPlan.estimate.planning.maximumRowsExamined + 1 + (7 * 513),
  "a batch estimate must include one compact plan, one state row and only its seven indexed dates");
assert.ok(chunkPlan.estimate.eachBatch.maximumRowsExamined < 100000,
  "one bounded batch must remain below the At a glance five-minute hard stop");
const staleChunkPuts = chunkedR2.puts;
const staleChunk = await runFacilityPublicationStep(chunkContext, "mmc", { mode: "build-batch", termStart: "2026-08-03", operationRevision: "stale", batchIndex: 0 });
assert.equal(staleChunk.stalePlan, true);
assert.equal(chunkedR2.puts, staleChunkPuts, "a stale chunk plan must write no R2 objects");
db.sql = [];
const firstChunk = await runFacilityPublicationStep(chunkContext, "mmc", { mode: "build-batch", termStart: "2026-08-03", operationRevision: chunkPlan.operationRevision, batchIndex: 0 });
assert.equal(firstChunk.ok, true);
const chunkDayQueries = db.sql.filter((sql) => /roster_events\.source_type = \?[\s\S]*roster_events\.start_date = \?/.test(sql));
assert.ok(chunkDayQueries.length <= 7 && chunkDayQueries.length === chunkPlan.batches[0].length, "one batch must issue only its exact hospital/date queries");
const firstChunkPuts = chunkedR2.puts;
db.rowsWritten = 0;
const repeatedChunkPublication = await runFacilityPublicationStep(chunkContext, "mmc", { mode: "build-batch", termStart: "2026-08-03", operationRevision: chunkPlan.operationRevision, batchIndex: 0 });
assert.equal(repeatedChunkPublication.unchanged, true);
assert.equal(chunkedR2.puts, firstChunkPuts, "repeating a completed batch must write no R2 objects");
assert.equal(db.rowsWritten, 0, "repeating a completed batch must write no D1 rows");
assert.equal((await loadPublishedFacilityDays(chunkedR2, ["mmc"], "2026-08-03")).preparing, true, "staged batches must remain invisible to readers");
const outOfOrder = await runFacilityPublicationStep(chunkContext, "mmc", { mode: "build-batch", termStart: "2026-08-03", operationRevision: chunkPlan.operationRevision, batchIndex: 2 });
assert.equal(outOfOrder.reason, "previous-batch-required", "an out-of-order batch must be rejected");
for (let batchIndex = 1; batchIndex < chunkPlan.batchCount; batchIndex += 1) {
  db.sql = [];
  const batch = await runFacilityPublicationStep(chunkContext, "mmc", { mode: "build-batch", termStart: "2026-08-03", operationRevision: chunkPlan.operationRevision, batchIndex });
  assert.equal(batch.ok, true);
  assert.ok(db.sql.filter((sql) => /roster_events\.source_type = \?[\s\S]*roster_events\.start_date = \?/.test(sql)).length <= 7, "each request must query at most seven exact dates");
}
for (const month of chunkPlan.months) {
  const monthResult = await runFacilityPublicationStep(chunkContext, "mmc", { mode: "build-month", termStart: "2026-08-03", operationRevision: chunkPlan.operationRevision, month });
  assert.equal(monthResult.ok, true);
  const monthPuts = chunkedR2.puts;
  const repeatedMonth = await runFacilityPublicationStep(chunkContext, "mmc", { mode: "build-month", termStart: "2026-08-03", operationRevision: chunkPlan.operationRevision, month });
  assert.equal(repeatedMonth.unchanged, true);
  assert.equal(chunkedR2.puts, monthPuts, "repeating a completed month must write no R2 objects");
}
db.failRunIncludes = "SET status = 'complete'";
await assert.rejects(runFacilityPublicationStep(chunkContext, "mmc", { mode: "finalize", termStart: "2026-08-03", operationRevision: chunkPlan.operationRevision }), /Injected D1 statement failure/);
assert.equal((await loadPublishedFacilityDays(chunkedR2, ["mmc"], "2026-08-03")).preparing, false, "only finalisation may expose staged days");
db.failRunIncludes = "";
db.rowsWritten = 0;
const repairedFinalize = await runFacilityPublicationStep(chunkContext, "mmc", { mode: "finalize", termStart: "2026-08-03", operationRevision: chunkPlan.operationRevision });
assert.equal(repairedFinalize.repaired, true, "a repeated finalisation must repair status without rebuilding");
assert.ok(db.rowsWritten <= 1, "finalisation repair must update at most one logical row");
sqlite.prepare("UPDATE roster_file_coverage SET coverage_end=? WHERE file_id=?").run(originalChunkCoverageEnd, file.id);

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
assert.equal(unchangedPublishedRange.unchanged, undefined, "a partly covered range must not be reported as a complete unchanged answer");
assert.ok(unchangedPublishedRange.missing.length, "dates outside the visible term must report missing coverage");
assert.ok(r2.gets - rangeReadsBeforeRevalidation <= 4, "coverage revalidation remains bounded to the manifest and its publication objects");
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
    env: { CREATOR_DIRECTORY_ENABLED: options.creatorDirectory ? "true" : "false", ROSTER_DB: db, ROSTER_FILES: options.r2 || r2, FACILITY_OVERVIEW_MAINTENANCE_MODE: "false", FACILITY_SHARED_ROLLOUT_ACTIVE: "true", FACILITY_SHARED_EMERGENCY_PAUSED: "false", FACILITY_LEGACY_READS_PAUSED: "true", FACILITY_ACCESS_MATERIALIZATION_ENABLED: "true", FACILITY_SHARED_METADATA_ENABLED: "true", FACILITY_SHARED_DAYS_ENABLED: "true", FACILITY_SHARED_READER_SOURCE_ALLOWLIST: "mmc", FACILITY_SHARED_READER_COHORT: "all" },
    waitUntil() {},
  });
  const payload = await response.json();
  assert.equal(response.status, options.status || 200, JSON.stringify(payload));
  assert.equal(db.sql.some((sql) => /\broster_events\b/i.test(sql)), false, `${body.action} must not query roster_events`);
  return payload;
}
// Restored Creator surfaces use indexed account pages and visible R2 membership.
const doctorDirectory = await loadPublishedRosterDoctors(r2, "2026-08-04");
assert.equal(doctorDirectory.preparing, false);
assert.ok(doctorDirectory.doctors.some(person => person.key === "TERM TRAINEE" && person.sourceType === "mmc"));
assert.deepEqual((await loadPublishedRosterDoctors(new LocalR2(), "2026-08-04")).doctors, []);
for (let index = 0; index < 105; index += 1) {
  const email = `page-${String(index).padStart(3, "0")}@example.com`;
  sqlite.prepare("INSERT INTO account_profiles (email, real_name) VALUES (?, ?)").run(email, `Page ${index}`);
}
const page1 = await listAccountDirectoryPage(db);
assert.equal(page1.records.length, 100);
assert.ok(page1.nextCursor);
const page2 = await listAccountDirectoryPage(db, page1.nextCursor);
assert.equal(page2.nextCursor, "");
assert.equal(new Set([...page1.records, ...page2.records].map(record => record.email)).size, 106);
assert.ok(page1.records.find(record => record.email === "doctor@example.com").claims.length);
assert.ok(page1.records.every(record => !record.passwordHash && !record.subscriptionToken));
const plan = sqlite.prepare("EXPLAIN QUERY PLAN SELECT email FROM account_profiles WHERE email > ? ORDER BY email LIMIT 101").all("");
assert.ok(plan.some(row => /SEARCH.*INDEX.*email>/.test(row.detail)), "directory must use an indexed cursor range");
sqlite.prepare("UPDATE account_profiles SET role = 'creator' WHERE email = 'doctor@example.com'").run();
const creatorDoctors = await callSharedAction({ action: "listRosterDoctors" }, { creatorDirectory: true });
assert.equal(creatorDoctors.unavailable, false);
assert.ok(creatorDoctors.availableDoctors.length);
const creatorUsers = await callSharedAction({ action: "listUsers" }, { creatorDirectory: true });
assert.equal(creatorUsers.users.length, 100);
assert.ok(creatorUsers.nextCursor);
assert.equal(db.sql.some(sql => /\broster_(daily_presence|term_members|doctor_directory)\b/.test(sql)), false);
// The fixture publishes only one month, despite retaining other visible terms.
// Profiles must not replace a personal calendar from that incomplete history.
const profileCalendar = await callSharedAction({ action: "loadDoctorProfile", profileId: "fixture-trainee", doctorKey: "TERM TRAINEE", displayName: "Term Trainee", sourceTypes: ["mmc"], aliases: [{ sourceType: "mmc", key: "TERM TRAINEE", displayName: "Term Trainee" }] }, { creatorDirectory: true });
assert.equal(profileCalendar.snapshotSource, "published-roster");
assert.equal(profileCalendar.snapshotAvailable, false);
assert.equal(db.sql.some(sql => /\broster_(daily_presence|term_members|doctor_directory|file_doctors)\b/.test(sql)), false, "profile switching must not discover roster history");
const resolvedDoctor = await callSharedAction({ action: "resolveDoctorAccount", doctor: { key: "TERM TRAINEE", sourceTypes: ["mmc"] } }, { creatorDirectory: true });
assert.equal(resolvedDoctor.mode, "doctor-profile", "Creator accounts must not be treated as claimed clinician accounts");
assert.ok(db.sql.some(sql => /idx_account_claims_source_doctor_email/.test(sql)), "account resolution must use site and doctor index");
sqlite.prepare("UPDATE account_profiles SET role = 'user' WHERE email = 'doctor@example.com'").run();
await callSharedAction({ action: "listRosterDoctors" }, { creatorDirectory: true, status: 403 });
const handlerMetadata = await callSharedAction({ action: "queryFacilityOverviewMetadata", sourceTypes: ["mmc"] });
assert.ok(handlerMetadata.catalogEvents.length > 0);
const handlerStaff = await callSharedAction({ action: "queryFacilityOverviewStaff", facilityKey: "mmc", termStart: "2026-08-03", termEnd: "2026-11-01" });
assert.equal(handlerStaff.members.length, 2);
const handlerDay = await callSharedAction({ action: "queryFacilityOverviewOnShift", facilityKey: "mmc", date: "2026-08-03", includeClinicalSupport: true });
assert.equal(handlerDay.events.length, 1, "On shift handler must filter the shared day object using existing working-shift rules");
const unchangedHandlerDay = await callSharedAction({ action: "queryFacilityOverviewOnShift", facilityKey: "mmc", date: "2026-08-03", includeClinicalSupport: true, cachedRevision: handlerDay.revision });
assert.equal(unchangedHandlerDay.rosterUnchanged, true, "an unchanged On shift revision must not retransmit roster events");
assert.equal(unchangedHandlerDay.events, undefined);
const handlerRange = await callSharedAction({ action: "queryFacilityOverviewByStream", startDate: "2026-08-03", endDate: "2026-08-31", selections: [{ id: "day", facilityKey: "mmc", streamKey: "day", seniority: "ALL" }] });
assert.equal(handlerRange.events.length, 1, "By stream must use the shared monthly object and existing working-shift filtering");
assert.equal(db.sql.some((sql) => /roster_events|roster_daily_presence/i.test(sql)), false, "By stream shared reads must not query roster history");
assert.equal((await callSharedAction({ action: "queryFacilityOverviewByStream", startDate: "2026-08-03", endDate: "2026-08-31", selections: [{ id: "day", facilityKey: "mmc", streamKey: "day", seniority: "ALL" }], cachedRevision: handlerRange.revision })).unchanged, undefined, "incomplete coverage must be revalidated");
const handlerTogether = await callSharedAction({ action: "queryFacilityOverviewWorkingTogether", startDate: "2026-08-01", endDate: "2026-08-31", sourceTypes: ["mmc"], doctorKeys: ["PERMANENT SMS"] });
assert.equal(handlerTogether.events.length, 1, "Working together must filter the shared monthly object by doctor");
assert.equal(db.sql.some((sql) => /roster_events|roster_daily_presence/i.test(sql)), false, "Working together shared reads must not query roster history");
assert.equal((await callSharedAction({ action: "queryFacilityOverviewWorkingTogether", startDate: "2026-08-01", endDate: "2026-08-31", sourceTypes: ["mmc"], doctorKeys: ["PERMANENT SMS"], cachedRevision: handlerTogether.revision })).unchanged, undefined, "incomplete coverage must be revalidated");
sqlite.prepare("UPDATE account_profiles SET insights_enabled=1 WHERE email='doctor@example.com'").run();
const cachedWho = await callSharedAction({ action: "queryRosterInsights", startDate: "2026-08-01", endDate: "2026-08-31", sourceTypes: ["mmc"] });
assert.equal(cachedWho.source, "published-roster");
assert.ok(Array.isArray(cachedWho.coworkers));
assert.equal(db.sql.some(sql => /roster_events|roster_daily_presence/i.test(sql)), false, "colleague tools must never query roster history");
const cachedWhen = await callSharedAction({ action: "queryRosterOverlapDoctors", startDate: "2026-08-01", endDate: "2026-08-31", sourceTypes: ["mmc"], overlapDoctorKeys: ["PERMANENT SMS"] });
assert.equal(cachedWhen.source, "published-roster");
assert.ok(Array.isArray(cachedWhen.doctors));
await callSharedAction({ action: "queryRosterInsights", startDate: "2026-08-01", endDate: "2027-08-31", sourceTypes: ["mmc"] }, { status: 400 });
await callSharedAction({ action: "queryRosterInsights", startDate: "2026-08-01", endDate: "2026-08-31", sourceTypes: ["mmc"] }, { r2: new LocalR2(), status: 503 });
sqlite.prepare("UPDATE account_profiles SET insights_enabled=0 WHERE email='doctor@example.com'").run();
await callSharedAction({ action: "queryRosterInsights", startDate: "2026-08-01", endDate: "2026-08-31", sourceTypes: ["mmc"] }, { status: 403 });
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
const scopedAllStaff = await callSharedAction(
  { action: "queryFacilityOverviewStaff", facilityKey: "all", termStart: "2026-08-03", termEnd: "2026-11-01" },
);
assert.equal(scopedAllStaff.members.length, handlerStaff.members.length, "All my hospitals must stay within the authorised hospital set");
const missingHandlerStaff = await callSharedAction(
  { action: "queryFacilityOverviewStaff", facilityKey: "mmc", termStart: "2026-08-03", termEnd: "2026-11-01" },
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

async function runCompleteRoute(runId, incomingFileId, events, options = {}) {
  if (options.seedRun !== false) {
    routeSqlite.prepare(`INSERT INTO roster_sync_runs (id, source_id, trigger_type, file_id, source_file_id, status, started_at) VALUES (?, ?, ?, ?, ?, 'queued', ?)`)
      .run(runId, "monash-adults", "automatic", incomingFileId, incomingFileId, new Date().toISOString());
  }
  const deferred = [];
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
    env: { ROSTER_DB: routeDb, ROSTER_AUTOMATION_TOKEN: token, ROSTER_AUTOMATION_WRITES_ENABLED: "true", ROSTER_AUTOMATION_QUEUE_ENABLED: "true", ROSTER_AUTOMATION_SOURCE_ALLOWLIST: "monash-adults", ROSTER_AUTOMATION_REVIEWED_FACT_LIMIT: "1250" },
    waitUntil(promise) { deferred.push(promise); },
  });
  await Promise.all(deferred);
  const payload = await response.json();
  assert.equal(response.status, options.expectedStatus || 200, JSON.stringify(payload));
  return payload;
}

routeDb.sql = [];
const routeFirst = await runCompleteRoute("route-run-1", "route-file-1", initialEvents);
assert.equal(routeFirst.fileId, "route-file-1");
const disabledFanoutStatements = routeDb.sql.filter((sql) => /canonical_doctors|snapshot_registry|FROM account_profiles/i.test(sql));
assert.deepEqual(disabledFanoutStatements, [], `disabled identity and snapshot features must not run after ingestion: ${JSON.stringify(disabledFanoutStatements)}`);
routeDb.rowsWritten = 0;
const duplicateRouteFirst = await runCompleteRoute("route-run-1", "route-file-1", initialEvents, { seedRun: false });
assert.equal(duplicateRouteFirst.duplicate, true, "a duplicate completion callback must be recognised before roster work");
assert.equal(routeDb.rowsWritten, 0, "a duplicate completion callback must write zero D1 rows");
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

const currentTermBefore = routeSqlite.prepare("SELECT id, event_json FROM roster_events WHERE file_id = 'route-file-1' ORDER BY id").all();
const nextTermEvents = {
  "TERM TRAINEE": [event("next-trainee", "2026-11-02", "Day")],
  "PERMANENT SMS": [event("next-sms", "2026-11-03", "Day")],
};
const nextTerm = await runCompleteRoute("route-next-term", "route-next-file", nextTermEvents);
assert.equal(nextTerm.fileId, "route-next-file", "a new term needs its own file, not the current source pointer");
assert.deepEqual(routeSqlite.prepare("SELECT id, event_json FROM roster_events WHERE file_id = 'route-file-1' ORDER BY id").all(), currentTermBefore, "new-term ingestion must preserve all current-term events");
assert.equal(routeSqlite.prepare("SELECT active FROM roster_files WHERE id = 'route-file-1'").get().active, 1);
assert.equal(routeSqlite.prepare("SELECT visible_from FROM facility_term_visibility WHERE source_type='mmc' AND term_start='2026-11-02'").get().visible_from, "2026-10-19");

const nextRepeat = await runCompleteRoute("route-next-repeat", "route-next-repeat-file", nextTermEvents);
assert.equal(nextRepeat.unchanged, true);
routeDb.rowsWritten = 0;
const nextCallbackReplay = await runCompleteRoute("route-next-repeat", "route-next-repeat-file", nextTermEvents, { seedRun: false });
assert.equal(nextCallbackReplay.duplicate, true, "a completion retry must recognise the original input after the target id changes");
assert.equal(routeDb.rowsWritten, 0);
const olderCorrection = await runCompleteRoute("route-older-correction", "route-older-input", initialEvents);
assert.equal(olderCorrection.fileId, "route-file-1", "a correction for the retained current term must not target the upcoming term");
assert.equal(routeSqlite.prepare("SELECT active_file_id FROM roster_sources WHERE id='monash-adults'").get().active_file_id, "route-next-file", "an older-term correction must preserve the latest-term source pointer");
assert.equal(routeSqlite.prepare("SELECT COUNT(*) AS count FROM roster_events WHERE file_id='route-next-file'").get().count, 2);

// A complete new-term payload must not bypass the reviewed fact ceiling just
// because it has no stored predecessor. Use a realistic many-shift delivery.
const largeNewTerm = { "TERM TRAINEE": Array.from({ length: 1251 }, (_, index) => event(`large-${index}`, "2027-02-01", "Day")) };
const retainedBeforeFailure = routeSqlite.prepare("SELECT id, event_json FROM roster_events ORDER BY id").all();
const oversizedRoute = await runCompleteRoute("route-large-new-term", "route-large-file", largeNewTerm, { expectedStatus: 422 });
assert.match(oversizedRoute.error, /1250-fact safety budget/);
assert.deepEqual(routeSqlite.prepare("SELECT id, event_json FROM roster_events ORDER BY id").all(), retainedBeforeFailure);
assert.equal(routeSqlite.prepare("SELECT COUNT(*) AS count FROM roster_files WHERE id='route-large-file'").get().count, 0);

const overlappingTerms = { "TERM TRAINEE": [...initialEvents["TERM TRAINEE"], ...nextTermEvents["TERM TRAINEE"]] };
const rejectedOverlap = await runCompleteRoute("route-overlapping-terms", "route-overlapping-input", overlappingTerms, { expectedStatus: 422 });
assert.match(rejectedOverlap.error, /overlaps a different retained term range/);
assert.deepEqual(routeSqlite.prepare("SELECT id, event_json FROM roster_events ORDER BY id").all(), retainedBeforeFailure);
const targetLookup = routeDb.sql.find((sql) => sql.includes("FROM roster_files f INDEXED BY idx_roster_files_source_active"));
const targetQueryPlan = routeSqlite.prepare(`EXPLAIN QUERY PLAN ${targetLookup}`).all("mmc").map((row) => row.detail);
assert.ok(targetQueryPlan.some((line) => /SEARCH f USING INDEX idx_roster_files_source_active/.test(line)));
assert.equal(targetQueryPlan.some((line) => /SCAN|TEMP B-TREE/.test(line)), false, "target selection must use compact indexed rows without scanning event history");

const lateFailure = await saveAutomatedDerivedRoster({
  request: new Request("http://local/api/automation/derived", {
    method: "POST", headers: { authorization: `Bearer ${token}`, "content-type": "application/json" },
    body: JSON.stringify({ runId: "route-next-term", sourceId: "monash-adults", phase: "failed", file: { id: "route-next-file" } }),
  }),
  env: { ROSTER_DB: routeDb, ROSTER_AUTOMATION_TOKEN: token, ROSTER_AUTOMATION_WRITES_ENABLED: "true", ROSTER_AUTOMATION_QUEUE_ENABLED: "true", ROSTER_AUTOMATION_SOURCE_ALLOWLIST: "monash-adults" },
});
assert.equal(lateFailure.status, 409);
assert.deepEqual(routeSqlite.prepare("SELECT id, event_json FROM roster_events ORDER BY id").all(), retainedBeforeFailure, "late failure reporting must not delete completed live data");

// Simulate successful activation followed by a failed bookkeeping write.
// Cleanup must retain the live roster so retry can repair the small records.
routeDb.failRunIncludes = "UPDATE roster_sync_runs";
const activatedEvents = {
  "TERM TRAINEE": [event("activated", "2027-02-01", "Day")],
  "PERMANENT SMS": [event("activated-sms", "2027-02-02", "Day")],
};
await runCompleteRoute("route-bookkeeping-failure", "route-activated-file", activatedEvents, { expectedStatus: 422 });
assert.equal(routeSqlite.prepare("SELECT active FROM roster_files WHERE id='route-activated-file'").get().active, 1);
assert.equal(routeSqlite.prepare("SELECT COUNT(*) AS count FROM roster_events WHERE file_id='route-activated-file'").get().count, 2);
const repairedBookkeeping = await runCompleteRoute("route-bookkeeping-failure", "route-activated-file", activatedEvents, { seedRun: false });
assert.equal(repairedBookkeeping.unchanged, true);

async function withNextImportBudget(operation) {
  let result = await operation();
  if (result.deferred) {
    routeSqlite.prepare("UPDATE roster_import_daily_budget SET reserved_writes=0, reserved_reads=0 WHERE utc_day=?").run(new Date().toISOString().slice(0, 10));
    result = await operation();
  }
  assert.notEqual(result.deferred, true);
  return result;
}

const stagedFile = { ...file, id: "bounded-staged-term", sourceId: "monash-adults", active: false };
const stagedEvents = { "TERM TRAINEE": Array.from({ length: 1800 }, (_, index) => event(`bounded-${index}`, "2027-05-03", "Day")) };
const batchPlan = await planRosterImportBatches({ file: stagedFile, doctors, eventsByDoctor: stagedEvents, issuesByDoctor: {} });
await beginBoundedRosterImport(routeDb, "bounded-run", stagedFile, batchPlan);
assert.equal(routeSqlite.prepare("SELECT active FROM roster_files WHERE id=?").get(stagedFile.id).active, 0);
await assert.rejects(stageBoundedRosterBatch(routeDb, "bounded-run", stagedFile, batchPlan.revision, batchPlan.batches[1]), /processed in order/);
const changedBatch = structuredClone(batchPlan.batches[0]);
changedBatch.eventsByDoctor["TERM TRAINEE"][0].title = "Changed during retry";
await assert.rejects(stageBoundedRosterBatch(routeDb, "bounded-run", stagedFile, batchPlan.revision, changedBatch), /differs from the pinned plan/);
routeDb.failRunIncludes = "INSERT INTO roster_import_batch_receipts";
await assert.rejects(stageBoundedRosterBatch(routeDb, "bounded-run", stagedFile, batchPlan.revision, batchPlan.batches[0]), /Injected D1 statement failure/);
assert.equal(routeSqlite.prepare("SELECT COUNT(*) AS count FROM roster_events WHERE file_id=?").get(stagedFile.id).count, 0, "receipt failure must roll back all event writes in its batch");
assert.equal(routeSqlite.prepare("SELECT next_batch FROM roster_import_jobs WHERE run_id='bounded-run'").get().next_batch, 0);
await withNextImportBudget(() => stageBoundedRosterBatch(routeDb, "bounded-run", stagedFile, batchPlan.revision, batchPlan.batches[0]));
routeDb.rowsWritten = 0;
assert.equal((await stageBoundedRosterBatch(routeDb, "bounded-run", stagedFile, batchPlan.revision, batchPlan.batches[0])).duplicate, true);
assert.equal(routeDb.rowsWritten, 0, "receipt replay must write nothing");
assert.equal((await beginBoundedRosterImport(routeDb, "bounded-run", stagedFile, batchPlan)).nextBatch, 1, "resume must not clear previously persisted events");
for (const batch of batchPlan.batches.slice(1)) await withNextImportBudget(() => stageBoundedRosterBatch(routeDb, "bounded-run", stagedFile, batchPlan.revision, batch));
assert.equal(routeSqlite.prepare("SELECT COUNT(*) AS count FROM roster_events WHERE file_id=?").get(stagedFile.id).count, 1800);
assert.equal(routeSqlite.prepare("SELECT event_count FROM roster_import_jobs WHERE run_id='bounded-run'").get().event_count, 1800);
assert.equal(routeSqlite.prepare("SELECT active FROM roster_files WHERE id=?").get(stagedFile.id).active, 0, "even a fully staged import stays hidden until activation is implemented and verified");
assert.equal(routeSqlite.prepare("SELECT COUNT(*) AS count FROM roster_daily_presence WHERE event_id >= ? AND event_id < ?").get(`${stagedFile.id}:`, `${stagedFile.id};`).count, 0, "staging must not leak upcoming clinicians through the presence index");
await assert.rejects(activateBoundedRosterTerm(routeDb, "bounded-run", batchPlan.revision), /not ready for activation/);
await assert.rejects(prepareBoundedRosterMetadata(routeDb, "bounded-run", batchPlan.revision), /Incomplete staged import/);
for (const batch of batchPlan.batches) await withNextImportBudget(() => prepareBoundedRosterPresence(routeDb, "bounded-run", stagedFile, batchPlan.revision, batch));
routeDb.rowsWritten = 0;
assert.equal((await prepareBoundedRosterPresence(routeDb, "bounded-run", stagedFile, batchPlan.revision, batchPlan.batches[0])).duplicate, true);
assert.equal(routeDb.rowsWritten, 0);
const futureOptions = { sourceTypes: ["mmc"], startDate: "2027-05-03", endDate: "2027-05-03" };
assert.deepEqual(await queryCoworkerEvents(routeDb, futureOptions), [], "prepared presence must not expose the inactive roster");
assert.deepEqual(await queryCoworkerEvents(routeDb, { ...futureOptions, overlapDoctorKeys: ["TERM TRAINEE"] }), []);
assert.deepEqual(await queryOverlapDoctors(routeDb, { ...futureOptions, overlapDoctorKeys: ["TERM TRAINEE"] }), []);
await withNextImportBudget(() => prepareBoundedRosterMetadata(routeDb, "bounded-run", batchPlan.revision));
assert.equal((await queryMaterializedFacilityCoverage(routeDb, { sourceType: "mmc" })).some((row) => row.fileId === stagedFile.id), false);
routeSqlite.prepare("UPDATE roster_import_daily_budget SET reserved_writes=0, reserved_reads=0").run();
routeDb.failRunIncludes = "UPDATE roster_import_jobs SET activated";
await assert.rejects(activateBoundedRosterTerm(routeDb, "bounded-run", batchPlan.revision), /Injected D1 statement failure/);
assert.equal(routeSqlite.prepare("SELECT active FROM roster_files WHERE id=?").get(stagedFile.id).active, 0, "activation control rows must roll back together");
const activatedTerm = await activateBoundedRosterTerm(routeDb, "bounded-run", batchPlan.revision);
assert.equal(activatedTerm.events, 1800);
assert.equal((await queryCoworkerEvents(routeDb, futureOptions)).length, 1800, "only the final activation exposes the complete roster");
routeDb.rowsWritten = 0;
assert.equal((await activateBoundedRosterTerm(routeDb, "bounded-run", batchPlan.revision)).duplicate, true);
assert.equal(routeDb.rowsWritten, 0);
assert.equal((await beginBoundedRosterImport(routeDb, "bounded-run", stagedFile, batchPlan)).completed, true);

// Full-term execution for each ingestion class, with exhausted daily budget
// and later resumption simulated locally. No production data is submitted.
for (const [sourceId, sourceType] of [["monash-adults", "mmc"], ["monash-paeds", "mch"], ["dandenong-findmyshift", "ddh"], ["vhh-active-medical-roster", "vhh"]]) {
  const fullFile = { ...file, id: `full-${sourceType}`, sourceId, sourceType, name: `Full ${sourceType}.xlsx` };
  const fullDoctors = Array.from({ length: 200 }, (_, index) => ({ key: `FULL ${index}`, displayName: `Full ${index}`, seniority: "HMO" }));
  const fullEvents = Object.fromEntries(fullDoctors.map((doctor) => [doctor.key, Array.from({ length: 90 }, (_, day) => {
    const date = new Date(Date.UTC(2027, 7, 2 + day)).toISOString().slice(0, 10);
    return { ...event(`day-${day}`, date, "Day"), source: sourceType };
  })]));
  const fullPlan = await planRosterImportBatches({ file: fullFile, doctors: fullDoctors, eventsByDoctor: fullEvents, issuesByDoctor: {} });
  const runId = `full-run-${sourceType}`;
  const quotaDay = new Date().toISOString().slice(0, 10);
  routeSqlite.prepare("UPDATE roster_import_daily_budget SET reserved_writes=10000 WHERE utc_day=?").run(quotaDay);
  const deferredStart = await beginBoundedRosterImport(routeDb, runId, fullFile, fullPlan);
  assert.equal(deferredStart.deferred, true);
  assert.equal(routeSqlite.prepare("SELECT COUNT(*) AS count FROM roster_files WHERE id=?").get(fullFile.id).count, 0);
  let exhaustedDays = 0;
  const resume = async (operation) => {
    let result = await operation();
    if (result.deferred) {
      assert.equal(routeSqlite.prepare("SELECT reserved_writes FROM roster_import_daily_budget WHERE utc_day=?").get(quotaDay).reserved_writes <= 10000, true);
      exhaustedDays++;
      routeSqlite.prepare("UPDATE roster_import_daily_budget SET reserved_writes=0, reserved_reads=0 WHERE utc_day=?").run(quotaDay);
      result = await operation();
    }
    assert.notEqual(result.deferred, true);
    return result;
  };
  await resume(() => beginBoundedRosterImport(routeDb, runId, fullFile, fullPlan));
  for (const batch of fullPlan.batches) await resume(() => stageBoundedRosterBatch(routeDb, runId, fullFile, fullPlan.revision, batch));
  for (const batch of fullPlan.batches) await resume(() => prepareBoundedRosterPresence(routeDb, runId, fullFile, fullPlan.revision, batch));
  assert.equal(routeSqlite.prepare("SELECT active FROM roster_files WHERE id=?").get(fullFile.id).active, 0);
  await resume(() => prepareBoundedRosterMetadata(routeDb, runId, fullPlan.revision));
  await resume(() => activateBoundedRosterTerm(routeDb, runId, fullPlan.revision));
  assert.ok(exhaustedDays >= 3, "large imports must carry progress across budget windows, not raise the daily ceiling");
  assert.equal(routeSqlite.prepare("SELECT COUNT(*) AS count FROM roster_events WHERE file_id=?").get(fullFile.id).count, 18000);
  assert.equal(routeSqlite.prepare("SELECT active FROM roster_files WHERE id=?").get(fullFile.id).active, 1);
}

console.log("Facility materialisation and automated-handler checks passed unchanged, correction, overlap, SMS continuity, and 14-day visibility checks.");

// An overnight shift may attend the first day of the next term without
// owning that term. Both target selection and activation must preserve it.
const previousNightFile = { ...file, id: "http-previous-night", sourceId: "monash-adults" };
await replaceDerivedRosterFile(routeDb, previousNightFile, doctors, { "TERM TRAINEE": [{ ...event("previous-night", "2028-02-06", "Night"), start: "2028-02-06T22:00:00", end: "2028-02-07T08:00:00" }] });
assert.equal(routeSqlite.prepare("SELECT coverage_end FROM roster_file_coverage WHERE file_id=?").get(previousNightFile.id).coverage_end, "2028-02-07");

// Exercise the production HTTP protocol through middleware, interruption,
// completion bookkeeping failure, and durable cursor replay.
const httpFile = { ...file, id: "http-bounded-file", contentHash: "http-content", sourceId: "monash-adults", active: false };
const httpPlan = await planRosterImportBatches({ file: httpFile, doctors, eventsByDoctor: { "TERM TRAINEE": Array.from({ length: 700 }, (_, i) => event(`http-${i}`, "2028-02-07")) }, issuesByDoctor: {}, contentHash: "http-content" });
routeSqlite.prepare("INSERT INTO roster_sync_runs (id, source_id, trigger_type, file_id, source_file_id, content_hash, status, started_at) VALUES (?, ?, 'automatic', ?, ?, ?, 'queued', ?)").run("http-bounded-run", "monash-adults", httpFile.id, httpFile.id, "http-content", new Date().toISOString());
const httpEnv = { ROSTER_DB: routeDb, ROSTER_FILES: new LocalR2(), ROSTER_AUTOMATION_TOKEN: token, ROSTER_AUTOMATION_WRITES_ENABLED: "true", ROSTER_AUTOMATION_QUEUE_ENABLED: "true", ROSTER_AUTOMATION_SOURCE_ALLOWLIST: "monash-adults", ROSTER_AUTOMATION_REVIEWED_FACT_LIMIT: "1250", ROSTER_AUTOMATION_BOUNDED_IMPORT_ENABLED: "true", FACILITY_AUTOMATIC_PUBLICATION_ENABLED: "true", FACILITY_SHARED_ROLLOUT_ACTIVE: "true", FACILITY_SHARED_EMERGENCY_PAUSED: "false", FACILITY_MATERIALIZATION_SOURCE_ALLOWLIST: "mmc" };
async function httpStep(body, env = httpEnv, handler = saveAutomatedDerivedRoster, path = "derived") {
  const pending = [];
  const context = { request: new Request(`http://local/api/automation/${path}`, { method: "POST", headers: { authorization: `Bearer ${token}`, "content-type": "application/json" }, body: JSON.stringify(body) }), env: { ...env }, waitUntil(p) { pending.push(p); } };
  context.next = () => handler(context);
  const response = await middleware(context);
  await Promise.all(pending);
  return { status: response.status, data: await response.json() };
}
const protocolCallback = async (step) => {
  const response = await httpStep({ ...step, runId: "http-bounded-run", sourceId: "monash-adults", file: httpFile });
  if (response.status !== 200) throw new Error(response.data.error || JSON.stringify(response.data));
  return response.data;
};
const originalConsoleLog = console.log;
console.log = () => {};
try {
  const beforeDisabled = routeDb.sql.length;
  const disabled = await httpStep({ phase: "bounded-begin", sourceId: "monash-adults", file: httpFile }, { ...httpEnv, ROSTER_AUTOMATION_BOUNDED_IMPORT_ENABLED: "false" });
  assert.equal(disabled.status, 503);
  assert.equal(routeDb.sql.slice(beforeDisabled).some((sql) => /roster_import_|facility_refresh_jobs/.test(sql)), false);
  let interrupted = false;
  await assert.rejects(executeBoundedRosterImport(httpPlan, async (step) => {
    const result = await protocolCallback(step);
    if (step.phase === "bounded-events" && !interrupted) { interrupted = true; throw new Error("Simulated lost HTTP response"); }
    return result;
  }), /lost HTTP response/);
  routeSqlite.prepare("UPDATE roster_import_daily_budget SET reserved_writes=0, reserved_reads=0").run();
  routeDb.failRunIncludes = "UPDATE roster_sync_runs";
  await assert.rejects(executeBoundedRosterImport(httpPlan, async (step) => {
    routeSqlite.prepare("UPDATE roster_import_daily_budget SET reserved_writes=0, reserved_reads=0").run();
    return protocolCallback(step);
  }), /Injected D1 statement failure/);
  assert.equal(routeSqlite.prepare("SELECT active FROM roster_files WHERE id=?").get(httpFile.id).active, 1);
  const completed = await executeBoundedRosterImport(httpPlan, protocolCallback);
  assert.equal(completed.completed, true);
  assert.equal(routeSqlite.prepare("SELECT status FROM roster_sync_runs WHERE id=?").get("http-bounded-run").status, "success");
  assert.equal((await executeBoundedRosterImport(httpPlan, protocolCallback)).completed, true);
  assert.equal(routeSqlite.prepare("SELECT COUNT(*) AS n FROM roster_events WHERE file_id=?").get(httpFile.id).n, 700);

  // Publish a complete fixture, then rebuild one date. Other dates in that
  // month must survive and completed jobs must not requeue old term dates.
  await routeDb.batch(facilityRefreshStatements(routeDb, "mmc", ["2026-08-03", "2026-08-04"], "refresh-fixture-1"));
  async function drainRefresh() {
    for (let i = 0; i < 24; i += 1) {
      const result = await httpStep({ sourceId: "monash-adults" }, httpEnv, refreshFacility, "facility-refresh");
      assert.equal(result.status, 200, JSON.stringify(result.data));
      if (result.data.completed) return;
      assert.equal(result.data.deferred, undefined);
    }
    assert.fail("Refresh did not complete within bounded steps");
  }
  routeSqlite.prepare("UPDATE roster_import_daily_budget SET reserved_writes=0, reserved_reads=0").run();
  await drainRefresh();
  const beforePartial = await loadPublishedFacilityDays(httpEnv.ROSTER_FILES, ["mmc"], "2026-08-04");
  await routeDb.batch(facilityRefreshStatements(routeDb, "mmc", ["2026-08-03"], "refresh-fixture-2"));
  assert.deepEqual(JSON.parse(routeSqlite.prepare("SELECT dates_json FROM facility_refresh_jobs WHERE source_type='mmc' AND term_start='2026-08-03'").get().dates_json), ["2026-08-03"]);
  httpEnv.ROSTER_FILES.failPointerOnce = true;
  await assert.rejects(drainRefresh(), /Injected manifest failure/);
  assert.deepEqual(await loadPublishedFacilityDays(httpEnv.ROSTER_FILES, ["mmc"], "2026-08-04"), beforePartial);
  await drainRefresh();
  assert.deepEqual(await loadPublishedFacilityDays(httpEnv.ROSTER_FILES, ["mmc"], "2026-08-04"), beforePartial);
  const allowance = routeSqlite.prepare("SELECT reserved_reads FROM roster_import_daily_budget").get();
  assert.ok(allowance.reserved_reads < 500000, "complete metadata refunds unused read reservations");
  routeSqlite.prepare("UPDATE roster_import_daily_budget SET reserved_reads=500000").run();
  const paused = await httpStep({ sourceId: "monash-adults" }, httpEnv, refreshFacility, "facility-refresh");
  assert.equal(paused.data.deferred, true);
  const seedDeferred = await httpStep({ sourceId: "monash-adults", seedCurrent: true }, httpEnv, refreshFacility, "facility-refresh");
  assert.equal(seedDeferred.data.deferred, true);
  routeSqlite.prepare("UPDATE roster_import_daily_budget SET reserved_writes=0, reserved_reads=0").run();
  routeSqlite.prepare("DELETE FROM roster_file_coverage WHERE file_id=?").run(previousNightFile.id);
  routeSqlite.prepare("DELETE FROM roster_file_status_summaries WHERE file_id=?").run(previousNightFile.id);
  const prepared = await httpStep({ sourceId: "monash-adults", mode: "prepare-coverage" }, httpEnv, refreshFacility, "facility-refresh");
  assert.equal(prepared.status, 200, JSON.stringify(prepared.data));
  assert.equal(prepared.data.prepared, true);
  assert.equal(routeSqlite.prepare("SELECT derived_state FROM roster_file_status_summaries WHERE file_id=?").get(previousNightFile.id).derived_state, "ready");
  const preparedReplay = await httpStep({ sourceId: "monash-adults", mode: "prepare-coverage" }, httpEnv, refreshFacility, "facility-refresh");
  assert.equal(preparedReplay.data.idle, true);
  const seeded = await httpStep({ sourceId: "monash-adults", seedCurrent: true }, httpEnv, refreshFacility, "facility-refresh");
  assert.equal(seeded.status, 200);
  assert.equal(seeded.data.seeded, true);
  assert.ok(seeded.data.dateCount <= 120);
  const invalidSourceReads = routeDb.sql.length;
  assert.equal((await httpStep({ sourceId: "unknown", seedCurrent: true }, httpEnv, refreshFacility, "facility-refresh")).status, 503);
  assert.equal(routeDb.sql.length, invalidSourceReads);
} finally { console.log = originalConsoleLog; }
console.log("Bounded HTTP continuation, indexed reservation envelopes and incremental automatic publication checks passed.");

const sessionDb = new LocalD1(routeSqlite);
routeSqlite.prepare("UPDATE roster_import_daily_budget SET reserved_reads=0,reserved_writes=0").run();
configureRosterMaintenanceBudget(sessionDb, { utcDay: new Date().toISOString().slice(0,10), baseReads: 0, baseWrites: 0 });
assert.equal(await reserveRosterMaintenanceBudget(sessionDb, 5001, 0), false);
assert.equal(await reserveRosterMaintenanceBudget(sessionDb, 0, 250001), false);
assert.equal(await reserveRosterMaintenanceBudget(sessionDb, 4990, 100), true);
assert.equal(await reserveRosterMaintenanceBudget(sessionDb, 10, 0), false);
console.log("Per-pass SQL reservations enforce the admitted 5,000-write / 250,000-read ceiling.");

// Plan compact repair before touching active events; include stale contribution
// rows and stop before writes when the exact mutation reservation is refused.
const planningFile = { ...file, id: "compact-planning-fixture" };
await replaceDerivedRosterFile(db, planningFile, doctors, initialEvents);
const staffCols = sqlite.prepare("PRAGMA table_info(facility_term_staff_contributions)").all().map((row) => row.name);
const staffValues = staffCols.map((name) => name === "doctor_key" ? "'ORPHAN'" : name);
sqlite.prepare(`INSERT INTO facility_term_staff_contributions (${staffCols.join(",")}) SELECT ${staffValues.join(",")} FROM facility_term_staff_contributions WHERE file_id=? LIMIT 1`).run(planningFile.id);
db.rowsWritten = 0;
const compactDry = await refreshFacilityOverviewMaterializationForFile(db, planningFile.id, { dryRun: true, maximumWrites: 750 });
assert.equal(compactDry.ok, true);
assert.ok(compactDry.proposedWrites > 0);
assert.equal(db.rowsWritten, 0);
const compactRefused = await refreshFacilityOverviewMaterializationForFile(db, planningFile.id, { maximumWrites: 750, reserveMutationWrites: async () => false });
assert.equal(compactRefused.deferred, true);
assert.equal(db.rowsWritten, 0);
await refreshFacilityOverviewMaterializationForFile(db, planningFile.id, { maximumWrites: 750 });
assert.equal(db.rowsWritten, compactDry.proposedWrites);
assert.equal(sqlite.prepare("SELECT COUNT(*) AS n FROM facility_term_staff_contributions WHERE file_id=? AND doctor_key='ORPHAN'").get(planningFile.id).n, 0);
console.log("Exact compact planning accounts for stale rows and defers before any mutation.");
