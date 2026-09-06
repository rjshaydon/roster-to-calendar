import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFile, readdir } from "node:fs/promises";
import { DatabaseSync } from "node:sqlite";
import {
  deleteDerivedRosterFile,
  queryMaterializedFacilityCoverage,
  queryMaterializedFacilityTermStaff,
  replaceDerivedRosterFile,
} from "../functions/_lib/d1-calendar.js";
import { onRequestPost as saveAutomatedDerivedRoster } from "../functions/api/automation/derived.js";
import { onRequestPost as stateHandler } from "../functions/api/state.js";
import { loadPublishedFacilityDays, loadPublishedFacilityMetadata, loadPublishedFacilityRange, loadPublishedFacilityStaff, publishFacilityDays, publishFacilityStaffMetadata } from "../functions/_lib/facility-overview-cache.js";

class LocalD1 {
  constructor(sqlite) { this.sqlite = sqlite; this.rowsWritten = 0; this.sql = []; }
  prepare(sql) {
    const owner = this;
    owner.sql.push(sql.replace(/\s+/g, " ").trim());
    return {
      args: [],
      bind(...args) { this.args = args; return this; },
      async run() { const result = owner.sqlite.prepare(sql).run(...this.args); owner.rowsWritten += Number(result.changes || 0); return { success: true, meta: { changes: Number(result.changes || 0) } }; },
      async all() { return { success: true, results: owner.sqlite.prepare(sql).all(...this.args) }; },
      async first() { return owner.sqlite.prepare(sql).get(...this.args) || null; },
    };
  }
  async batch(statements) { const results = []; for (const statement of statements) results.push(await statement.run()); return results; }
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
for (const name of (await readdir(new URL("../migrations", import.meta.url))).filter((name) => name.endsWith(".sql")).sort()) {
  sqlite.exec(await readFile(new URL(`../migrations/${name}`, import.meta.url), "utf8"));
}
const db = new LocalD1(sqlite);
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
assert.ok(db.rowsWritten <= 6, `one correction wrote ${db.rowsWritten} rows`);

db.rowsWritten = 0;
await assert.rejects(
  replaceDerivedRosterFile(db, file, doctors, {
    "PERMANENT SMS": [event("sms-1", "2026-08-03", "Changed SMS shift")],
    "TERM TRAINEE": [event("trainee-1", "2026-08-03", "Changed trainee shift")],
  }, {}, { maximumIncrementalFacts: 1 }),
  (error) => error?.code === "ROSTER_INCREMENTAL_BUDGET" && error.changedFactCount === 2,
);
assert.equal(db.rowsWritten, 0, "an over-budget automatic revision must stop before writes");

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
    env: { ROSTER_DB: db, ROSTER_FILES: options.r2 || r2, FACILITY_ACCESS_MATERIALIZATION_ENABLED: "true", FACILITY_SHARED_METADATA_ENABLED: "true", FACILITY_SHARED_DAYS_ENABLED: "true" },
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
    env: { ROSTER_DB: routeDb, ROSTER_AUTOMATION_TOKEN: token, ROSTER_AUTOMATION_WRITES_ENABLED: "true" },
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
