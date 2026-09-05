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
import { loadPublishedFacilityMetadata, loadPublishedFacilityStaff, publishFacilityStaffMetadata } from "../functions/_lib/facility-overview-cache.js";

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
  constructor() { this.objects = new Map(); this.puts = 0; this.gets = 0; }
  async put(key, value, options = {}) {
    this.puts += 1;
    const bytes = value instanceof Uint8Array ? value : new Uint8Array(value);
    this.objects.set(key, { bytes, options });
  }
  async get(key) {
    this.gets += 1;
    const item = this.objects.get(key);
    if (!item) return null;
    return { arrayBuffer: async () => item.bytes.buffer.slice(item.bytes.byteOffset, item.bytes.byteOffset + item.bytes.byteLength) };
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
    env: { ROSTER_DB: db, ROSTER_FILES: options.r2 || r2, FACILITY_ACCESS_MATERIALIZATION_ENABLED: "true", FACILITY_SHARED_METADATA_ENABLED: "true" },
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
