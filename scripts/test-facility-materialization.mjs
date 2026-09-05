import assert from "node:assert/strict";
import { readFile, readdir } from "node:fs/promises";
import { DatabaseSync } from "node:sqlite";
import {
  deleteDerivedRosterFile,
  queryMaterializedFacilityCoverage,
  queryMaterializedFacilityTermStaff,
  replaceDerivedRosterFile,
} from "../functions/_lib/d1-calendar.js";
import { onRequestPost as saveAutomatedDerivedRoster } from "../functions/api/automation/derived.js";

class LocalD1 {
  constructor(sqlite) { this.sqlite = sqlite; this.rowsWritten = 0; }
  prepare(sql) {
    const owner = this;
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
