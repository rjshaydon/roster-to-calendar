import assert from "node:assert/strict";
import { readFile, readdir } from "node:fs/promises";
import { DatabaseSync } from "node:sqlite";
import {
  deleteDerivedRosterFile,
  queryMaterializedFacilityCoverage,
  queryMaterializedFacilityTermStaff,
  replaceDerivedRosterFile,
} from "../functions/_lib/d1-calendar.js";

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

const secondFile = { ...file, id: "incremental-b", name: "Overlapping.xlsx" };
await replaceDerivedRosterFile(db, secondFile, [doctors[1]], { "TERM TRAINEE": [event("trainee-2", "2026-08-04")] });
let staff = await queryMaterializedFacilityTermStaff(db, { sourceType: "mmc", termStart: "2026-08-03" });
assert.equal(staff.find((row) => row.doctorKey === "TERM TRAINEE").contributionCount, 2);
await deleteDerivedRosterFile(db, file.id);
staff = await queryMaterializedFacilityTermStaff(db, { sourceType: "mmc", termStart: "2026-08-03" });
assert.equal(staff.find((row) => row.doctorKey === "TERM TRAINEE").contributionCount, 1, "deleting one file must preserve another contribution");
assert.equal(sqlite.prepare("SELECT COUNT(*) AS count FROM facility_sms_memberships WHERE doctor_key='PERMANENT SMS'").get().count, 1, "SMS continuity must survive a missing term/file");

console.log("Facility materialisation passed unchanged, correction, overlap, SMS continuity, and 14-day visibility checks.");
