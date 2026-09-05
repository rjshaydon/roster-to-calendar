import assert from "node:assert/strict";
import { DatabaseSync } from "node:sqlite";

const db = new DatabaseSync(":memory:");
db.exec(`
  CREATE TABLE roster_files (id TEXT PRIMARY KEY, source_type TEXT NOT NULL, active INTEGER NOT NULL);
  CREATE TABLE roster_file_doctors (file_id TEXT, source_type TEXT, doctor_key TEXT, display_name TEXT, seniority TEXT, membership_source TEXT, PRIMARY KEY(file_id, source_type, doctor_key));
  CREATE TABLE roster_events (id TEXT PRIMARY KEY, file_id TEXT, source_type TEXT, doctor_key TEXT, display_name TEXT, seniority TEXT, start_date TEXT, end_date TEXT, start_ts TEXT, title TEXT, event_json TEXT);
  CREATE INDEX idx_events_date_source ON roster_events(start_date, source_type);
  CREATE INDEX idx_events_source_range ON roster_events(source_type, start_date, end_date);
  CREATE INDEX idx_events_doctor_range ON roster_events(doctor_key, start_date, end_date);
  CREATE INDEX idx_events_file ON roster_events(file_id);
  CREATE INDEX idx_file_doctors_source_file ON roster_file_doctors(source_type, file_id);
  CREATE TABLE contact_list_files (id TEXT PRIMARY KEY, source_id TEXT, received_at TEXT);
  CREATE INDEX idx_contacts_source_received ON contact_list_files(source_id, received_at DESC);
  CREATE TABLE roster_file_coverage (file_id TEXT PRIMARY KEY, source_type TEXT, coverage_start TEXT, coverage_end TEXT);
  CREATE INDEX idx_compact_coverage ON roster_file_coverage(source_type, coverage_start, coverage_end, file_id);
  CREATE TABLE facility_term_staff_contributions (source_type TEXT, term_start TEXT, doctor_key TEXT, file_id TEXT, display_name TEXT, PRIMARY KEY(source_type,term_start,doctor_key,file_id));
  CREATE INDEX idx_compact_staff ON facility_term_staff_contributions(source_type,term_start,doctor_key);
`);

const sources = ["mmc", "ddh", "casey", "mch", "vhh"];
const doctorsPerSource = 120;
const days = 182;
db.exec("BEGIN");
const addFile = db.prepare("INSERT INTO roster_files VALUES (?, ?, 1)");
const addCoverage = db.prepare("INSERT INTO roster_file_coverage VALUES (?, ?, '2026-01-01', '2026-07-01')");
const addCompactStaff = db.prepare("INSERT INTO facility_term_staff_contributions VALUES (?, '2026-05-04', ?, ?, ?)");
const addDoctor = db.prepare("INSERT INTO roster_file_doctors VALUES (?, ?, ?, ?, ?, 'roster')");
const addEvent = db.prepare("INSERT INTO roster_events VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'Day', '{}')");
for (const source of sources) {
  const file = `fixture-${source}`;
  addFile.run(file, source);
  addCoverage.run(file, source);
  for (let doctor = 0; doctor < doctorsPerSource; doctor += 1) {
    const key = `${source.toUpperCase()} DOCTOR ${String(doctor).padStart(3, "0")}`;
    addDoctor.run(file, source, key, key, doctor < 25 ? "SMS" : "Registrar");
    addCompactStaff.run(source, key, file, key);
    for (let day = 0; day < days; day += 1) {
      const date = new Date(Date.UTC(2026, 0, 1 + day)).toISOString().slice(0, 10);
      addEvent.run(`${file}:${doctor}:${day}`, file, source, key, key, doctor < 25 ? "SMS" : "Registrar", date, date, `${date}T08:00:00`);
    }
  }
}
const addContact = db.prepare("INSERT INTO contact_list_files VALUES (?, ?, ?)");
for (let index = 0; index < 40; index += 1) addContact.run(`contact-${index}`, "monash", `2026-06-${String(30 - (index % 30)).padStart(2, "0")}`);
db.exec("COMMIT");

const cases = {
  staff: [`SELECT DISTINCT d.doctor_key FROM roster_file_doctors d JOIN roster_files f ON f.id=d.file_id WHERE f.active=1 AND f.source_type=? AND EXISTS (SELECT 1 FROM roster_events e WHERE e.file_id=f.id AND e.start_date<=? AND e.end_date>=? LIMIT 1)`, ["mmc", "2026-06-30", "2026-04-01"]],
  coverage: [`SELECT e.source_type, MIN(e.start_date), MAX(e.start_date) FROM roster_events e JOIN roster_files f ON f.id=e.file_id WHERE f.active=1 AND e.source_type=? GROUP BY e.source_type`, ["mmc"]],
  onShift: [`SELECT e.doctor_key FROM roster_events e JOIN roster_files f ON f.id=e.file_id WHERE f.active=1 AND e.source_type=? AND e.start_date=?`, ["mmc", "2026-06-01"]],
  access: [`SELECT e.source_type,e.doctor_key FROM roster_events e JOIN roster_files f ON f.id=e.file_id WHERE f.active=1 AND e.doctor_key=? AND e.start_date<=? AND e.end_date>=?`, ["MMC DOCTOR 001", "2026-06-30", "2026-04-01"]],
  contacts: [`SELECT id FROM contact_list_files WHERE source_id=? ORDER BY received_at DESC LIMIT 8`, ["monash"]],
  compactCoverage: [`SELECT c.file_id,c.coverage_start,c.coverage_end FROM roster_file_coverage c JOIN roster_files f ON f.id=c.file_id WHERE f.active=1 AND c.source_type=?`, ["mmc"]],
  compactStaff: [`SELECT s.doctor_key,MAX(s.display_name) FROM facility_term_staff_contributions s JOIN roster_files f ON f.id=s.file_id WHERE f.active=1 AND s.source_type=? AND s.term_start=? GROUP BY s.doctor_key`, ["mmc", "2026-05-04"]],
};

function plan(sql, bindings) {
  return db.prepare(`EXPLAIN QUERY PLAN ${sql}`).all(...bindings).map((row) => String(row.detail));
}
const report = {};
for (const [label, [sql, bindings]] of Object.entries(cases)) {
  const started = performance.now();
  const returned = db.prepare(sql).all(...bindings).length;
  report[label] = { returned, elapsedMs: Number((performance.now() - started).toFixed(3)), plan: plan(sql, bindings) };
}

const totalEvents = sources.length * doctorsPerSource * days;
report.fixture = { files: sources.length, doctors: sources.length * doctorsPerSource, events: totalEvents };
report.estimates = {
  staff: "candidate membership rows plus repeated event-index probes; grows with files/membership and covered history",
  coverage: `${doctorsPerSource * days} ED event rows examined to return one aggregate row`,
  onShift: `${doctorsPerSource} ED/date event rows examined and returned`,
  access: `${days} indexed rows for one doctor before range filtering`,
  contacts: "at most eight returned, but the existing endpoint repeats authentication/access and R2 discovery on every poll",
  compactCoverage: "one compact row per contributing active file; zero roster-event rows",
  compactStaff: "one compact contribution per doctor/file for the selected ED/term; zero roster-event rows",
};

assert.equal(totalEvents, 109200);
assert.ok(report.onShift.plan.some((line) => /SEARCH e USING INDEX idx_events_(?:date_source|source_range).*(?:start_date|source_type)/.test(line)), "On shift must remain an indexed ED/date lookup");
assert.ok(report.access.plan.some((line) => /SEARCH e USING INDEX idx_events_doctor_range/.test(line)), "access must use its doctor/date index");
assert.ok(report.contacts.plan.some((line) => /SEARCH contact_list_files USING INDEX idx_contacts_source_received/.test(line)), "contact discovery must use its source/revision index");
assert.ok(report.coverage.plan.some((line) => /SEARCH e USING/.test(line)), "baseline coverage plan should be recorded as a broad source-history index walk");
assert.ok(report.compactCoverage.plan.every((line) => !/roster_events/.test(line)), "compact coverage must not access roster events");
assert.ok(report.compactStaff.plan.every((line) => !/roster_events/.test(line)), "compact staff must not access roster events");

console.log(JSON.stringify(report, null, 2));
