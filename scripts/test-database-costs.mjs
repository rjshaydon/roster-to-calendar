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
  CREATE TABLE roster_file_status_summaries (file_id TEXT PRIMARY KEY, source_type TEXT, active INTEGER, updated_at TEXT, event_count INTEGER);
  CREATE INDEX idx_compact_roster_status ON roster_file_status_summaries(active,source_type,updated_at,file_id);
  CREATE TABLE roster_sources (id TEXT PRIMARY KEY, label TEXT);
  CREATE INDEX idx_roster_sources_label_id ON roster_sources(label,id);
  CREATE TABLE roster_sync_runs (id TEXT PRIMARY KEY, source_id TEXT, provider_version TEXT, content_hash TEXT, file_id TEXT, source_file_id TEXT, status TEXT, started_at TEXT, completed_at TEXT);
  CREATE INDEX idx_roster_sync_runs_source_started_id ON roster_sync_runs(source_id,started_at DESC,id DESC);
  CREATE INDEX idx_roster_sync_runs_source_hash ON roster_sync_runs(source_id,content_hash,status);
  CREATE INDEX idx_roster_sync_runs_source_version_status_started ON roster_sync_runs(source_id,provider_version,status,started_at DESC);
  CREATE INDEX idx_roster_sync_runs_source_status_started ON roster_sync_runs(source_id,status,started_at,id);
  CREATE TABLE raw_roster_files (file_id TEXT PRIMARY KEY, name TEXT);
  CREATE TABLE roster_dispatches (id TEXT PRIMARY KEY, requested_at TEXT);
  CREATE INDEX idx_roster_dispatches_requested_id ON roster_dispatches(requested_at DESC,id DESC);
  CREATE TABLE account_profiles (email TEXT PRIMARY KEY, real_name TEXT, role TEXT);
  CREATE TABLE account_claims (email TEXT, source_type TEXT, doctor_key TEXT, display_name TEXT, PRIMARY KEY(email, source_type, doctor_key));
  CREATE INDEX idx_account_claims_source_doctor_email ON account_claims(source_type,doctor_key,email);
  CREATE TABLE account_states (email TEXT PRIMARY KEY, session_json TEXT);
  CREATE TABLE snapshot_registry (owner_type TEXT, owner_id TEXT, doctor_key TEXT, range_key TEXT, status TEXT, built_revision TEXT, PRIMARY KEY(owner_type, owner_id, doctor_key, range_key));
`);

const sources = ["mmc", "ddh", "casey", "mch", "vhh"];
const doctorsPerSource = 120;
const days = 182;
db.exec("BEGIN");
const addFile = db.prepare("INSERT INTO roster_files VALUES (?, ?, 1)");
const addCoverage = db.prepare("INSERT INTO roster_file_coverage VALUES (?, ?, '2026-01-01', '2026-07-01')");
const addCompactStaff = db.prepare("INSERT INTO facility_term_staff_contributions VALUES (?, '2026-05-04', ?, ?, ?)");
const addRosterStatus = db.prepare("INSERT INTO roster_file_status_summaries VALUES (?, ?, 1, '2026-09-07T00:00:00Z', ?)");
const addDoctor = db.prepare("INSERT INTO roster_file_doctors VALUES (?, ?, ?, ?, ?, 'roster')");
const addEvent = db.prepare("INSERT INTO roster_events VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'Day', '{}')");
for (const source of sources) {
  const file = `fixture-${source}`;
  addFile.run(file, source);
  addCoverage.run(file, source);
  addRosterStatus.run(file, source, doctorsPerSource * days);
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
const addSource = db.prepare("INSERT INTO roster_sources VALUES (?, ?)");
for (let index = 0; index < 16; index += 1) addSource.run(`source-${index}`, `Source ${String(index).padStart(2, "0")}`);
const addRawFile = db.prepare("INSERT INTO raw_roster_files VALUES (?, ?)");
for (let index = 0; index < 16; index += 1) addRawFile.run(`raw-${index}`, `Roster-${index}.xlsx`);
const addRun = db.prepare("INSERT INTO roster_sync_runs VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)");
const addDispatch = db.prepare("INSERT INTO roster_dispatches VALUES (?, ?)");
for (let index = 0; index < 10000; index += 1) {
  const stamp = `2026-09-${String(1 + (index % 7)).padStart(2, "0")}T${String(index % 24).padStart(2, "0")}:${String(index % 60).padStart(2, "0")}:00Z`;
  const sourceIndex = index % 16;
  addRun.run(`run-${String(index).padStart(5, "0")}`, `source-${sourceIndex}`, `version-${index}`, `hash-${index}`, `raw-${sourceIndex}`, `raw-${sourceIndex}`, index % 5 === 0 ? "queued" : "success", stamp, stamp);
  addDispatch.run(`dispatch-${String(index).padStart(5, "0")}`, stamp);
}
const addInactiveRosterStatus = db.prepare("INSERT INTO roster_file_status_summaries VALUES (?, ?, 0, '2026-01-01T00:00:00Z', 0)");
for (let index = 0; index < 150; index += 1) addInactiveRosterStatus.run(`historical-${index}`, "mmc");
const addAccount = db.prepare("INSERT INTO account_profiles VALUES (?, ?, ?)");
const addClaim = db.prepare("INSERT INTO account_claims VALUES (?, ?, ?, ?)");
const addAccountState = db.prepare("INSERT INTO account_states VALUES (?, '{}')");
for (let index = 0; index < 10000; index += 1) {
  const email = index === 0 ? "rhaydon@gmail.com" : `synthetic-${String(index).padStart(5, "0")}@example.test`;
  addAccount.run(email, `Synthetic User ${index}`, index === 0 ? "creator" : "user");
  addAccountState.run(email);
  addClaim.run(email, "mmc", `MMC DOCTOR ${String(index % doctorsPerSource).padStart(3, "0")}`, `MMC Doctor ${index % doctorsPerSource}`);
  addClaim.run(email, "ddh", `DDH DOCTOR ${String(index % doctorsPerSource).padStart(3, "0")}`, `DDH Doctor ${index % doctorsPerSource}`);
}
db.prepare("INSERT INTO snapshot_registry VALUES (?, ?, ?, ?, 'ready', 'revision-1')").run("creator-account", "rhaydon@gmail.com", "RICHARD HAYDON", "2026-01-01:2026-12-31");
db.exec("COMMIT");

const cases = {
  staff: [`SELECT DISTINCT d.doctor_key FROM roster_file_doctors d JOIN roster_files f ON f.id=d.file_id WHERE f.active=1 AND f.source_type=? AND EXISTS (SELECT 1 FROM roster_events e WHERE e.file_id=f.id AND e.start_date<=? AND e.end_date>=? LIMIT 1)`, ["mmc", "2026-06-30", "2026-04-01"]],
  coverage: [`SELECT e.source_type, MIN(e.start_date), MAX(e.start_date) FROM roster_events e JOIN roster_files f ON f.id=e.file_id WHERE f.active=1 AND e.source_type=? GROUP BY e.source_type`, ["mmc"]],
  onShift: [`SELECT e.doctor_key FROM roster_events e JOIN roster_files f ON f.id=e.file_id WHERE f.active=1 AND e.source_type=? AND e.start_date=?`, ["mmc", "2026-06-01"]],
  access: [`SELECT e.source_type,e.doctor_key FROM roster_events e JOIN roster_files f ON f.id=e.file_id WHERE f.active=1 AND e.doctor_key=? AND e.start_date<=? AND e.end_date>=?`, ["MMC DOCTOR 001", "2026-06-30", "2026-04-01"]],
  contacts: [`SELECT id FROM contact_list_files WHERE source_id=? ORDER BY received_at DESC LIMIT 8`, ["monash"]],
  compactCoverage: [`SELECT c.file_id,c.coverage_start,c.coverage_end FROM roster_file_coverage c JOIN roster_files f ON f.id=c.file_id WHERE f.active=1 AND c.source_type=?`, ["mmc"]],
  compactStaff: [`SELECT s.doctor_key,MAX(s.display_name) FROM facility_term_staff_contributions s JOIN roster_files f ON f.id=s.file_id WHERE f.active=1 AND s.source_type=? AND s.term_start=? GROUP BY s.doctor_key`, ["mmc", "2026-05-04"]],
  bootstrapEvents: [`SELECT id, doctor_key, display_name, seniority, start_date, end_date, event_json FROM roster_events WHERE file_id=? LIMIT ?`, ["fixture-mmc", 25001]],
  compactRosterStatus: [`SELECT file_id, source_type, event_count FROM roster_file_status_summaries WHERE active=1 ORDER BY source_type,updated_at,file_id LIMIT 100`, []],
  expectedRosterStatus: [`SELECT file_id, source_type, event_count FROM roster_file_status_summaries WHERE file_id IN (?, ?) LIMIT 100`, ["fixture-mmc", "historical-149"]],
  boundedSources: [`SELECT * FROM roster_sources ORDER BY label,id LIMIT 16`, []],
  latestSourceRun: [`SELECT * FROM roster_sync_runs WHERE source_id=? ORDER BY started_at DESC,id DESC LIMIT 1`, ["source-3"]],
  exactProviderVersion: [`SELECT roster_sync_runs.id FROM roster_sync_runs INNER JOIN raw_roster_files ON raw_roster_files.file_id=COALESCE(NULLIF(roster_sync_runs.source_file_id,''),roster_sync_runs.file_id) WHERE roster_sync_runs.source_id=? AND roster_sync_runs.provider_version=? AND LOWER(raw_roster_files.name)=LOWER(?) AND roster_sync_runs.status IN ('success','queued','processing','failed') ORDER BY CASE roster_sync_runs.status WHEN 'success' THEN 0 WHEN 'processing' THEN 1 WHEN 'queued' THEN 2 ELSE 3 END, roster_sync_runs.started_at DESC LIMIT 1`, ["source-3", "version-3", "Roster-3.xlsx"]],
  exactContentHash: [`SELECT id FROM roster_sync_runs WHERE source_id=? AND content_hash=? AND status='success' ORDER BY completed_at DESC LIMIT 1`, ["source-3", "hash-3"]],
  exactSourceQueue: [`SELECT id FROM roster_sync_runs WHERE source_id=? AND status IN ('queued','processing') LIMIT 1`, ["source-3"]],
  latestDispatch: [`SELECT * FROM roster_dispatches ORDER BY requested_at DESC,id DESC LIMIT 1`, []],
  creatorAccount: [`SELECT p.email, p.real_name, p.role, c.source_type, c.doctor_key, c.display_name, s.session_json FROM account_profiles p LEFT JOIN account_claims c ON c.email = p.email LEFT JOIN account_states s ON s.email = p.email WHERE p.email = ? ORDER BY c.source_type, c.display_name`, ["rhaydon@gmail.com"]],
  affectedRosterClaim: [`SELECT p.email FROM account_claims c INDEXED BY idx_account_claims_source_doctor_email INNER JOIN account_profiles p ON p.email=c.email WHERE c.source_type=? AND c.doctor_key=? LIMIT 41`, ["mmc", "MMC DOCTOR 001"]],
  creatorSnapshotRegistry: [`SELECT status, built_revision FROM snapshot_registry WHERE owner_type = ? AND owner_id = ? AND doctor_key = ? AND range_key = ?`, ["creator-account", "rhaydon@gmail.com", "RICHARD HAYDON", "2026-01-01:2026-12-31"]],
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
report.fixture = { files: sources.length, doctors: sources.length * doctorsPerSource, events: totalEvents, accounts: 10000, accountClaims: 20000, syncRuns: 10000, dispatches: 10000, inactiveRetainedSummaries: 150 };
report.estimates = {
  staff: "candidate membership rows plus repeated event-index probes; grows with files/membership and covered history",
  coverage: `${doctorsPerSource * days} ED event rows examined to return one aggregate row`,
  onShift: `${doctorsPerSource} ED/date event rows examined and returned`,
  access: `${days} indexed rows for one doctor before range filtering`,
  contacts: "at most eight returned, but the existing endpoint repeats authentication/access and R2 discovery on every poll",
  compactCoverage: "one compact row per contributing active file; zero roster-event rows",
  compactStaff: "one compact contribution per doctor/file for the selected ED/term; zero roster-event rows",
  bootstrapEvents: "at most 25,001 rows from one exact file-index walk, with no database sort",
  compactRosterStatus: "at most 100 compact rows; zero roster-event rows and independent of roster history",
  expectedRosterStatus: "at most 100 primary-key probes for explicitly expected files",
  boundedSources: "at most 16 rows from an ordered source index walk",
  latestSourceRun: "one row from an exact source/latest-run index probe",
  exactProviderVersion: "one exact source/version index probe plus one raw-file primary-key lookup; independent of other sources and versions",
  exactContentHash: "one exact source/hash/status index probe; independent of other source history",
  exactSourceQueue: "one exact source/status queue probe; independent of other sources",
  latestDispatch: "one row from the requested-at dispatch index",
  creatorAccount: "one exact account primary-key probe plus only that account's claim and state rows; independent of the other 9,999 accounts and 109,200 roster events",
  affectedRosterClaim: "only claims for one exact ED/doctor pair; independent of all other users and roster history",
  creatorSnapshotRegistry: "one exact composite-primary-key probe; independent of roster and account history",
};

assert.equal(totalEvents, 109200);
assert.ok(report.onShift.plan.some((line) => /SEARCH e USING INDEX idx_events_(?:date_source|source_range).*(?:start_date|source_type)/.test(line)), "On shift must remain an indexed ED/date lookup");
assert.ok(report.access.plan.some((line) => /SEARCH e USING INDEX idx_events_doctor_range/.test(line)), "access must use its doctor/date index");
assert.ok(report.contacts.plan.some((line) => /SEARCH contact_list_files USING INDEX idx_contacts_source_received/.test(line)), "contact discovery must use its source/revision index");
assert.ok(report.coverage.plan.some((line) => /SEARCH e USING/.test(line)), "baseline coverage plan should be recorded as a broad source-history index walk");
assert.ok(report.compactCoverage.plan.every((line) => !/roster_events/.test(line)), "compact coverage must not access roster events");
assert.ok(report.compactStaff.plan.every((line) => !/roster_events/.test(line)), "compact staff must not access roster events");
assert.ok(report.bootstrapEvents.plan.some((line) => /SEARCH roster_events USING INDEX idx_events_file \(file_id=\?\)/.test(line)), "bootstrap must use the exact file index");
assert.equal(report.bootstrapEvents.plan.some((line) => /SCAN roster_events|TEMP B-TREE/i.test(line)), false, "bootstrap must not scan or sort roster history before its sentinel limit");
assert.ok(report.compactRosterStatus.plan.some((line) => /idx_compact_roster_status/.test(line)), "roster status must use its compact active index");
assert.equal(report.compactRosterStatus.plan.some((line) => /TEMP B-TREE|roster_events/i.test(line)), false, "roster status must not sort or inspect roster history");
assert.ok(report.expectedRosterStatus.plan.some((line) => /sqlite_autoindex_roster_file_status_summaries_1/.test(line)), "expected roster status must use primary-key probes");
assert.ok(report.boundedSources.plan.some((line) => /idx_roster_sources_label_id/.test(line)), "source status must use its bounded ordering index");
assert.ok(report.latestSourceRun.plan.some((line) => /idx_roster_sync_runs_source_started_id/.test(line)), "latest source run must use the exact source/history index");
assert.equal(report.latestSourceRun.plan.some((line) => /TEMP B-TREE|SCAN roster_sync_runs/i.test(line)), false, "latest source run must not scan or sort run history");
assert.ok(report.exactProviderVersion.plan.some((line) => /idx_roster_sync_runs_source_version_status_started/.test(line)), "provider-version idempotency must use the exact source/version index");
assert.equal(report.exactProviderVersion.plan.some((line) => /SCAN roster_sync_runs/i.test(line)), false, "provider-version idempotency must not scan sync history");
assert.ok(report.exactContentHash.plan.some((line) => /idx_roster_sync_runs_source_hash/.test(line)), "content idempotency must use the exact source/hash index");
assert.equal(report.exactContentHash.plan.some((line) => /SCAN roster_sync_runs/i.test(line)), false, "content idempotency must not scan sync history");
assert.ok(report.exactSourceQueue.plan.some((line) => /idx_roster_sync_runs_source_status_started/.test(line)), "queue polling must use the exact source/status index");
assert.equal(report.exactSourceQueue.plan.some((line) => /SCAN roster_sync_runs/i.test(line)), false, "queue polling must not scan global sync history");
assert.ok(report.latestDispatch.plan.some((line) => /idx_roster_dispatches_requested_id/.test(line)), "latest dispatch must use its history index");
assert.equal(report.latestDispatch.plan.some((line) => /TEMP B-TREE/i.test(line)), false, "latest dispatch must use its ordered index without a temporary sort");
assert.ok(report.creatorAccount.plan.some((line) => /SEARCH p USING INDEX sqlite_autoindex_account_profiles_1 \(email=\?\)/i.test(line)), "Creator authentication must use the account email primary key");
assert.ok(report.creatorAccount.plan.some((line) => /SEARCH c USING INDEX sqlite_autoindex_account_claims_1 \(email=\?\)/i.test(line)), "Creator claims must use the email-leading primary key");
assert.equal(report.creatorAccount.plan.some((line) => /roster_events|roster_file_doctors|SCAN account_profiles|SCAN account_claims/i.test(line)), false, "minimal Creator account loading must not scan account or roster history");
assert.ok(report.affectedRosterClaim.plan.some((line) => /idx_account_claims_source_doctor_email/.test(line)), "issue propagation must use the exact source/doctor claim index");
assert.equal(report.affectedRosterClaim.plan.some((line) => /SCAN account_claims|SCAN account_profiles/i.test(line)), false, "issue propagation must not scan all accounts or claims");
assert.ok(report.creatorSnapshotRegistry.plan.some((line) => /snapshot_registry.*(?:PRIMARY KEY|sqlite_autoindex_snapshot_registry_1)/i.test(line)), "Creator snapshot metadata must use its composite primary key");
assert.equal(report.creatorSnapshotRegistry.plan.some((line) => /SCAN|roster_events|roster_file_doctors/i.test(line)), false, "Creator snapshot metadata must be one exact lookup");

console.log(JSON.stringify(report, null, 2));
