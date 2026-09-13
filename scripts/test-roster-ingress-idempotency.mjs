import assert from "node:assert/strict";
import { readFile, readdir } from "node:fs/promises";
import { DatabaseSync } from "node:sqlite";
import { sha256Hex } from "../functions/_lib/automation-import.js";
import { normaliseVhhRosterExtract } from "../functions/_lib/vhh-roster.js";
import { onRequestPost as ingestRoster } from "../functions/api/automation/ingest.js";
import { onRequestPost as ingestVhhRoster } from "../functions/api/automation/vhh-roster-extract.js";
import { onRequestPost as checkFindmyshift } from "../functions/api/automation/findmyshift-check.js";
import { onRequestPost as saveDerivedRoster } from "../functions/api/automation/derived.js";

class LocalD1 {
  constructor(sqlite) { this.sqlite = sqlite; this.rowsWritten = 0; this.statements = []; }
  prepare(sql) {
    const owner = this;
    owner.statements.push(sql.replace(/\s+/g, " ").trim());
    return {
      args: [],
      bind(...args) { this.args = args; return this; },
      async run() {
        const result = owner.sqlite.prepare(sql).run(...this.args);
        owner.rowsWritten += Number(result.changes || 0);
        return { success: true, meta: { changes: Number(result.changes || 0) } };
      },
      async all() { return { success: true, results: owner.sqlite.prepare(sql).all(...this.args) }; },
      async first() { return owner.sqlite.prepare(sql).get(...this.args) || null; },
    };
  }
  async batch(statements) { return Promise.all(statements.map((statement) => statement.run())); }
  resetMetrics() { this.rowsWritten = 0; this.statements = []; }
}

class LocalR2 {
  constructor() { this.puts = 0; this.objects = new Map(); }
  async put(key, value) { this.puts += 1; this.objects.set(key, value); }
}

const sqlite = new DatabaseSync(":memory:");
for (const name of (await readdir(new URL("../migrations", import.meta.url))).filter((name) => name.endsWith(".sql")).sort()) {
  sqlite.exec(await readFile(new URL(`../migrations/${name}`, import.meta.url), "utf8"));
}
// Production already has these compatibility columns. The numbered legacy
// migrations predate them, so add them to this isolated fresh-store fixture.
for (const sql of [
  "ALTER TABLE raw_roster_files ADD COLUMN name TEXT NOT NULL DEFAULT ''",
  "ALTER TABLE raw_roster_files ADD COLUMN source_type TEXT NOT NULL DEFAULT ''",
  "ALTER TABLE raw_roster_files ADD COLUMN size INTEGER NOT NULL DEFAULT 0",
  "ALTER TABLE raw_roster_files ADD COLUMN last_modified INTEGER NOT NULL DEFAULT 0",
]) sqlite.exec(sql);
const db = new LocalD1(sqlite);
const r2 = new LocalR2();
const token = "local-automation-token";
const baseEnv = {
  ROSTER_DB: db,
  ROSTER_FILES: r2,
  ROSTER_AUTOMATION_TOKEN: token,
  ROSTER_AUTOMATION_WRITES_ENABLED: "true",
  ROSTER_AUTOMATION_QUEUE_ENABLED: "false",
};

function seedSuccessfulRun({ sourceId, sourceType, fileName, providerVersion, contentHash, suffix }) {
  const fileId = `raw-${suffix}`;
  sqlite.prepare(`INSERT INTO roster_sources
    (id, provider, source_type, label, enabled, config_json, cursor_json, provider_version, last_success_at)
    VALUES (?, 'test', ?, ?, 1, '{}', '{}', ?, '2026-09-01T00:00:00Z')`).run(sourceId, sourceType, sourceId, providerVersion);
  sqlite.prepare(`INSERT INTO raw_roster_files
    (file_id, name, source_type, size, last_modified, object_key, type, data_url, uploaded_at)
    VALUES (?, ?, ?, 1, 1, ?, 'application/octet-stream', '', '2026-09-01T00:00:00Z')`).run(fileId, fileName, sourceType, `automation/${sourceId}/${suffix}`);
  sqlite.prepare(`INSERT INTO roster_sync_runs
    (id, source_id, trigger_type, provider_version, content_hash, file_id, source_file_id, status, started_at, completed_at)
    VALUES (?, ?, 'test', ?, ?, ?, ?, 'success', '2026-09-01T00:00:00Z', '2026-09-01T00:01:00Z')`)
    .run(`run-${suffix}`, sourceId, providerVersion, contentHash, fileId, fileId);
  return fileId;
}

async function jsonPayload(response, expectedStatus = 200) {
  const payload = await response.json();
  assert.equal(response.status, expectedStatus, JSON.stringify(payload));
  return payload;
}

const workbookBytes = new TextEncoder().encode("deterministic synthetic workbook");
const workbookHash = await sha256Hex(workbookBytes);
seedSuccessfulRun({ sourceId: "monash-adults", sourceType: "mmc", fileName: "Adult.xlsx", providerVersion: "etag-1", contentHash: workbookHash, suffix: "mmc" });

async function callWorkbook(providerVersion, bytes = workbookBytes) {
  return ingestRoster({
    request: new Request("http://local/api/automation/ingest", {
      method: "POST",
      headers: { authorization: `Bearer ${token}`, "content-type": "application/json" },
      body: JSON.stringify({ sourceId: "monash-adults", fileName: "Adult.xlsx", providerVersion, contentBase64: Buffer.from(bytes).toString("base64") }),
    }),
    env: { ...baseEnv, ROSTER_AUTOMATION_SOURCE_ALLOWLIST: "monash-adults" },
  });
}

db.resetMetrics();
let putsBefore = r2.puts;
assert.equal((await jsonPayload(await callWorkbook("etag-1"))).status, "unchanged");
assert.equal(db.rowsWritten, 0, "an identical workbook provider version must write zero D1 rows");
assert.equal(r2.puts, putsBefore, "an identical workbook provider version must write zero R2 objects");

db.resetMetrics();
putsBefore = r2.puts;
assert.equal((await jsonPayload(await callWorkbook("etag-2"))).status, "unchanged");
assert.equal(db.rowsWritten, 0, "identical workbook content under a new provider version must write zero D1 rows");
assert.equal(r2.puts, putsBefore, "identical workbook content under a new provider version must write zero R2 objects");

const changedWorkbookBytes = new TextEncoder().encode("deterministic synthetic workbook with one correction");
db.resetMetrics();
putsBefore = r2.puts;
assert.equal((await jsonPayload(await callWorkbook("etag-3", changedWorkbookBytes), 202)).status, "queued");
assert.ok(db.rowsWritten <= 4, `a changed workbook ingress wrote ${db.rowsWritten} D1 rows`);
assert.equal(r2.puts - putsBefore, 1, "a changed workbook ingress must retain exactly one R2 object");
db.resetMetrics();
putsBefore = r2.puts;
assert.equal((await jsonPayload(await callWorkbook("etag-3", changedWorkbookBytes), 202)).status, "queued");
assert.equal(db.rowsWritten, 0, "a repeated queued workbook version must write zero D1 rows");
assert.equal(r2.puts, putsBefore, "a repeated queued workbook version must write zero R2 objects");

const vhhExtract = {
  sourceId: "vhh-active-medical-roster",
  fileName: "Active Medical Roster.json",
  providerVersion: "vhh-etag-1",
  providerModifiedAt: "2026-09-01T00:00:00Z",
  blocks: [{
    sheetName: "Roster", blockIndex: 1, visible: true,
    dates: [{ sourceColumn: "B", displayedDate: "1-Sep", date: "2026-09-01" }],
    rows: [{ sourceRow: 2, sourceShiftLabel: "AM REG", shiftLabel: "AM REG", assignments: [{ date: "2026-09-01", displayedDate: "1-Sep", namesText: "Doctor, Test", sourceCell: "B2" }] }],
  }],
};
const normalizedVhh = normaliseVhhRosterExtract(vhhExtract);
const vhhHash = await sha256Hex(new TextEncoder().encode(JSON.stringify({ ...normalizedVhh, providerVersion: "", providerModifiedAt: "" })));
seedSuccessfulRun({ sourceId: vhhExtract.sourceId, sourceType: "vhh", fileName: normalizedVhh.fileName, providerVersion: vhhExtract.providerVersion, contentHash: vhhHash, suffix: "vhh" });
async function callVhh(payload) {
  return ingestVhhRoster({
    request: new Request("http://local/api/automation/vhh-roster-extract", { method: "POST", headers: { authorization: `Bearer ${token}`, "content-type": "application/json" }, body: JSON.stringify(payload) }),
    env: { ...baseEnv, ROSTER_AUTOMATION_SOURCE_ALLOWLIST: vhhExtract.sourceId },
  });
}

db.resetMetrics();
putsBefore = r2.puts;
assert.equal((await jsonPayload(await callVhh(vhhExtract))).status, "unchanged");
assert.equal(db.rowsWritten, 0, "an identical VHH provider version must write zero D1 rows");
assert.equal(r2.puts, putsBefore, "an identical VHH provider version must write zero R2 objects");

db.resetMetrics();
putsBefore = r2.puts;
assert.equal((await jsonPayload(await callVhh({ ...vhhExtract, providerVersion: "vhh-etag-2" }))).status, "unchanged");
assert.equal(db.rowsWritten, 0, "identical VHH content under a new provider version must write zero D1 rows");
assert.equal(r2.puts, putsBefore, "identical VHH content under a new provider version must write zero R2 objects");

const changedVhhExtract = structuredClone(vhhExtract);
changedVhhExtract.providerVersion = "vhh-etag-3";
changedVhhExtract.blocks[0].rows[0].assignments[0].namesText = "Doctor, Changed";
db.resetMetrics();
putsBefore = r2.puts;
assert.equal((await jsonPayload(await callVhh(changedVhhExtract), 202)).status, "queued");
assert.ok(db.rowsWritten <= 4, `a changed VHH ingress wrote ${db.rowsWritten} D1 rows`);
assert.equal(r2.puts - putsBefore, 1, "a changed VHH ingress must retain exactly one R2 object");
db.resetMetrics();
putsBefore = r2.puts;
assert.equal((await jsonPayload(await callVhh(changedVhhExtract))).status, "queued");
assert.equal(db.rowsWritten, 0, "a repeated queued VHH version must write zero D1 rows");
assert.equal(r2.puts, putsBefore, "a repeated queued VHH version must write zero R2 objects");

const ddhVersion = "2026-09-01T00:00:00.000Z";
const ddhRange = { from: "2026-08-03", to: "2026-11-01" };
const ddhFileName = `Dandenong-FindMyShift-stream-paired-v7-${ddhRange.from}-to-${ddhRange.to}.xlsx`;
const ddhFileId = seedSuccessfulRun({ sourceId: "dandenong-findmyshift", sourceType: "ddh", fileName: ddhFileName, providerVersion: ddhVersion, contentHash: "ddh-hash", suffix: "ddh" });
sqlite.prepare(`UPDATE roster_sources SET provider_version=?, active_file_id=?, cursor_json=? WHERE id='dandenong-findmyshift'`)
  .run(ddhVersion, ddhFileId, JSON.stringify({ findmyshiftRange: { ...ddhRange, providerVersion: ddhVersion, importFormat: "stream-paired-v7", status: "queued" } }));
const originalFetch = globalThis.fetch;
let providerRequests = 0;
globalThis.fetch = async () => { providerRequests += 1; return Response.json({ modified: ddhVersion }); };
try {
  db.resetMetrics();
  const response = await checkFindmyshift({
    request: new Request("http://local/api/automation/findmyshift-check", { method: "POST", headers: { authorization: `Bearer ${token}`, "content-type": "application/json" }, body: JSON.stringify({ range: ddhRange }) }),
    env: { ...baseEnv, ROSTER_AUTOMATION_SOURCE_ALLOWLIST: "dandenong-findmyshift", FINDMYSHIFT_API_KEY: "test-key", FINDMYSHIFT_TEAM_ID: "test-team" },
  });
  assert.equal((await jsonPayload(response)).status, "unchanged");
  assert.equal(providerRequests, 1, "an unchanged FindMyShift check should make only its version request");
  assert.equal(db.rowsWritten, 0, "an unchanged FindMyShift provider version must write zero D1 rows");
  assert.equal(r2.puts, putsBefore, "an unchanged FindMyShift provider version must write zero R2 objects");
} finally {
  globalThis.fetch = originalFetch;
}

db.resetMetrics();
const oversizedDoctors = Array.from({ length: 513 }, (_, index) => ({ key: `DOCTOR ${index}`, displayName: `Doctor ${index}` }));
const oversizedResponse = await saveDerivedRoster({
  request: new Request("http://local/api/automation/derived", {
    method: "POST",
    headers: { authorization: `Bearer ${token}`, "content-type": "application/json" },
    body: JSON.stringify({ runId: "not-loaded", sourceId: "monash-adults", phase: "complete", file: { id: "not-loaded", sourceId: "monash-adults" }, doctors: oversizedDoctors, eventsByDoctor: {} }),
  }),
  env: { ...baseEnv, ROSTER_AUTOMATION_QUEUE_ENABLED: "true", ROSTER_AUTOMATION_SOURCE_ALLOWLIST: "monash-adults" },
});
assert.equal(oversizedResponse.status, 413);
assert.equal(db.rowsWritten, 0, "an oversized derived payload must be rejected before D1 writes");
assert.equal(db.statements.length, 0, "an oversized derived payload must be rejected before any D1 statement");

console.log("Roster ingress idempotency checks passed for Monash, VHH and FindMyShift with zero unchanged D1/R2 writes.");
