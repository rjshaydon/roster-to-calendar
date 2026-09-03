import assert from "node:assert/strict";

import { onRequestPost as handleDerivedRoster } from "../functions/api/automation/derived.js";

const token = "queue-test-token";
const run = {
  id: "sync:future-source:1",
  source_id: "future-source",
  trigger_type: "scheduled",
  provider_version: "future-version",
  content_hash: "future-hash",
  file_id: "future-file",
  source_file_id: "future-file",
  status: "queued",
  message: "Queued.",
  doctor_count: 0,
  event_count: 0,
  started_at: "2026-08-29T00:00:00.000Z",
  completed_at: "",
};

function derivedRequest(phase) {
  return new Request("https://example.test/api/automation/derived", {
    method: "POST",
    headers: { authorization: `Bearer ${token}`, "content-type": "application/json" },
    body: JSON.stringify({
      runId: run.id,
      sourceId: run.source_id,
      phase,
      file: { id: run.file_id, name: "Future Roster.xlsx", sourceId: run.source_id, sourceType: "future" },
      message: "Processor did not recognise this source.",
    }),
  });
}

class QueueFailureDb {
  constructor(initialRun) {
    this.run = { ...initialRun };
  }

  prepare(sql) {
    return new QueueFailureStatement(this, sql);
  }
}

class QueueFailureStatement {
  constructor(db, sql) {
    this.db = db;
    this.sql = String(sql || "").replace(/\s+/g, " ").trim();
    this.args = [];
  }

  bind(...args) {
    this.args = args;
    return this;
  }

  async first() {
    if (this.sql.includes("sqlite_master")) {
      return {
        has_account_invites: 1,
        has_staff_overrides: 1,
        has_roster_dispatches: 1,
        has_contact_list_files: 1,
        has_contact_resolutions: 1,
        has_contact_resolution_history: 1,
      };
    }
    if (this.sql === "SELECT * FROM roster_sync_runs WHERE id = ?") {
      return this.args[0] === this.db.run.id ? { ...this.db.run } : null;
    }
    if (this.sql === "SELECT source_type FROM roster_files WHERE id = ?") return null;
    return null;
  }

  async all() {
    return { results: [] };
  }

  async run() {
    if (this.sql.startsWith("UPDATE roster_sync_runs SET status = ?, message = ?")) {
      const [status, message, fileId, doctorCount, eventCount, completedAt, runId] = this.args;
      if (runId === this.db.run.id) {
        Object.assign(this.db.run, {
          status,
          message,
          file_id: fileId,
          doctor_count: doctorCount,
          event_count: eventCount,
          completed_at: completedAt,
        });
      }
    }
    return { success: true, meta: { changes: 1 } };
  }
}

const failedDb = new QueueFailureDb(run);
const failedResponse = await handleDerivedRoster({
  request: derivedRequest("failed"),
  env: { ROSTER_AUTOMATION_TOKEN: token, ROSTER_AUTOMATION_WRITES_ENABLED: "true", ROSTER_DB: failedDb },
});
assert.equal(failedResponse.status, 200, "an unrecognised queued source must still accept terminal failure reporting");
assert.equal((await failedResponse.json()).phase, "failed");
assert.equal(failedDb.run.status, "failed", "failure reporting must remove the run from the active queue");
assert.match(failedDb.run.message, /processor did not recognise this source/i);
assert.ok(failedDb.run.completed_at, "a failed run must receive a completion timestamp");

const processingDb = new QueueFailureDb(run);
const processingResponse = await handleDerivedRoster({
  request: derivedRequest("start"),
  env: { ROSTER_AUTOMATION_TOKEN: token, ROSTER_AUTOMATION_WRITES_ENABLED: "true", ROSTER_DB: processingDb },
});
assert.equal(processingResponse.status, 400, "an unrecognised source must not be allowed to save derived roster data");
assert.equal(processingDb.run.status, "queued", "rejecting derived data must not falsely mark the run successful");

console.log("Roster queue failure safeguards passed.");
