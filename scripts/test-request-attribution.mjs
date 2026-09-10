import assert from "node:assert/strict";
import { onRequest } from "../functions/_middleware.js";

const originalLog = console.log;
const logs = [];
console.log = (value) => logs.push(String(value));

try {
  let nextCalls = 0;
  const forbidden = new Proxy({}, { get() { throw new Error("Paused route touched D1"); } });
  const pausedPaths = [
    "/api/automation/contact-list-extract",
    "/api/automation/ingest",
    "/api/automation/pending",
    "/api/automation/dispatch",
    "/api/automation/facility-bootstrap",
    "/api/automation/facility-materialize",
  ];
  for (const path of pausedPaths) {
    const blocked = await onRequest({
      request: new Request(`https://example.test${path}`, {
        method: "POST", headers: { authorization: "Bearer secret-value" },
        body: JSON.stringify({ sourceId: "private-source", contacts: [{ phone: "0400000000" }] }),
      }),
      env: { ROSTER_DB: forbidden, CONTACT_AUTOMATION_WRITES_ENABLED: "false", ROSTER_AUTOMATION_WRITES_ENABLED: "false", CF_PAGES_COMMIT_SHA: "test-commit" },
      next: async () => { nextCalls += 1; return Response.json({ ok: true }); },
    });
    assert.equal(blocked.status, 503, `${path} should stop in middleware`);
  }
  assert.equal(nextCalls, 0);

  const db = {
    prepare() {
      return {
        bind() { return this; },
        async all() { return { results: [{ ok: true }], meta: { rows_read: 3, rows_written: 0 } }; },
      };
    },
  };
  const stateContext = {
    request: new Request("https://example.test/api/state", {
      method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({ action: "loadCalendarEvents", email: "private@example.test", password: "secret-password" }),
    }),
    env: { ROSTER_DB: db, CF_PAGES_COMMIT_SHA: "test-commit" },
    next: async () => {
      await stateContext.env.ROSTER_DB.prepare("SELECT secret").all();
      return Response.json({ ok: true });
    },
  };
  const stateResponse = await onRequest(stateContext);
  assert.equal(stateResponse.status, 200);
  const joined = logs.join("\n");
  assert.match(joined, /"route":"\/api\/state"/);
  assert.match(joined, /"action":"loadCalendarEvents"/);
  assert.doesNotMatch(joined, /private@example|secret-password|secret-value|0400000000|private-source/);
  assert.match(joined, /"d1Statements":1/);
  assert.match(joined, /"d1RowsRead":3/);

  const budgetDb = {
    prepare() {
      return { async all() { return { results: [], meta: { rows_read: 0, rows_written: 0 } }; } };
    },
  };
  const budgetContext = {
    request: new Request("https://example.test/api/state", {
      method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({ action: "login" }),
    }),
    env: { ROSTER_DB: budgetDb },
    next: async () => {
      for (let index = 0; index < 33; index += 1) await budgetContext.env.ROSTER_DB.prepare("SELECT 1").all();
      return Response.json({ ok: true });
    },
  };
  const overBudget = await onRequest(budgetContext);
  assert.equal(overBudget.status, 503, "the 33rd login statement must be rejected before execution");

  await onRequest({
    request: new Request("https://example.test/api/state", {
      method: "POST", headers: { "content-type": "application/json" },
      body: JSON.stringify({ action: "do-not-log-this-secret" }),
    }),
    env: {},
    next: async () => Response.json({ error: "unsupported" }, { status: 400 }),
  });
  assert.match(logs.at(-1), /"action":"unknown-action"/);
  assert.doesNotMatch(logs.join("\n"), /do-not-log-this-secret/);
} finally {
  console.log = originalLog;
}

console.log("Request attribution and pre-handler containment tests passed.");
