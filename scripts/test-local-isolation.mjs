import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

import { requestQueuedRosterProcessing } from "../functions/_lib/automation-dispatch.js";
import { findmyshiftLastModified } from "../functions/_lib/findmyshift.js";
import {
  guardedFetch,
  localFeatureDisabledResponse,
  localOnlyEnabled,
} from "../functions/_lib/outbound-network.js";
import { onRequestPost as dispatchAutomation } from "../functions/api/automation/dispatch.js";
import { onRequestPost as checkFindmyshift } from "../functions/api/automation/findmyshift-check.js";
import { onRequestPost as ingestAutomation } from "../functions/api/automation/ingest.js";
import { onRequestPost as ingestVhhAutomation } from "../functions/api/automation/vhh-roster-extract.js";
import { onRequestGet as pollPendingAutomation } from "../functions/api/automation/pending.js";

const originalFetch = globalThis.fetch;
const attemptedUrls = [];
globalThis.fetch = async (input) => {
  const url = input instanceof Request ? input.url : String(input);
  attemptedUrls.push(url);
  return new Response(JSON.stringify({ ok: true }), {
    status: 200,
    headers: { "Content-Type": "application/json" },
  });
};

try {
  assert.equal(localOnlyEnabled({ LOCAL_ONLY: "true" }), true);
  assert.equal(localOnlyEnabled({ LOCAL_ONLY: "false" }), false);

  await assert.rejects(
    guardedFetch({ LOCAL_ONLY: "true" }, "https://api.github.com/example", {}, { label: "GitHub test" }),
    (error) => error?.code === "LOCAL_OUTBOUND_BLOCKED" && /No external service was contacted/.test(error.message),
  );
  await assert.rejects(
    guardedFetch({ LOCAL_ONLY: "true" }, "https://api.postmarkapp.com/email", {}, { label: "Postmark test" }),
    (error) => error?.code === "LOCAL_OUTBOUND_BLOCKED",
  );
  await assert.rejects(
    findmyshiftLastModified("synthetic-key", "synthetic-team", { env: { LOCAL_ONLY: "true" } }),
    (error) => error?.code === "LOCAL_OUTBOUND_BLOCKED",
  );
  assert.deepEqual(attemptedUrls, [], "blocked external requests must not reach fetch");

  await guardedFetch({ LOCAL_ONLY: "true" }, "http://127.0.0.1:9876/internal-test");
  assert.deepEqual(attemptedUrls, ["http://127.0.0.1:9876/internal-test"], "loopback requests remain available locally");
  attemptedUrls.length = 0;

  const inaccessibleBinding = new Proxy({}, {
    get() { throw new Error("A disabled local endpoint touched D1 or R2."); },
  });
  const localEnv = {
    LOCAL_ONLY: "true",
    ROSTER_AUTOMATION_ENABLED: "false",
    ROSTER_AUTOMATION_WRITES_ENABLED: "false",
    EMAIL_DELIVERY_ENABLED: "false",
    ROSTER_AUTOMATION_TOKEN: "synthetic-automation-token",
    ROSTER_WATCHDOG_TOKEN: "synthetic-watchdog-token",
    VHH_AUTOMATION_TOKEN: "synthetic-vhh-token",
    FINDMYSHIFT_API_KEY: "synthetic-findmyshift-key",
    FINDMYSHIFT_TEAM_ID: "synthetic-team-id",
    GITHUB_ACTIONS_TOKEN: "synthetic-github-token",
    POSTMARK_API_TOKEN: "synthetic-postmark-token",
    ROSTER_DB: inaccessibleBinding,
    ROSTER_FILES: inaccessibleBinding,
  };

  const directDispatch = await requestQueuedRosterProcessing({
    ...localEnv,
    ROSTER_AUTOMATION_WRITES_ENABLED: "true",
    ROSTER_AUTOMATION_QUEUE_ENABLED: "true",
    ROSTER_AUTOMATION_SOURCE_ALLOWLIST: "monash-adults",
  }, { sourceId: "monash-adults", reason: "isolation-test" });
  assert.deepEqual(directDispatch, { ok: false, dispatched: false, reason: "local-disabled", dispatch: null });

  const disabledResponse = localFeatureDisabledResponse(localEnv, "Email invitations");
  assert.equal(disabledResponse.status, 503);
  assert.match((await disabledResponse.json()).error, /disabled in local development.*No external service was contacted/i);

  const endpointCases = [
    ["dispatch", dispatchAutomation, "/api/automation/dispatch", "synthetic-watchdog-token"],
    ["FindMyShift", checkFindmyshift, "/api/automation/findmyshift-check", "synthetic-automation-token"],
    ["ingestion", ingestAutomation, "/api/automation/ingest", "synthetic-automation-token"],
    ["VHH", ingestVhhAutomation, "/api/automation/vhh-roster-extract", "synthetic-vhh-token"],
  ];
  for (const [label, handler, pathname, token] of endpointCases) {
    const response = await handler({
      request: new Request(`http://127.0.0.1:9876${pathname}`, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: "{}",
      }),
      env: localEnv,
    });
    const payload = await response.json();
    assert.equal(response.status, 503, `${label} must fail closed locally`);
    assert.equal(payload.status, "local-disabled", `${label} must explain that the local block is deliberate`);
    assert.match(payload.error, /No external service was contacted/i);
  }
  const pendingResponse = await pollPendingAutomation({
    request: new Request("http://127.0.0.1:9876/api/automation/pending", {
      headers: { Authorization: "Bearer synthetic-automation-token" },
    }),
    env: localEnv,
  });
  assert.equal(pendingResponse.status, 503, "queue polling must fail closed locally");
  assert.equal((await pendingResponse.json()).status, "local-disabled");
  assert.deepEqual(attemptedUrls, [], "disabled automation endpoints must not reach fetch");

  const localDevSource = await readFile(new URL("./local-dev.mjs", import.meta.url), "utf8");
  assert.match(localDevSource, /"--persist-to", LOCAL_STATE_DIRECTORY/);
  assert.match(localDevSource, /"d1", "migrations", "apply", "ROSTER_DB", "--local"/);
  assert.match(localDevSource, /"d1", "execute", "ROSTER_DB", "--local"/);
  assert.match(localDevSource, /"--binding", "LOCAL_ONLY=true"/);
  assert.match(localDevSource, /"--binding", "FACILITY_OVERVIEW_MAINTENANCE_MODE=false"/);
  assert.match(localDevSource, /"--binding", "IDENTITY_DISCOVERY_ENABLED=false"/);
  assert.match(localDevSource, /"--binding", "ACCOUNT_SNAPSHOT_BUILD_ENABLED=false"/);
  assert.match(localDevSource, /"--binding", "EMAIL_DELIVERY_ENABLED=false"/);
  assert.doesNotMatch(localDevSource, /localDevArguments[\s\S]{0,800}"--remote"/);

  const serverSources = await Promise.all([
    "../functions/api/state.js",
    "../functions/api/automation/findmyshift-check.js",
    "../functions/_lib/automation-dispatch.js",
    "../functions/_lib/findmyshift.js",
    "./process-roster-queue.mjs",
  ].map((source) => readFile(new URL(source, import.meta.url), "utf8")));
  const rawFetches = serverSources.flatMap((source) => [...source.matchAll(/\bfetch\s*\(/g)]);
  assert.equal(rawFetches.length, 0, "server integrations must use the shared outbound guard instead of raw fetch");
  const stateSource = serverSources[0];
  const inviteAction = stateSource.slice(stateSource.indexOf('action === "adminSendInvite"'), stateSource.indexOf('action === "resolveAccountClaims"'));
  assert.match(inviteAction, /localFeatureDisabledResponse\(context\.env, "Email invitations"\)/);
  assert.ok(
    inviteAction.indexOf("localFeatureDisabledResponse") < inviteAction.indexOf("upsertAccountMirror"),
    "local email blocking must happen before invite account or D1 writes",
  );
  assert.match(inviteAction, /guardedFetch\(context\.env, "https:\/\/api\.postmarkapp\.com\/email"/);
} finally {
  globalThis.fetch = originalFetch;
}

console.log("Local isolation checks passed: non-loopback network, D1, R2, automation, provider, workflow, and email paths are fail-closed.");
