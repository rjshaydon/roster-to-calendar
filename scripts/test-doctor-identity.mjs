import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  harmlessRosterIdentityKey,
  isHarmlessRosterNameVariant,
  rosterIdentityFeatures,
  scoreRosterIdentityCandidate,
  suggestedRosterPersonId,
  validateRosterPersonId,
} from "../functions/_lib/d1-calendar.js";

assert.equal(harmlessRosterIdentityKey("Toby O’Brien"), harmlessRosterIdentityKey("Toby O BRIEN"));
assert.equal(isHarmlessRosterNameVariant("Toby OBRIEN", "Toby O’Brien"), true);
assert.equal(scoreRosterIdentityCandidate({ displayName: "Toby OBRIEN" }, { displayName: "Toby O’Brien" }), null, "formatting variants are not candidate merges");

const aeshan = scoreRosterIdentityCandidate({ sourceType: "ddh", key: "AESHAN KULARATNE", displayName: "Aeshan KULARATNE" }, { sourceType: "vhh", key: "AESHAN KULURATNE", displayName: "Aeshan KULURATNE" });
assert.ok(aeshan && aeshan.score >= 55, "one-letter surname variants must be suggestions");
assert.ok(aeshan.reasons.includes("exact-given-surname-one-letter"));
assert.equal(suggestedRosterPersonId("Aeshan KULARATNE"), "person:kularatne-aeshan");
assert.equal(validateRosterPersonId("person:kularatne-aeshan"), "person:kularatne-aeshan");
assert.equal(validateRosterPersonId("person:Kularatne-Aeshan"), "");
assert.equal(rosterIdentityFeatures({ displayName: "Aeshan KULARATNE" }).surnameKey, "kularatne");

const migration = await readFile(new URL("../migrations/0026_doctor_identity_operations.sql", import.meta.url), "utf8");
for (const table of ["roster_person_redirects", "roster_identity_operations", "roster_identity_operation_items", "roster_identity_candidates", "roster_identity_features", "roster_identity_audit_runs"]) {
  assert.match(migration, new RegExp(`CREATE TABLE IF NOT EXISTS ${table}`));
}
const stateSource = await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8");
assert.match(stateSource, /action === "claimOwnRosterIdentityAlias"[\s\S]*?never accepts targetEmail[\s\S]*?SELF_ALIAS_OWNED/, "self-service alias claims must not merge another account");
assert.match(stateSource, /runIdentityAuditBatch\(context\.env\.ROSTER_DB, run\.auditRunId, \{ maxRows: 10, maxCandidates: 20, maxMs: 2000 \}\)/, "interactive duplicate checks must stay tightly bounded");
const appSource = await readFile(new URL("../public/static/app.js", import.meta.url), "utf8");
assert.match(appSource, /Link another roster spelling[\s\S]*?Names linked to other accounts are excluded/, "the self-service control must explain its account boundary");
assert.match(appSource, /Choose one hospital first[\s\S]*?sourceTypes: \[sourceType\]/, "duplicate checks must require one hospital scope");
const auditWorkerSource = await readFile(new URL("../worker/identity-audit.js", import.meta.url), "utf8");
assert.doesNotMatch(auditWorkerSource, /runScheduledIdentityAudit/, "the scheduled worker must not consume D1 reads");
console.log("Doctor identity invariants passed.");
