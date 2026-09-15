import assert from "node:assert/strict";
import { execFileSync, spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const script = fileURLToPath(new URL("./submit-monash-adults-canary.mjs", import.meta.url));
const fixture = fileURLToPath(new URL("../fixtures/AdultTerm1.2026.xlsx", import.meta.url));
const bytes = readFileSync(fixture);
const hash = createHash("sha256").update(bytes).digest("hex");
const baseArgs = [script, "--file", fixture, "--provider-version", "fixture-1", "--provider-modified-at", "2026-01-01T00:00:00Z"];

const prepared = JSON.parse(execFileSync(process.execPath, baseArgs, { encoding: "utf8" }));
assert.equal(prepared.mode, "prepare-only");
assert.equal(prepared.sourceId, "monash-adults");
assert.equal(prepared.submittedFileName, "AdultTerm3.2026.xlsx");
assert.equal(prepared.endpoint, "https://roster-to-calendar.pages.dev/api/automation/ingest");
assert.equal(prepared.sha256, hash);

const automationImportSource = readFileSync(fileURLToPath(new URL("../functions/_lib/automation-import.js", import.meta.url)), "utf8");
assert.match(automationImportSource, /contentHash: String\(contentHash \|\| ""\)\.toLowerCase\(\)/, "the reviewed workbook hash must survive parsing and reach the derived-save guard");

const missingConfirmation = spawnSync(process.execPath, [...baseArgs, "--execute", "--expected-sha256", hash], {
  encoding: "utf8",
  env: { ...process.env, ROSTER_AUTOMATION_TOKEN: "synthetic-never-sent" },
});
assert.notEqual(missingConfirmation.status, 0);
assert.match(missingConfirmation.stderr, /Execution requires --confirm SUBMIT_MONASH_ADULTS_ONCE/);

const wrongHash = spawnSync(process.execPath, [...baseArgs, "--execute", "--expected-sha256", "0".repeat(64), "--confirm", "SUBMIT_MONASH_ADULTS_ONCE"], {
  encoding: "utf8",
  env: { ...process.env, ROSTER_AUTOMATION_TOKEN: "synthetic-never-sent" },
});
assert.notEqual(wrongHash.status, 0);
assert.match(wrongHash.stderr, /does not match/);

console.log("Monash Adults one-shot canary preparation and fail-closed execution checks passed.");
