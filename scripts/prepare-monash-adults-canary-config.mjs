import { chmod, readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";

const args = process.argv.slice(2);
const outputPath = resolve(args[0] || "/private/tmp/wrangler.monash-adults-canary.toml");
const reviewedContentSha256 = String(args[1] || "").trim().toLowerCase();
if (!/^[a-f0-9]{64}$/.test(reviewedContentSha256)) {
  throw new Error("A reviewed 64-character SHA-256 is required as the second argument.");
}
const sourcePath = resolve("wrangler.toml");
const source = await readFile(sourcePath, "utf8");
const previewMarker = "[env.preview.vars]";
const markerIndex = source.indexOf(previewMarker);
if (markerIndex < 0) throw new Error("Preview configuration marker is missing.");

const production = source.slice(0, markerIndex);
const preview = source.slice(markerIndex);
const requiredClosedSettings = new Map([
  ["ROSTER_AUTOMATION_WRITES_ENABLED", "false"],
  ["ROSTER_AUTOMATION_QUEUE_ENABLED", "false"],
  ["ROSTER_AUTOMATION_SOURCE_ALLOWLIST", ""],
  ["ROSTER_AUTOMATION_REVIEWED_FACT_LIMIT", "0"],
  ["ROSTER_AUTOMATION_REVIEWED_CONTENT_SHA256", ""],
  ["MANUAL_ROSTER_WRITES_ENABLED", "false"],
  ["ROSTER_ADVANCED_MAINTENANCE_ENABLED", "false"],
  ["CONTACT_AUTOMATION_WRITES_ENABLED", "false"],
  ["FACILITY_BOOTSTRAP_INSPECTION_ENABLED", "false"],
  ["FACILITY_BOOTSTRAP_EXECUTION_ENABLED", "false"],
  ["ACCOUNT_SNAPSHOT_BUILD_ENABLED", "false"],
  ["IDENTITY_DISCOVERY_ENABLED", "false"],
  ["ROSTER_INSIGHT_READS_ENABLED", "false"],
]);
for (const [key, value] of requiredClosedSettings) {
  const pattern = new RegExp(`^${key} = "${value}"$`, "m");
  if (!pattern.test(production)) throw new Error(`Production ${key} is not at its required closed value.`);
}

let canary = production
  .replace(/^pages_build_output_dir = "public"$/m, `pages_build_output_dir = ${JSON.stringify(resolve("public"))}`)
  .replace(/^ROSTER_AUTOMATION_WRITES_ENABLED = "false"$/m, 'ROSTER_AUTOMATION_WRITES_ENABLED = "true"')
  .replace(/^ROSTER_AUTOMATION_QUEUE_ENABLED = "false"$/m, 'ROSTER_AUTOMATION_QUEUE_ENABLED = "true"')
  .replace(/^ROSTER_AUTOMATION_SOURCE_ALLOWLIST = ""$/m, 'ROSTER_AUTOMATION_SOURCE_ALLOWLIST = "monash-adults"')
  .replace(/^ROSTER_AUTOMATION_REVIEWED_FACT_LIMIT = "0"$/m, 'ROSTER_AUTOMATION_REVIEWED_FACT_LIMIT = "1250"')
  .replace(/^ROSTER_AUTOMATION_REVIEWED_CONTENT_SHA256 = ""$/m, `ROSTER_AUTOMATION_REVIEWED_CONTENT_SHA256 = "${reviewedContentSha256}"`);
canary += preview;

for (const expected of [
  'ROSTER_AUTOMATION_WRITES_ENABLED = "true"',
  'ROSTER_AUTOMATION_QUEUE_ENABLED = "true"',
  'ROSTER_AUTOMATION_SOURCE_ALLOWLIST = "monash-adults"',
  'ROSTER_AUTOMATION_REVIEWED_FACT_LIMIT = "1250"',
  `ROSTER_AUTOMATION_REVIEWED_CONTENT_SHA256 = "${reviewedContentSha256}"`,
]) {
  if (!canary.slice(0, canary.indexOf(previewMarker)).includes(expected)) throw new Error(`Canary setting was not generated: ${expected}`);
}
if (!preview.includes('ROSTER_AUTOMATION_WRITES_ENABLED = "false"')
  || !preview.includes('ROSTER_AUTOMATION_QUEUE_ENABLED = "false"')
  || !preview.includes('ROSTER_AUTOMATION_SOURCE_ALLOWLIST = ""')) {
  throw new Error("Preview roster automation is not closed.");
}

await writeFile(outputPath, canary, { encoding: "utf8", mode: 0o600 });
await chmod(outputPath, 0o600);
process.stdout.write(`${JSON.stringify({
  mode: "prepare-only",
  outputPath,
  productionChanges: {
    ROSTER_AUTOMATION_WRITES_ENABLED: "true",
    ROSTER_AUTOMATION_QUEUE_ENABLED: "true",
    ROSTER_AUTOMATION_SOURCE_ALLOWLIST: "monash-adults",
    ROSTER_AUTOMATION_REVIEWED_FACT_LIMIT: "1250",
    ROSTER_AUTOMATION_REVIEWED_CONTENT_SHA256: reviewedContentSha256,
  },
  unchangedSafetyControlsVerified: requiredClosedSettings.size - 3,
  previewRemainsClosed: true,
}, null, 2)}\n`);
