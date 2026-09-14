import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import { basename, resolve } from "node:path";

const SOURCE_ID = "monash-adults";
const SUBMITTED_FILE_NAME = "AdultTerm3.2026.xlsx";
const ENDPOINT = "https://roster-to-calendar.pages.dev/api/automation/ingest";
const EXECUTION_CONFIRMATION = "SUBMIT_MONASH_ADULTS_ONCE";
const XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const MAX_BYTES = 15 * 1024 * 1024;

const options = parseArguments(process.argv.slice(2));
if (options.help) {
  process.stdout.write(usage());
  process.exit(0);
}

const filePath = resolve(required(options.file, "--file is required"));
const providerVersion = required(options.providerVersion, "--provider-version is required");
const providerModifiedAt = required(options.providerModifiedAt, "--provider-modified-at is required");
const modifiedAtMs = Date.parse(providerModifiedAt);
if (!Number.isFinite(modifiedAtMs)) throw new Error("--provider-modified-at must be an ISO-8601 timestamp.");

const bytes = await readFile(filePath);
if (!bytes.length || bytes.length > MAX_BYTES) throw new Error("Workbook is empty or exceeds the 15 MiB ingress limit.");
if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new Error("Workbook does not have an XLSX/ZIP signature.");
const sha256 = createHash("sha256").update(bytes).digest("hex");
const manifest = {
  mode: options.execute ? "execute" : "prepare-only",
  endpoint: ENDPOINT,
  sourceId: SOURCE_ID,
  submittedFileName: SUBMITTED_FILE_NAME,
  localFileName: basename(filePath),
  providerVersion,
  providerModifiedAt: new Date(modifiedAtMs).toISOString(),
  bytes: bytes.length,
  sha256,
};

if (!options.execute) {
  process.stdout.write(`${JSON.stringify(manifest, null, 2)}\n`);
  process.exit(0);
}

const expectedHash = required(options.expectedSha256, "--expected-sha256 is required for execution").toLowerCase();
if (!/^[a-f0-9]{64}$/.test(expectedHash) || expectedHash !== sha256) {
  throw new Error("Workbook SHA-256 does not match the separately reviewed expected hash.");
}
if (options.confirm !== EXECUTION_CONFIRMATION) {
  throw new Error(`Execution requires --confirm ${EXECUTION_CONFIRMATION}.`);
}
const token = String(process.env.ROSTER_AUTOMATION_TOKEN || "").trim();
if (!token) throw new Error("ROSTER_AUTOMATION_TOKEN is required only for execution.");

const response = await fetch(ENDPOINT, {
  method: "POST",
  headers: { authorization: `Bearer ${token}`, "content-type": "application/json" },
  body: JSON.stringify({
    sourceId: SOURCE_ID,
    fileName: SUBMITTED_FILE_NAME,
    contentType: XLSX_CONTENT_TYPE,
    contentBase64: bytes.toString("base64"),
    providerVersion,
    providerModifiedAt: new Date(modifiedAtMs).toISOString(),
    lastModified: modifiedAtMs,
  }),
});
let payload = null;
try { payload = await response.json(); } catch { /* Never print an HTML response body. */ }
const result = {
  ...manifest,
  httpStatus: response.status,
  ok: response.ok && payload?.ok === true,
  status: String(payload?.status || ""),
  sourceId: String(payload?.sourceId || SOURCE_ID),
  fileId: String(payload?.fileId || ""),
  runId: String(payload?.runId || ""),
  processorDispatch: payload?.processorDispatch || null,
  error: String(payload?.error || ""),
};
process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
if (!response.ok || payload?.ok !== true) process.exitCode = 2;

function required(value, message) {
  const result = String(value || "").trim();
  if (!result) throw new Error(message);
  return result;
}

function parseArguments(args) {
  const parsed = { execute: false, help: false };
  for (let index = 0; index < args.length; index += 1) {
    const key = args[index];
    if (key === "--execute") { parsed.execute = true; continue; }
    if (key === "--help") { parsed.help = true; continue; }
    if (!key.startsWith("--") || index + 1 >= args.length) throw new Error(`Invalid argument: ${key}`);
    const name = key.slice(2).replace(/-([a-z])/g, (_match, letter) => letter.toUpperCase());
    if (!["file", "providerVersion", "providerModifiedAt", "expectedSha256", "confirm"].includes(name)) {
      throw new Error(`Unknown argument: ${key}`);
    }
    parsed[name] = args[++index];
  }
  return parsed;
}

function usage() {
  return `Prepare the fixed Monash Adults one-shot roster canary.\n\n`+
    `Preparation performs no network request:\n`+
    `  npm run canary:monash-adults -- --file <xlsx> --provider-version <version> --provider-modified-at <ISO timestamp>\n\n`+
    `Execution additionally requires the reviewed SHA-256, the exact confirmation phrase, and ROSTER_AUTOMATION_TOKEN.\n`;
}
