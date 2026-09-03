import { createHash } from "node:crypto";
import { spawn, spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { lstat, mkdir, readFile, rm, unlink, writeFile } from "node:fs/promises";
import net from "node:net";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const SCRIPT_PATH = fileURLToPath(import.meta.url);
export const REPOSITORY_ROOT = path.resolve(path.dirname(SCRIPT_PATH), "..");
export const LOCAL_STATE_DIRECTORY = path.join(REPOSITORY_ROOT, ".wrangler", "local-safe");
export const LOCAL_HOST = "127.0.0.1";
export const DEFAULT_LOCAL_PORT = 8788;
export const LOCAL_CREATOR_EMAIL = "rhaydon@gmail.com";
export const LOCAL_CREATOR_PASSWORD = "local-creator-only";
export const LOCAL_USER_EMAIL = "test.doctor@example.test";
export const LOCAL_USER_PASSWORD = "local-user-only";
const LOCAL_UNCLAIMED_EMAIL = "unclaimed@example.test";
const PID_FILE = path.join(LOCAL_STATE_DIRECTORY, "dev.pid");
const WRANGLER_PATH = path.join(REPOSITORY_ROOT, "node_modules", ".bin", "wrangler");
const FORBIDDEN_LOCAL_ARGUMENTS = new Set(["--remote", "--preview"]);
const SECRET_ENVIRONMENT_KEYS = [
  "CLOUDFLARE_API_KEY",
  "CLOUDFLARE_API_TOKEN",
  "CLOUDFLARE_ACCOUNT_ID",
  "CF_API_KEY",
  "CF_API_TOKEN",
  "ROSTER_AUTOMATION_TOKEN",
  "ROSTER_WATCHDOG_TOKEN",
  "FINDMYSHIFT_API_KEY",
  "FINDMYSHIFT_TEAM_ID",
  "GITHUB_ACTIONS_TOKEN",
  "POSTMARK_SERVER_TOKEN",
];

function localPort(value = process.env.LOCAL_DEV_PORT) {
  const parsed = Number.parseInt(String(value || DEFAULT_LOCAL_PORT), 10);
  if (!Number.isInteger(parsed) || parsed < 1024 || parsed > 65535) {
    throw new Error("LOCAL_DEV_PORT must be a number between 1024 and 65535.");
  }
  return parsed;
}

function assertExactLocalStatePath() {
  const expected = path.join(REPOSITORY_ROOT, ".wrangler", "local-safe");
  if (LOCAL_STATE_DIRECTORY !== expected
    || path.basename(LOCAL_STATE_DIRECTORY) !== "local-safe"
    || path.dirname(LOCAL_STATE_DIRECTORY) !== path.join(REPOSITORY_ROOT, ".wrangler")
    || LOCAL_STATE_DIRECTORY === REPOSITORY_ROOT
    || LOCAL_STATE_DIRECTORY === os.homedir()
    || LOCAL_STATE_DIRECTORY === path.parse(LOCAL_STATE_DIRECTORY).root) {
    throw new Error("Refusing to use an unsafe local-state path.");
  }
}

function safeChildEnvironment() {
  const environment = { ...process.env };
  for (const key of SECRET_ENVIRONMENT_KEYS) delete environment[key];
  environment.LOCAL_ONLY = "true";
  environment.ROSTER_AUTOMATION_ENABLED = "false";
  environment.ROSTER_AUTOMATION_WRITES_ENABLED = "false";
  environment.EMAIL_DELIVERY_ENABLED = "false";
  environment.NO_UPDATE_NOTIFIER = "1";
  environment.WRANGLER_LOG_PATH = path.join(LOCAL_STATE_DIRECTORY, "wrangler.log");
  return environment;
}

async function assertNoDevVarsSecrets() {
  for (const fileName of [".dev.vars", ".dev.vars.local", ".env", ".env.local"]) {
    const filePath = path.join(REPOSITORY_ROOT, fileName);
    if (!existsSync(filePath)) continue;
    const contents = await readFile(filePath, "utf8");
    const secretNames = SECRET_ENVIRONMENT_KEYS.filter((key) => new RegExp(`^\\s*${key}\\s*=`, "m").test(contents));
    if (secretNames.length) {
      throw new Error(`${fileName} contains remote credentials (${secretNames.join(", ")}). Local development refuses to load it.`);
    }
  }
}

async function assertPortAvailable(port) {
  await new Promise((resolve, reject) => {
    const server = net.createServer();
    server.unref();
    server.once("error", (error) => {
      if (error?.code === "EADDRINUSE") {
        reject(new Error(`Port ${port} is already in use. Stop the existing local server before continuing.`));
        return;
      }
      reject(error);
    });
    server.listen({ host: LOCAL_HOST, port }, () => server.close(resolve));
  });
}

async function activeLocalDevPid() {
  if (!existsSync(PID_FILE)) return 0;
  const value = Number.parseInt((await readFile(PID_FILE, "utf8")).trim(), 10);
  if (!Number.isInteger(value) || value <= 1) {
    await unlink(PID_FILE).catch(() => {});
    return 0;
  }
  try {
    process.kill(value, 0);
    return value;
  } catch {
    await unlink(PID_FILE).catch(() => {});
    return 0;
  }
}

export async function checkLocalSafety(options = {}) {
  assertExactLocalStatePath();
  if (!existsSync(WRANGLER_PATH)) {
    throw new Error("Wrangler is not installed locally. Run npm install first.");
  }
  const unexpectedArguments = process.argv.slice(3).filter((argument) => FORBIDDEN_LOCAL_ARGUMENTS.has(argument));
  if (unexpectedArguments.length) throw new Error(`Remote arguments are forbidden: ${unexpectedArguments.join(", ")}`);
  const configuration = await readFile(path.join(REPOSITORY_ROOT, "wrangler.toml"), "utf8");
  if (/^\s*remote\s*=\s*true\s*$/im.test(configuration)) {
    throw new Error("wrangler.toml contains remote = true. Local development is blocked.");
  }
  await assertNoDevVarsSecrets();
  const loadedSecretNames = SECRET_ENVIRONMENT_KEYS.filter((key) => String(process.env[key] || "").trim());
  if (loadedSecretNames.length) {
    throw new Error(`Remote credentials are loaded in this shell (${loadedSecretNames.join(", ")}). Start local development from a clean shell.`);
  }
  const port = localPort(options.port);
  if (options.requirePortAvailable === true) await assertPortAvailable(port);
  return { host: LOCAL_HOST, port, stateDirectory: LOCAL_STATE_DIRECTORY };
}

function runWrangler(argumentsList, options = {}) {
  for (const argument of argumentsList) {
    if (FORBIDDEN_LOCAL_ARGUMENTS.has(argument)) throw new Error(`Refusing forbidden Wrangler argument: ${argument}`);
  }
  const result = spawnSync(WRANGLER_PATH, argumentsList, {
    cwd: REPOSITORY_ROOT,
    env: safeChildEnvironment(),
    encoding: "utf8",
    stdio: options.capture === true ? "pipe" : "inherit",
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    const details = [result.stdout, result.stderr].filter(Boolean).join("\n").trim();
    throw new Error(`Local Wrangler command failed${details ? `:\n${details}` : "."}`);
  }
  return result.stdout || "";
}

export async function resetLocalState(options = {}) {
  const { port } = await checkLocalSafety({ port: options.port });
  const activePid = await activeLocalDevPid();
  if (activePid) throw new Error(`Local development is still running as process ${activePid}. Stop it before resetting.`);
  if (options.checkPort !== false) await assertPortAvailable(port);
  assertExactLocalStatePath();
  if (existsSync(LOCAL_STATE_DIRECTORY)) {
    const info = await lstat(LOCAL_STATE_DIRECTORY);
    if (info.isSymbolicLink()) throw new Error("Refusing to remove a symbolic-link local-state directory.");
    await rm(LOCAL_STATE_DIRECTORY, { recursive: true, force: false });
  }
  await mkdir(LOCAL_STATE_DIRECTORY, { recursive: true });
  console.log(`Reset disposable local state only: ${LOCAL_STATE_DIRECTORY}`);
}

export async function migrateLocalDatabase() {
  await checkLocalSafety();
  await mkdir(LOCAL_STATE_DIRECTORY, { recursive: true });
  console.log(`Applying migrations to local D1 only: ${LOCAL_STATE_DIRECTORY}`);
  runWrangler([
    "d1", "migrations", "apply", "ROSTER_DB", "--local",
    "--persist-to", LOCAL_STATE_DIRECTORY,
  ], { capture: true });
  console.log("Local D1 migrations applied successfully.");
}

function sqlString(value) {
  return `'${String(value ?? "").replaceAll("'", "''")}'`;
}

function sqlValues(rows) {
  return rows.map((row) => `(${row.map(sqlString).join(", ")})`).join(",\n  ");
}

function passwordHash(password, salt) {
  return createHash("sha256").update(`${salt}:${password}`).digest("hex");
}

function localSeedSql() {
  const now = "2026-09-03T00:00:00.000Z";
  const doctors = [
    { sourceType: "mmc", key: "ABI THANIKASALAM", displayName: "Abi THANIKASALAM", seniority: "Consultant" },
    { sourceType: "ddh", key: "ABHI THANIKASALAM", displayName: "Abhi THANIKASALAM", seniority: "Consultant" },
    { sourceType: "mmc", key: "AESHAN KULARATNE", displayName: "Aeshan KULARATNE", seniority: "Registrar" },
    { sourceType: "ddh", key: "AESHAN KULURATNE", displayName: "Aeshan KULURATNE", seniority: "Registrar" },
    { sourceType: "vhh", key: "JAY WEERARATNE", displayName: "Jay WEERARATNE", seniority: "Consultant" },
    { sourceType: "ddh", key: "JAYANTHA WEERARATNE", displayName: "Jayantha WEERARATNE", seniority: "Consultant" },
    { sourceType: "casey", key: "TOBY O BRIEN", displayName: "Toby O BRIEN", seniority: "Registrar" },
    { sourceType: "mch", key: "TOBY OBRIEN", displayName: "Toby OBRIEN", seniority: "Registrar" },
    { sourceType: "mmc", key: "MAYA PATEL", displayName: "Maya PATEL", seniority: "Consultant" },
    { sourceType: "vhh", key: "SAM LEE", displayName: "Sam LEE", seniority: "Registrar" },
    { sourceType: "ddh", key: "SAM LI", displayName: "Sam LI", seniority: "Registrar" },
  ];
  const files = ["mmc", "ddh", "casey", "mch", "vhh"].map((sourceType) => ({
    id: `local-${sourceType}`,
    name: `Local ${sourceType.toUpperCase()} roster.xlsx`,
    sourceType,
  }));
  const events = doctors.map((doctor, index) => {
    const fileId = `local-${doctor.sourceType}`;
    const day = String(7 + (index % 5)).padStart(2, "0");
    const id = `${fileId}:${doctor.key.toLowerCase().replace(/[^a-z0-9]+/g, "-")}`;
    const event = {
      id,
      fileId,
      sourceType: doctor.sourceType,
      doctorKey: doctor.key,
      displayName: doctor.displayName,
      startDate: `2026-09-${day}`,
      endDate: `2026-09-${day}`,
      startTime: "08:00",
      endTime: "16:00",
      title: "Day shift",
      rawValue: "D",
      seniority: doctor.seniority,
      location: `${doctor.sourceType.toUpperCase()} Emergency Department`,
      allDay: false,
      timeLabel: "08:00–16:00",
    };
    return { ...doctor, ...event, eventJson: JSON.stringify(event) };
  });
  const creatorSalt = "local-creator-salt-v1";
  const userSalt = "local-user-salt-v1";
  const unclaimedSalt = "local-unclaimed-salt-v1";
  const accountRows = [
    [LOCAL_CREATOR_EMAIL, "Local Creator", "creator", "1", "1", "0", "0", "local-creator-feed", creatorSalt, passwordHash(LOCAL_CREATOR_PASSWORD, creatorSalt), "[]", "[]", now, now],
    [LOCAL_USER_EMAIL, "Maya Patel", "user", "1", "1", "0", "0", "local-user-feed", userSalt, passwordHash(LOCAL_USER_PASSWORD, userSalt), "[]", "[]", now, now],
    [LOCAL_UNCLAIMED_EMAIL, "Unclaimed Local User", "user", "0", "0", "1", "0", "local-unclaimed-feed", unclaimedSalt, passwordHash("local-unclaimed-only", unclaimedSalt), "[]", "[]", now, now],
  ];
  const canonicalRows = doctors.map((doctor) => [
    `${doctor.sourceType}:${doctor.key}`,
    doctor.displayName,
    doctor.sourceType,
    JSON.stringify([doctor.sourceType]),
    JSON.stringify([{ key: doctor.key, displayName: doctor.displayName, sourceType: doctor.sourceType }]),
    "1",
    now,
  ]);
  return `-- Synthetic, test-only local data. Never apply this file remotely.
BEGIN TRANSACTION;

INSERT INTO account_profiles (
  email, real_name, role, insights_enabled, facility_overview_enabled,
  non_clinical, director_view_enabled, subscription_token, password_salt,
  password_hash, admin_issues_json, local_parser_extensions_json, created_at, updated_at
) VALUES
  ${sqlValues(accountRows)}
ON CONFLICT(email) DO UPDATE SET
  real_name = excluded.real_name,
  role = excluded.role,
  insights_enabled = excluded.insights_enabled,
  facility_overview_enabled = excluded.facility_overview_enabled,
  non_clinical = excluded.non_clinical,
  director_view_enabled = excluded.director_view_enabled,
  subscription_token = excluded.subscription_token,
  password_salt = excluded.password_salt,
  password_hash = excluded.password_hash,
  admin_issues_json = excluded.admin_issues_json,
  local_parser_extensions_json = excluded.local_parser_extensions_json,
  updated_at = excluded.updated_at;

INSERT INTO account_states (email, session_json, updated_at) VALUES
  ${sqlValues(accountRows.map(([email]) => [email, JSON.stringify({ version: 1, session: { customEvents: [] } }), now]))}
ON CONFLICT(email) DO UPDATE SET session_json = excluded.session_json, updated_at = excluded.updated_at;

DELETE FROM account_claims WHERE email IN (${sqlString(LOCAL_USER_EMAIL)}, ${sqlString(LOCAL_UNCLAIMED_EMAIL)});
INSERT INTO account_claims (email, source_type, doctor_key, display_name, matched_at, updated_at)
VALUES (${sqlString(LOCAL_USER_EMAIL)}, 'mmc', 'MAYA PATEL', 'Maya PATEL', ${sqlString(now)}, ${sqlString(now)});

INSERT INTO account_hospital_locations (email, source_type, location, updated_at) VALUES
  (${sqlString(LOCAL_CREATOR_EMAIL)}, 'mmc', 'Local MMC', ${sqlString(now)}),
  (${sqlString(LOCAL_CREATOR_EMAIL)}, 'ddh', 'Local DDH', ${sqlString(now)}),
  (${sqlString(LOCAL_USER_EMAIL)}, 'mmc', 'Local MMC', ${sqlString(now)})
ON CONFLICT(email, source_type) DO UPDATE SET location = excluded.location, updated_at = excluded.updated_at;

INSERT INTO custom_events (
  owner_email, id, title, start_date, end_date, all_day, start_time, end_time,
  location, include, updated_at
) VALUES (
  ${sqlString(LOCAL_USER_EMAIL)}, 'local-custom-event', 'Local teaching session',
  '2026-09-09', '2026-09-09', 0, '17:00', '18:00', 'Local meeting room', 1, ${sqlString(now)}
)
ON CONFLICT(owner_email, id) DO UPDATE SET title = excluded.title, updated_at = excluded.updated_at;

INSERT INTO roster_files (
  id, name, source_type, source_id, active, size, last_modified, added_at,
  uploaded_at, uploaded_by, parsed_at, parser_version
) VALUES
  ${sqlValues(files.map((file, index) => [file.id, file.name, file.sourceType, `local-${file.sourceType}`, "1", String(1000 + index), "1788739200000", now, now, LOCAL_CREATOR_EMAIL, now, "local-seed-v1"]))}
ON CONFLICT(id) DO UPDATE SET
  name = excluded.name, source_type = excluded.source_type, source_id = excluded.source_id,
  active = excluded.active, uploaded_at = excluded.uploaded_at, parsed_at = excluded.parsed_at,
  parser_version = excluded.parser_version;

INSERT INTO roster_doctors (source_type, doctor_key, display_name, updated_at) VALUES
  ${sqlValues(doctors.map((doctor) => [doctor.sourceType, doctor.key, doctor.displayName, now]))}
ON CONFLICT(source_type, doctor_key) DO UPDATE SET display_name = excluded.display_name, updated_at = excluded.updated_at;

INSERT INTO roster_file_doctors (
  file_id, source_type, doctor_key, display_name, seniority, membership_source
) VALUES
  ${sqlValues(doctors.map((doctor) => [`local-${doctor.sourceType}`, doctor.sourceType, doctor.key, doctor.displayName, doctor.seniority, "local-seed"]))}
ON CONFLICT(file_id, source_type, doctor_key) DO UPDATE SET
  display_name = excluded.display_name, seniority = excluded.seniority, membership_source = excluded.membership_source;

INSERT INTO roster_events (
  id, file_id, source_type, doctor_key, display_name, start_date, end_date,
  start_ts, end_ts, title, raw_value, seniority, location, all_day, time_label, event_json
) VALUES
  ${sqlValues(events.map((event) => [event.id, event.fileId, event.sourceType, event.key, event.displayName, event.startDate, event.endDate, `${event.startDate}T08:00:00`, `${event.endDate}T16:00:00`, event.title, event.rawValue, event.seniority, event.location, "0", event.timeLabel, event.eventJson]))}
ON CONFLICT(id) DO UPDATE SET event_json = excluded.event_json, start_date = excluded.start_date,
  end_date = excluded.end_date, start_ts = excluded.start_ts, end_ts = excluded.end_ts;

INSERT INTO roster_daily_presence (date, source_type, doctor_key, display_name, event_id) VALUES
  ${sqlValues(events.map((event) => [event.startDate, event.sourceType, event.key, event.displayName, event.id]))}
ON CONFLICT(date, source_type, doctor_key, event_id) DO UPDATE SET display_name = excluded.display_name;

INSERT INTO canonical_doctors (
  canonical_key, display_name, source_type, source_types_json, aliases_json, has_events, updated_at
) VALUES
  ${sqlValues(canonicalRows)}
ON CONFLICT(canonical_key) DO UPDATE SET
  display_name = excluded.display_name, source_type = excluded.source_type,
  source_types_json = excluded.source_types_json, aliases_json = excluded.aliases_json,
  has_events = excluded.has_events, updated_at = excluded.updated_at;

COMMIT;
`;
}

export async function seedLocalDatabase() {
  await checkLocalSafety();
  await mkdir(LOCAL_STATE_DIRECTORY, { recursive: true });
  const seedPath = path.join(LOCAL_STATE_DIRECTORY, "seed.sql");
  await writeFile(seedPath, localSeedSql(), { encoding: "utf8", mode: 0o600 });
  try {
    console.log(`Seeding synthetic local data only: ${LOCAL_STATE_DIRECTORY}`);
    runWrangler([
      "d1", "execute", "ROSTER_DB", "--local",
      "--persist-to", LOCAL_STATE_DIRECTORY,
      "--file", seedPath,
    ], { capture: true });
    console.log("Synthetic local seed applied successfully.");
  } finally {
    await unlink(seedPath).catch(() => {});
  }
}

function localDevArguments(port) {
  return [
    "pages", "dev", "public",
    "--ip", LOCAL_HOST,
    "--port", String(port),
    "--persist-to", LOCAL_STATE_DIRECTORY,
    "--binding", "LOCAL_ONLY=true",
    "--binding", "ROSTER_AUTOMATION_ENABLED=false",
    "--binding", "ROSTER_AUTOMATION_WRITES_ENABLED=false",
    "--binding", "EMAIL_DELIVERY_ENABLED=false",
  ];
}

export async function startLocalServer(options = {}) {
  const { port } = await checkLocalSafety({ port: options.port, requirePortAvailable: true });
  await mkdir(LOCAL_STATE_DIRECTORY, { recursive: true });
  const child = spawn(WRANGLER_PATH, localDevArguments(port), {
    cwd: REPOSITORY_ROOT,
    env: safeChildEnvironment(),
    stdio: options.capture === true ? ["ignore", "pipe", "pipe"] : "inherit",
  });
  let output = "";
  if (options.capture === true) {
    child.stdout?.on("data", (chunk) => { output = `${output}${chunk}`.slice(-30000); });
    child.stderr?.on("data", (chunk) => { output = `${output}${chunk}`.slice(-30000); });
  }
  child.localPort = port;
  child.localOutput = () => output;
  return child;
}

export async function waitForLocalServer(child, timeoutMs = 30000) {
  const startedAt = Date.now();
  const url = `http://${LOCAL_HOST}:${child.localPort}/`;
  while (Date.now() - startedAt < timeoutMs) {
    if (child.exitCode !== null) throw new Error(`Local Pages exited before becoming ready.\n${child.localOutput?.() || ""}`);
    try {
      const response = await fetch(url);
      if (response.ok) return url;
    } catch {}
    await new Promise((resolve) => setTimeout(resolve, 200));
  }
  throw new Error(`Local Pages did not become ready at ${url}.\n${child.localOutput?.() || ""}`);
}

export async function localLogin(port, credentials = {}) {
  const response = await fetch(`http://${LOCAL_HOST}:${port}/api/state`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      action: "login",
      mode: "login",
      responseMode: "fast",
      email: credentials.email || LOCAL_CREATOR_EMAIL,
      password: credentials.password || LOCAL_CREATOR_PASSWORD,
    }),
  });
  const payload = await response.json().catch(() => ({}));
  return { response, payload };
}

export async function stopLocalServer(child) {
  if (!child || child.exitCode !== null) return;
  child.kill("SIGINT");
  await Promise.race([
    new Promise((resolve) => child.once("exit", resolve)),
    new Promise((resolve) => setTimeout(resolve, 5000)),
  ]);
  if (child.exitCode === null) child.kill("SIGTERM");
}

export async function verifyLocalLogin(port) {
  const { response, payload } = await localLogin(port);
  if (!response.ok || payload?.ok !== true || payload?.role !== "creator") {
    throw new Error(`Seeded local Creator login failed (${response.status}): ${payload?.error || "unknown error"}`);
  }
  return payload;
}

export async function setupLocalEnvironment(options = {}) {
  const port = localPort(options.port);
  await resetLocalState({ port });
  await migrateLocalDatabase();
  await seedLocalDatabase();
  const child = await startLocalServer({ port, capture: true });
  try {
    await waitForLocalServer(child);
    await verifyLocalLogin(port);
  } finally {
    await stopLocalServer(child);
  }
  console.log("Local setup verified with a real Creator login.");
  printLocalDetails(port);
}

function printLocalDetails(port) {
  console.log(`Local URL: http://${LOCAL_HOST}:${port}/`);
  console.log(`Local Creator: ${LOCAL_CREATOR_EMAIL}`);
  console.log(`Local test password: ${LOCAL_CREATOR_PASSWORD}`);
  console.log(`Local data: ${LOCAL_STATE_DIRECTORY}`);
  console.log("Cloudflare D1/R2: not in use");
}

async function runLongLivedLocalDev() {
  const port = localPort();
  const child = await startLocalServer({ port });
  await writeFile(PID_FILE, String(process.pid), "utf8");
  printLocalDetails(port);
  const forwardSignal = (signal) => {
    if (child.exitCode === null) child.kill(signal);
  };
  process.once("SIGINT", () => forwardSignal("SIGINT"));
  process.once("SIGTERM", () => forwardSignal("SIGTERM"));
  const code = await new Promise((resolve) => child.once("exit", (exitCode) => resolve(exitCode ?? 0)));
  await unlink(PID_FILE).catch(() => {});
  process.exitCode = code;
}

async function main() {
  const command = String(process.argv[2] || "").trim();
  if (process.argv.length > 3) throw new Error("Local commands do not accept additional arguments. Use LOCAL_DEV_PORT only when a different local port is required.");
  if (command === "check") {
    const result = await checkLocalSafety({ requirePortAvailable: true });
    console.log(`Local safety check passed for http://${result.host}:${result.port}/`);
    console.log(`Local data: ${result.stateDirectory}`);
    console.log("Cloudflare D1/R2: not in use");
    return;
  }
  if (command === "reset") return await resetLocalState();
  if (command === "migrate") return await migrateLocalDatabase();
  if (command === "seed") return await seedLocalDatabase();
  if (command === "setup") return await setupLocalEnvironment();
  if (command === "dev") return await runLongLivedLocalDev();
  throw new Error("Use one of: check, reset, migrate, seed, setup, dev.");
}

if (path.resolve(process.argv[1] || "") === SCRIPT_PATH) {
  main().catch((error) => {
    console.error(`Local safety error: ${error?.message || error}`);
    process.exitCode = 1;
  });
}
