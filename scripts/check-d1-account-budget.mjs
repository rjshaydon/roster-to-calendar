import { chmod, readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { execFileSync } from "node:child_process";

import { d1AnalyticsQuery, evaluateD1Budget, settledUtcDayInterval, summarizeAnalyticsPayload } from "./d1-quota-budget-lib.mjs";

const options = parseArguments(process.argv.slice(2));
const generatedAt = new Date().toISOString();
const inventoryPath = resolve(options.inventory || "config/d1-database-inventory.json");
const inventory = JSON.parse(await readFile(inventoryPath, "utf8"));
const previousReport = options.previous ? JSON.parse(await readFile(resolve(options.previous), "utf8")) : null;
const accountId = String(process.env.CLOUDFLARE_ACCOUNT_ID || inventory.accountId || "").trim();
const analyticsCredential = loadAnalyticsCredential();
const token = analyticsCredential.token;

let analytics = { complete: false, reasons: [], databases: [], totals: {}, queryFingerprints: [] };
if (!accountId) analytics.reasons.push("cloudflare-account-id-required");
if (!token) analytics.reasons.push("account-analytics-read-token-required");
if (accountId && token) {
  try {
    const interval = settledUtcDayInterval(generatedAt);
    if (!interval.settled) {
      analytics = { ...analytics, reasons: ["analytics-settlement-window-before-utc-day"], interval };
    } else {
      const response = await fetch("https://api.cloudflare.com/client/v4/graphql", {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          query: d1AnalyticsQuery(),
          variables: { accountTag: accountId, start: interval.start, end: interval.observedUntil },
        }),
      });
      if (!response.ok) throw new Error(`analytics-http-${response.status}`);
      const payload = await response.json();
      if (options.rawOutput) {
        const rawPath = resolve(options.rawOutput);
        await writeFile(rawPath, `${JSON.stringify(payload, null, 2)}\n`, { encoding: "utf8", mode: 0o600 });
        await chmod(rawPath, 0o600);
      }
      analytics = { ...summarizeAnalyticsPayload(payload, inventory), interval };
    }
  } catch (error) {
    analytics = { complete: false, reasons: [`analytics-request-failed:${String(error?.message || "unknown")}`], databases: [], totals: {}, queryFingerprints: [] };
  }
}

const assessment = evaluateD1Budget({
  analytics,
  inventory,
  billing: { rowsRead: options.billingReads, rowsWritten: options.billingWrites, observedAt: options.billingObservedAt, unavailableReason: options.billingUnavailableReason },
  previousReport,
  now: generatedAt,
  optionalEstimate: { rowsRead: options.estimatedReads, rowsWritten: options.estimatedWrites },
});
const report = {
  schemaVersion: 1,
  generatedAt,
  inventory: { path: inventoryPath, complete: inventory.complete === true, databaseCount: (inventory.databases || []).length },
  credential: { source: analyticsCredential.source },
  analytics,
  billing: { rowsRead: numericOrNull(options.billingReads), rowsWritten: numericOrNull(options.billingWrites), observedAt: options.billingObservedAt || null, unavailableReason: options.billingUnavailableReason || null },
  ...assessment,
};
const serialized = `${JSON.stringify(report, null, 2)}\n`;
if (options.output) {
  const outputPath = resolve(options.output);
  await writeFile(outputPath, serialized, { encoding: "utf8", mode: 0o600 });
  await chmod(outputPath, 0o600);
}
process.stdout.write(serialized);
if (report.decision !== "GO") process.exitCode = 2;

function parseArguments(args) {
  const parsed = {};
  for (let index = 0; index < args.length; index += 1) {
    const key = args[index];
    if (!key.startsWith("--") || index + 1 >= args.length) throw new Error(`Invalid argument: ${key}`);
    const value = args[++index];
    const name = key.slice(2).replace(/-([a-z])/g, (_match, letter) => letter.toUpperCase());
    if (!["inventory", "previous", "output", "rawOutput", "billingReads", "billingWrites", "billingObservedAt", "billingUnavailableReason", "estimatedReads", "estimatedWrites"].includes(name)) throw new Error(`Unknown argument: ${key}`);
    parsed[name] = value;
  }
  return parsed;
}

function numericOrNull(value) {
  const number = Number(value);
  return Number.isFinite(number) && number >= 0 ? number : null;
}

function loadAnalyticsCredential() {
  const environmentToken = String(process.env.CLOUDFLARE_ACCOUNT_ANALYTICS_TOKEN || "").trim();
  if (environmentToken) return { token: environmentToken, source: "environment" };
  if (process.platform !== "darwin") return { token: "", source: null };

  try {
    const keychainToken = execFileSync(
      "/usr/bin/security",
      ["find-generic-password", "-s", "roster-d1-account-analytics", "-w"],
      { encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] },
    ).trim();
    return keychainToken ? { token: keychainToken, source: "macos-keychain" } : { token: "", source: null };
  } catch {
    return { token: "", source: null };
  }
}
