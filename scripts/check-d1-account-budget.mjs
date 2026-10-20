import {maintenanceDiagnosticEvidence,evaluateMaintenanceDiagnostic} from './maintenance-budget-diagnostic.mjs';
import { chmod, readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { execFileSync } from "node:child_process";

import { d1AnalyticsQuery, evaluateD1Budget, settledUtcDayInterval, summarizeAnalyticsPayload } from "./d1-quota-budget-lib.mjs";

const options = parseArguments(process.argv.slice(2));
const generatedAt = new Date().toISOString();
const inventoryPath = resolve(options.inventory || "config/d1-database-inventory.json");
const inventory = JSON.parse(await readFile(inventoryPath, "utf8"));
const previousReport = options.previous ? JSON.parse(await readFile(resolve(options.previous), "utf8")) : null;
const requestAttribution = options.requestAttribution ? JSON.parse(await readFile(resolve(options.requestAttribution), "utf8")) : null;
const reviewedFingerprintDocument = options.reviewedFingerprints ? JSON.parse(await readFile(resolve(options.reviewedFingerprints), "utf8")) : null;
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

let assessment;
try {
  if(!analytics.complete)throw Error('Account analytics are unavailable or inconsistent.');
  const config=await readFile(resolve('wrangler.toml'),'utf8');
  const productionConfig=config.split('[env.')[0];
  const previewEnabled=/^ACCOUNT_MAINTENANCE_PREVIEW_SUMMARY_ENABLED\s*=\s*"true"/m.test(productionConfig);
  const primary=inventory.databases.find(d=>d.environment==='production'&&d.name==='roster-converter-calendar');
  const preview=inventory.databases.find(d=>d.name==='roster-identity-batch2-preview');
  if(!primary||previewEnabled&&!preview)throw Error('Maintenance database inventory is incomplete.');
  const evidence=await maintenanceDiagnosticEvidence(readOnlyD1(primary.name),preview?readOnlyD1(preview.name):null,analytics,{previewEnabled});
  assessment=evaluateMaintenanceDiagnostic(analytics,evidence,new Date(generatedAt),{rowsRead:Number(options.estimatedReads||0),rowsWritten:Number(options.estimatedWrites||0)});
} catch(error){assessment={decision:'STOP',advisory:true,reasons:[String(error.message||error)]};}

if(options.mode==='canary') {
 const canaryAssessment=evaluateD1Budget({analytics,inventory,billing:{rowsRead:options.billingReads,rowsWritten:options.billingWrites,observedAt:options.billingObservedAt,unavailableReason:options.billingUnavailableReason},previousReport,now:generatedAt,
 canary:{readCeiling:options.canaryReadCeiling,requestId:options.requestId,requestAttribution,reviewedFingerprints:Array.isArray(reviewedFingerprintDocument)?reviewedFingerprintDocument:reviewedFingerprintDocument?.queries||[]}});
 assessment={...assessment,canary:canaryAssessment,decision:assessment.decision==='GO'&&canaryAssessment.decision==='GO'?'GO':'STOP',reasons:[...new Set([...assessment.reasons,...canaryAssessment.reasons])]};
}
const report = {
  schemaVersion: 2,
  sampleValid:analytics.complete===true && assessment.budget!==null && assessment.budget!==undefined,
  interval:analytics.interval,
  effectiveUsage:{rowsRead:analytics.totals?.rowsRead||0,rowsWritten:analytics.totals?.rowsWritten||0},
  generatedAt,
  inventory: { path: inventoryPath, complete: inventory.complete === true, databaseCount: (inventory.databases || []).length },
  credential: { source: analyticsCredential.source },
  analytics: {complete:analytics.complete,reasons:analytics.reasons,totals:analytics.totals,maximumFiveMinute:analytics.maximumFiveMinute,interval:analytics.interval},
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
    if (!["inventory", "previous", "output", "rawOutput", "billingReads", "billingWrites", "billingObservedAt", "billingUnavailableReason", "estimatedReads", "estimatedWrites", "mode", "canaryReadCeiling", "requestId", "requestAttribution", "reviewedFingerprints"].includes(name)) throw new Error(`Unknown argument: ${key}`);
    parsed[name] = value;
  }
  if (parsed.mode && !["baseline", "canary"].includes(parsed.mode)) throw new Error(`Invalid mode: ${parsed.mode}`);
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

// Known aggregate queries only; this adapter cannot mutate production or return
// clinical roster rows. Arguments are passed directly to Wrangler, never a shell.
function readOnlyD1(databaseId) {
 return {prepare(sql){
  if(!/^\s*(?:WITH|SELECT)\b/i.test(sql)||/\b(?:INSERT|UPDATE|DELETE|DROP|ALTER|PRAGMA)\b/i.test(sql))throw Error('Diagnostic SQL must be read-only.');
  return {args:[],bind(...args){this.args=args;return this;},async first(){
   let index=0;const bound=sql.replace(/\?/g,()=>{const value=this.args[index++];if(value===undefined)throw Error('Missing diagnostic binding.');return typeof value==='number'?String(value):"'"+String(value).replaceAll("'","''")+"'";});
   if(index!==this.args.length)throw Error('Unexpected diagnostic binding.');
   let result;
   for(let attempt=0;attempt<2;attempt++)try {
    result=JSON.parse(execFileSync('npx',['wrangler','d1','execute',databaseId,'--remote','--command',bound,'--json'],{encoding:'utf8',maxBuffer:1024*1024,stdio:['ignore','pipe','pipe']}));break;
   }catch(error){if(attempt||!String(error.stdout||error.stderr||'').includes('7403'))throw Error('Read-only maintenance summary unavailable: '+String(error.stdout||error.stderr||error.message||'').replace(/\x1b\[[0-9;]*m/g,'').slice(0,1000));}
   const row=result?.[0];if(!row?.success||!Array.isArray(row.results)||row.results.length>1)throw Error('Invalid maintenance summary response.');return row.results[0]||null;
  }};
 }};
}
