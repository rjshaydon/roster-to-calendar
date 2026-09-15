import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import { basename, resolve } from "node:path";
import { buildRosterView, doctorOptions, parseUploadForm, setParserExtensions } from "../public/static/roster.js";

const filePath = resolve(process.argv[2] || "");
const baselineId = String(process.argv[3] || "").trim();
const backupSqlPath = process.argv[4] ? resolve(process.argv[4]) : "";
if (!process.argv[2]) throw new Error("Usage: node scripts/inspect-roster-workbook.mjs <workbook> [baseline-file-id baseline-sql]");
if (baselineId && !backupSqlPath) throw new Error("A baseline SQL path is required with a baseline file id.");

const backupSql = backupSqlPath ? await readFile(backupSqlPath, "utf8") : "";
const sqlValues = (line) => [...line.matchAll(/'((?:[^']|'')*)'/g)].map((match) => match[1].replaceAll("''", "'"));
const productionRules = {};
for (const line of backupSql.split("\n")) {
  if (!line.startsWith('INSERT INTO "parser_rules"')) continue;
  const values = sqlValues(line);
  if (values[1] !== "global") continue;
  const rule = JSON.parse(values.at(-2) || "{}");
  const source = String(rule.source || "").toLowerCase();
  if (source) productionRules[source] = [...(productionRules[source] || []), rule];
}
if (Object.keys(productionRules).length) setParserExtensions(productionRules);

const bytes = await readFile(filePath);
const form = new FormData();
form.append("rosterFiles", new File([bytes], basename(filePath), {
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
}));
const { sources } = await parseUploadForm(new Request("http://local.test/api/analyze", { method: "POST", body: form }));

const summary = {
  fileName: basename(filePath),
  bytes: bytes.length,
  sha256: createHash("sha256").update(bytes).digest("hex"),
  sources: {},
};
for (const sourceType of ["mmc", "mch", "ddh", "casey"]) {
  const doctors = doctorOptions(sources.mmc, sources.ddh, sources.casey, sources.mch)
    .filter((doctor) => (doctor.sourceTypes || [doctor.sourceType]).includes(sourceType));
  const events = [];
  const issues = [];
  for (const doctor of doctors) {
    const view = buildRosterView(
      sources.mmc,
      sources.ddh,
      doctor.key,
      undefined,
      {},
      {},
      doctor.aliases || [],
      sources.casey,
      sources.mch,
    );
    events.push(...view.events
      .filter((event) => String(event.source || "").toLowerCase() === sourceType)
      .map((event) => ({ doctorKey: doctor.key, ...event })));
    issues.push(...(view.issues || []).filter((issue) => String(issue.source || "").toLowerCase() === sourceType));
  }
  if (doctors.length || events.length || issues.length) {
    const sourceSummary = { doctors: doctors.length, events: events.length, issues: issues.length };
    if (baselineId) {
      const oldEvents = [];
      for (const line of backupSql.split("\n")) {
        if (!line.startsWith('INSERT INTO "roster_events"')) continue;
        const values = sqlValues(line);
        if (values[1] === baselineId) oldEvents.push({ doctorKey: values[3], ...JSON.parse(values.at(-1) || "{}") });
      }
      const identity = (event) => `${event.doctorKey}|${event.id}`;
      const before = new Set(oldEvents.map(identity));
      const after = new Set(events.map(identity));
      sourceSummary.baseline = {
        fileId: baselineId,
        events: oldEvents.length,
        added: [...after].filter((id) => !before.has(id)).length,
        removed: [...before].filter((id) => !after.has(id)).length,
      };
    }
    summary.sources[sourceType] = sourceSummary;
  }
}

process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`);
