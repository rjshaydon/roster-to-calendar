import { readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { buildRosterView, doctorOptions, parseUploadForm, setParserExtensions } from "../public/static/roster.js";

const backupRoot = "backups/production-2026-08-12";
const active = [
  ["Casey Term 2 2026.xlsm:474906:1778218074702", "casey"],
  ["automation:dandenong-findmyshift:b0d33978d30d5ee10f48a86e", "ddh"],
  ["Dandenong_Emergency_Doctors'_Roster_02-02-2026_to_03-05-2026.xlsx:146512:1777464564005", "ddh"],
  ["Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007", "ddh"],
  ["Paeds - Term 1 2026.xlsx:500042:1778296700465", "mch"],
  ["automation:monash-paeds:ed3f4a068861fe848d5a1acb", "mch"],
  ["automation:monash-paeds:98632c7f2e7f1311f7a89a8e", "mch"],
  ["AdultTerm1.2026.xlsx:641068:1776812908257", "mmc"],
  ["automation:monash-adults:ac7f9d2e29c6bbb35e8a86df", "mmc"],
  ["automation:monash-adults:aec4fa0948b79e34f311b6c3", "mmc"],
];
const sql = await readFile(`${backupRoot}/roster-converter-calendar.sql`, "utf8");
const sqlValues = (line) => [...line.matchAll(/'((?:[^']|'')*)'/g)].map((match) => match[1].replaceAll("''", "'"));
const productionRules = {};
for (const line of sql.split("\n")) {
  if (!line.startsWith('INSERT INTO "parser_rules"')) continue;
  const values = sqlValues(line);
  if (values[1] !== "global") continue;
  try {
    const rule = JSON.parse(values.at(-2) || "{}");
    const source = String(rule.source || "").toLowerCase();
    if (source) productionRules[source] = [...(productionRules[source] || []), rule];
  } catch {
    throw new Error("Could not load a production parser rule from the backup.");
  }
}
if (!Object.keys(productionRules).length) throw new Error("No production parser rules were found in the backup.");
setParserExtensions(productionRules);
const rawRows = new Map();
for (const line of sql.split("\n")) {
  if (!line.startsWith('INSERT INTO "raw_roster_files"')) continue;
  const match = line.match(/VALUES\('((?:[^']|'')*)','((?:[^']|'')*)','[^']*','((?:[^']|'')*)','(?:[^']|'')*','((?:[^']|'')*)','((?:[^']|'')*)'/);
  if (!match) continue;
  const unescape = (value) => value.replaceAll("''", "'");
  rawRows.set(unescape(match[1]), { objectKey: unescape(match[2]), type: unescape(match[3]), name: unescape(match[4]), sourceType: unescape(match[5]) });
}
const old = new Map(active.map(([id, sourceType]) => [id, { sourceType, events: [], issues: [] }]));
for (const line of sql.split("\n")) {
  if (line.startsWith('INSERT INTO "roster_events"')) {
    const values = sqlValues(line);
    const id = values[1];
    if (old.has(id)) old.get(id).events.push({
      doctorKey: values[3],
      event: JSON.parse(values.at(-1) || "{}"),
    });
  }
  if (line.startsWith('INSERT INTO "roster_issues"')) {
    const match = line.match(/VALUES\('(?:[^']|'')*','((?:[^']|'')*)'/);
    const id = match?.[1]?.replaceAll("''", "'");
    if (old.has(id) && /'unknown'/.test(line)) old.get(id).issues.push(line);
  }
}
const results = [];
for (const [id, sourceType] of active) {
  const raw = rawRows.get(id);
  if (!raw) throw new Error(`Missing raw row for ${id}`);
  const hash = createHash("sha256").update(raw.objectKey).digest("hex");
  const bytes = await readFile(`${backupRoot}/roster-source-files/${hash}`);
  const form = new FormData();
  form.append("rosterFiles", new File([bytes], raw.name, { type: raw.type }));
  const parsed = await parseUploadForm(new Request("http://shadow.test/api/analyze", { method: "POST", body: form }));
  const sources = parsed.sources;
  const doctors = doctorOptions(sources.mmc, sources.ddh, sources.casey, sources.mch)
    .filter((doctor) => (doctor.sourceTypes || [doctor.sourceType]).includes(sourceType));
  const events = [];
  const issues = [];
  for (const doctor of doctors) {
    const view = buildRosterView(sources.mmc, sources.ddh, doctor.key, undefined, {}, {}, doctor.aliases || [], sources.casey, sources.mch);
    events.push(...view.events
      .filter((event) => String(event.source || "").toLowerCase() === sourceType)
      .map((event) => ({ ...event, doctorKey: doctor.key })));
    issues.push(...(view.issues || []).filter((issue) => String(issue.source || "").toLowerCase() === sourceType && String(issue.status || "").toLowerCase() === "unknown"));
  }
  // Parser event ids describe the shift occurrence but intentionally omit the
  // doctor. D1 makes them unique as file_id + doctor_key + event.id, so counting
  // by event.id alone collapses different doctors working the same shift.
  const eventIdentity = (doctorKey, event) => `${doctorKey}|${event.id}`;
  const beforeEvents = new Map(old.get(id).events.map(({ doctorKey, event }) => [eventIdentity(doctorKey, event), { doctorKey, ...event }]));
  const uniqueEvents = new Map(events.map((event) => [eventIdentity(event.doctorKey, event), event]));
  const removed = [...beforeEvents].filter(([identity]) => !uniqueEvents.has(identity)).map(([, event]) => event);
  const added = [...uniqueEvents].filter(([identity]) => !beforeEvents.has(identity)).map(([, event]) => event);
  const unknownCodes = new Set(issues.map((issue) => `${issue.source}|${issue.seniority}|${String(issue.rawValue || "").toUpperCase()}`));
  results.push({
    id,
    name: raw.name,
    sourceType,
    before: { events: old.get(id).events.length, unknownOccurrences: old.get(id).issues.length },
    after: { events: events.length, distinctStoredEventIdentities: uniqueEvents.size, unknownOccurrences: issues.length, distinctUnknownCodes: unknownCodes.size },
    delta: {
      addedCount: added.length,
      removedCount: removed.length,
      added: added.slice(0, 20),
      removed: removed.slice(0, 20),
    },
  });
}
await writeFile(`${backupRoot}/shadow-active-rosters.json`, `${JSON.stringify(results, null, 2)}\n`);
console.log(JSON.stringify(results, null, 2));
