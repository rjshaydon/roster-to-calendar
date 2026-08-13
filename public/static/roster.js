let XLSX = null;
let decompressSync = null;
let spreadsheetDependencyPromise = null;
let pdfDependencyPromise = null;

async function ensureSpreadsheetDependency() {
  if (XLSX) return XLSX;
  if (!spreadsheetDependencyPromise) {
    spreadsheetDependencyPromise = import("xlsx")
      .then((module) => {
        XLSX = module;
        return XLSX;
      })
      .catch((error) => {
        spreadsheetDependencyPromise = null;
        throw error;
      });
  }
  return spreadsheetDependencyPromise;
}

async function ensurePdfDependencies() {
  await ensureSpreadsheetDependency();
  if (decompressSync) return decompressSync;
  if (!pdfDependencyPromise) {
    pdfDependencyPromise = import("fflate")
      .then((module) => {
        decompressSync = module.decompressSync;
        return decompressSync;
      })
      .catch((error) => {
        pdfDependencyPromise = null;
        throw error;
      });
  }
  return pdfDependencyPromise;
}

const TIMEZONE = "Australia/Melbourne";

const MMC_LOCATION = "MMC Car Park, Tarella Road, Clayton VIC 3168, Australia";
const DDH_LOCATION = "DDH Car Park, 135 David St, Dandenong VIC 3175, Australia";
const CASEY_LOCATION = "Casey Hospital, 62-70 Kangan Drive, Berwick VIC 3806, Australia";
const MCH_LOCATION = "Monash Children's Hospital, 246 Clayton Road, Clayton VIC 3168, Australia";
const UNKNOWN_SENIORITY = "Unknown";
const SENIORITY_LABELS = [
  "SMS",
  "CMO",
  "Senior Registrar",
  "Transitional/Intermediate Registrar",
  "Junior Registrar",
  "HMO",
  "ENP",
  "AMP",
  "Intern",
  UNKNOWN_SENIORITY,
];

const MMC_TEAM_MAP = {
  G: "Green",
  A: "Amber",
  R: "Resus",
  C: "Clinic",
};

const DDH_LABEL_MAP = {
  "Clinical Support": "CS",
  "SSU SMS": "SSU",
  "Orange PM (on-call)": "Orange PM",
  "AVAO PM": "AVAO PM",
  "AVAO AM": "AVAO AM",
  "PM FAST IC": "FAST PM",
  "Orange AM IC": "Orange AM",
  "Silver AM IC": "Silver AM",
  "Silver PM IC": "Silver PM",
  "onsite CS": "CS onsite",
  "PHNW clinical": "PHNW",
  PHNW: "PHNW",
};

const KNOWN_DDH_DIRECT_LABELS = new Set([
  "CS",
  "CS onsite",
  "SSU",
  "Orange AM",
  "Orange PM",
  "Silver AM",
  "Silver PM",
  "FAST PM",
  "AVAO AM",
  "AVAO PM",
  "PHNW",
]);

const WEEKDAY_PREFIXES = ["Mon.", "Tue.", "Wed.", "Thu.", "Fri.", "Sat.", "Sun."];
const WEEKLY_LEAVE_LABELS = new Set(["ANNUAL LEAVE", "CONFERENCE LEAVE", "CME LEAVE", "CME/L"]);
const IGNORED_EXACT = new Set([
  "",
  "OFF",
  "AL",
  "A/L",
  "EXAM",
  "EXAM LEAVE",
  "PARENTAL LEAVE",
  "N/A",
]);
const IGNORED_CONTAINS = [
  "TEACHING",
  "EXAM",
  "MEETING",
  "MOCK",
  "HOLIDAY",
  "LABOUR DAY",
  "GOOD FRIDAY",
  "EASTER",
  "ANZAC",
  "BENDIGO",
  "ACEM",
  "REZA",
  "JENNY",
  "ARRHCHIE",
];
const MMC_SECTION_MARKERS = new Set([
  "GERIATRICIAN",
  "CMO",
  "SENIOR REG",
  "INTERMEDIATE REG",
  "JUNIOR REG",
  "HMO",
  "HMO MUST BE",
  "HMO MUST BE 111",
  "HMO - MUST BE 111",
  "ENP",
  "AMP",
  "EMERGENCY NURSE PRACTITIONER",
  "AMBULATORY MUSCULOSKELETAL PHYSIOTHERAPIST",
  "INTERN",
  "LOCUM",
]);
const MMC_STOP_SECTIONS = new Set([
  "INTERN",
  "LOCUM",
]);
const DDH_SECTION_MARKERS = new Set([
  "SENIOR MEDICAL STAFF",
  "JUNIOR MEDICAL STAFF",
  "REGISTRAR",
  "REGISTRARS",
  "HMO",
  "INTERN",
  "INTERNS",
  "PHYSIOTHERAPIST",
  "PHYSIOTHERAPISTS",
  "ENP",
  "AMP",
]);
const DDH_IGNORE_PREFIXES = [
  "YES",
  "NO",
  "OK",
  "N/A",
  "NA",
  "OFF",
  "WORKED",
  "FOR ",
  "IN LIEU",
  "PREFER",
  "RELUCTANT",
  "NOT ",
];
const DDH_IGNORE_CONTAINS = [
  "ORIENTATION",
  "TEACHING",
  "MEET AND GREET",
  "SKILLS",
  "SIM",
  "EXAM",
  "PLEASE",
  "NOT AVAILABLE",
  "NO CLINICAL",
  "NO AM",
  "NO PM",
  "AM ONLY",
  "PM ONLY",
  "AM/PM",
];
const CASEY_SECTION_MARKERS = new Set([
  "GERIATRICIAN",
  "SNR CMOS",
  "SENIOR CMOS",
  "ED GPS",
  "SNR REGS",
  "SENIOR REGS",
  "JNR REGS",
  "JUNIOR REGS",
  "PAEDS ED HMOS",
  "HMOS",
  "HMO",
  "INTERNS",
  "LOCUMS / EXTRA STAFF",
  "AMP COVER",
  "ROSTERED STAFF",
]);
const CASEY_STOP_SECTIONS = new Set([
  "ROSTERED STAFF",
]);
const CASEY_IGNORED_EXACT = new Set([
  "",
  "0",
  "1",
  "2",
  "OFF",
  "OFF CS",
  "MMC",
  "DDH",
  "TEACH",
  "TEACHING",
  "SIM",
  "EXAM",
  "EXAM LEAVE",
  "SABBATICAL",
  "SABBATICAL LEAVE",
  "LONG SERVICE LEAVE",
  "LSL",
  "PATERNITY LEAVE",
  "VHH",
  "TOX",
]);
const DEFAULT_SETTINGS = {
  showSourcePrefix: true,
  showAmPm: true,
  showTimes: true,
  showRawValues: false,
  showNormalizedTitles: true,
  includeLocations: true,
  includeAnnualLeave: true,
  includeConferenceLeave: true,
  includePublicHoliday: true,
  includeSickLeave: true,
  defaultLocationMmc: "MMC Car Park, Tarella Road, Clayton VIC 3168, Australia",
  defaultLocationDdh: "DDH Car Park, 135 David St, Dandenong VIC 3175, Australia",
  defaultLocationCasey: "Casey Hospital, 62-70 Kangan Drive, Berwick VIC 3806, Australia",
  defaultLocationMch: "Monash Children's Hospital, 246 Clayton Road, Clayton VIC 3168, Australia",
  hospitalFilter: "all",
  dateFrom: "",
  dateTo: "",
};

let MANUAL_PARSER_RULES = buildDefaultParserRules();

export function defaultSettings() {
  return { ...DEFAULT_SETTINGS };
}

export function setParserExtensions(value) {
  MANUAL_PARSER_RULES = sanitizeParserExtensions(value);
}

export function parserRuleDefaults() {
  return sanitizeParserExtensions(buildDefaultParserRules());
}

export function parserRuleSeniorities() {
  return [...SENIORITY_LABELS];
}

export async function parseUploadForm(request) {
  const formData = await request.formData();
  const uploads = formData.getAll("rosterFiles").filter((item) => item instanceof File);
  if (!uploads.length) {
    throw new Error("Upload at least one roster file.");
  }

  const importIds = formData.getAll("rosterFileId");
  const importAddedAt = formData.getAll("rosterFileAddedAt");
  const sources = { mmc: [], ddh: [], casey: [], mch: [] };
  for (let index = 0; index < uploads.length; index += 1) {
    const file = uploads[index];
    const workbook = await readWorkbook(file);
    const sourceType = detectSourceType(workbook, file.name);
    validateRosterDatesWithinSingleTerm(workbook, sourceType, file.name);
    sources[sourceType].push({
      id: String(importIds[index] || hashString(`${file.name}|${file.size}|${file.lastModified}|${index}`)),
      addedAt: String(importAddedAt[index] || ""),
      file,
      workbook,
    });
  }

  if (!hasAnyRosterSource(sources)) {
    throw new Error("Upload at least one MMC, DDH, Casey, or MCH roster.");
  }

  return {
    sources,
    doctorKey: String(formData.get("doctorKey") || ""),
    doctorDisplay: String(formData.get("doctorDisplay") || ""),
    doctorAliases: sanitizeDoctorAliases(parseJsonField(formData, "doctorAliases", [])),
    settings: sanitizeSettings(parseJsonField(formData, "settings", DEFAULT_SETTINGS)),
    overrides: sanitizeOverrides(parseJsonField(formData, "overrides", {})),
    customEvents: sanitizeCustomEvents(parseJsonField(formData, "customEvents", [])),
    conflictSelections: parseJsonField(formData, "conflictSelections", {}),
  };
}

// Server-side retained-source rebuilding uses the same parser. This remains a
// thin transport adapter: roster interpretation stays in this module.
export async function buildRosterViewFromStoredImports(imports, doctorKey, settings = DEFAULT_SETTINGS, overrides = {}, conflictSelections = {}, doctorAliases = []) {
  const sources = { mmc: [], ddh: [], casey: [], mch: [] };
  for (const item of Array.isArray(imports) ? imports : []) {
    if (!item?.dataUrl) continue;
    const workbook = await readWorkbookDataUrl(item.dataUrl, item.name || "roster.xlsx");
    const sourceType = String(item.sourceType || detectSourceType(workbook, item.name || "roster.xlsx")).toLowerCase();
    if (!isRosterSourceType(sourceType)) continue;
    sources[sourceType].push({
      id: String(item.id || item.repoId || hashString(`${item.name || "import"}|${item.lastModified || 0}`)),
      addedAt: String(item.addedAt || ""),
      file: { name: String(item.name || "roster.xlsx"), size: Number(item.size || 0), lastModified: Number(item.lastModified || 0) },
      workbook,
    });
  }
  return buildRosterView(sources.mmc, sources.ddh, doctorKey, settings, overrides, conflictSelections, doctorAliases, sources.casey, sources.mch);
}

export function doctorOptions(mmcSources, ddhSources, caseySources = [], mchSources = []) {
  const mmcEntries = normalizeSourceEntries(mmcSources);
  const ddhEntries = normalizeSourceEntries(ddhSources);
  const caseyEntries = normalizeSourceEntries(caseySources);
  const mchEntries = normalizeSourceEntries(mchSources);
  if (!mmcEntries.length && !ddhEntries.length && !caseyEntries.length && !mchEntries.length) {
    throw new Error("Upload at least one MMC, DDH, Casey, or MCH roster.");
  }
  const mmcNames = new Map();
  const ddhNames = new Map();
  const caseyNames = new Map();
  const mchNames = new Map();
  for (const entry of mmcEntries) {
    for (const [key, value] of extractMmcNames(entry.workbook)) {
      if (!mmcNames.has(key)) mmcNames.set(key, value);
    }
  }
  for (const entry of ddhEntries) {
    for (const [key, value] of extractDdhNames(entry.workbook)) {
      if (!ddhNames.has(key)) ddhNames.set(key, value);
    }
  }
  for (const entry of caseyEntries) {
    for (const [key, value] of extractCaseyNames(entry.workbook)) {
      if (!caseyNames.has(key)) caseyNames.set(key, value);
    }
  }
  for (const entry of mchEntries) {
    for (const [key, value] of extractMchNames(entry.workbook)) {
      if (!mchNames.has(key)) mchNames.set(key, value);
    }
  }
  return mergedDoctorOptions([
    ["mmc", mmcNames],
    ["ddh", ddhNames],
    ["casey", caseyNames],
    ["mch", mchNames],
  ]);
}

function mergedDoctorOptions(sourceMaps) {
  const groups = new Map();
  for (const [sourceType, names] of sourceMaps) {
    for (const [key, displayName] of names) {
      const identity = rosterIdentityKey(displayName || key);
      if (!identity) continue;
      if (!groups.has(identity)) groups.set(identity, []);
      groups.get(identity).push({ sourceType, key, displayName });
    }
  }
  return [...groups.values()].map((aliases) => {
    aliases.sort((left, right) => sourcePriority(left.sourceType) - sourcePriority(right.sourceType) || left.displayName.localeCompare(right.displayName));
    const primary = aliases[0];
    return {
      key: primary.key,
      displayName: formatDoctorDisplayName(primary.displayName),
      sourceTypes: [...new Set(aliases.map((alias) => alias.sourceType))],
      aliases: aliases.map((alias) => ({ ...alias, displayName: formatDoctorDisplayName(alias.displayName) })),
    };
  }).sort((left, right) => left.displayName.localeCompare(right.displayName));
}

function sourcePriority(sourceType) {
  return { mmc: 0, ddh: 1, casey: 2, mch: 3 }[sourceType] ?? 99;
}

export function buildRosterView(mmcSources, ddhSources, doctorKey, settings = DEFAULT_SETTINGS, overrides = {}, conflictSelections = {}, doctorAliases = [], caseySources = [], mchSources = []) {
  const records = [];
  const keysBySource = doctorKeysBySource(doctorKey, doctorAliases);
  for (const entry of normalizeSourceEntries(mmcSources)) {
    for (const key of keysBySource.mmc) {
      records.push(...attachImportMeta(parseMmcRecords(entry.workbook, key), entry));
    }
  }
  for (const entry of normalizeSourceEntries(ddhSources)) {
    for (const key of keysBySource.ddh) {
      records.push(...attachImportMeta(parseDdhRecords(entry.workbook, key), entry));
    }
  }
  for (const entry of normalizeSourceEntries(caseySources)) {
    for (const key of keysBySource.casey) {
      records.push(...attachImportMeta(parseCaseyRecords(entry.workbook, key), entry));
    }
  }
  for (const entry of normalizeSourceEntries(mchSources)) {
    for (const key of keysBySource.mch) {
      records.push(...attachImportMeta(parseMchRecords(entry.workbook, key), entry));
    }
  }

  records.sort((left, right) => {
    if (left.startDay !== right.startDay) return left.startDay.localeCompare(right.startDay);
    if (left.source !== right.source) return left.source.localeCompare(right.source);
    return left.rawValue.localeCompare(right.rawValue);
  });

  const merge = mergeRecordsAcrossImports(records, conflictSelections);
  const view = applySettings(merge.records, sanitizeSettings(settings), sanitizeOverrides(overrides));
  return {
    ...view,
    conflicts: merge.conflicts,
    imports: summarizeImports([...normalizeSourceEntries(mmcSources), ...normalizeSourceEntries(ddhSources), ...normalizeSourceEntries(caseySources), ...normalizeSourceEntries(mchSources)]),
  };
}

export function generateEvents(mmcSources, ddhSources, doctorKey, settings = DEFAULT_SETTINGS, overrides = {}, conflictSelections = {}, caseySources = [], mchSources = []) {
  return buildRosterView(mmcSources, ddhSources, doctorKey, settings, overrides, conflictSelections, [], caseySources, mchSources).events;
}

export async function inspectImportRecord(record) {
  if (!record?.dataUrl) {
    throw new Error("Import data is required for repository inspection.");
  }
  const workbook = await readWorkbookDataUrl(record.dataUrl, record.name || "roster.xlsx");
  const sourceType = detectSourceType(workbook, record.name || "roster.xlsx");
  validateRosterDatesWithinSingleTerm(workbook, sourceType, record.name || "roster.xlsx");
  const entry = {
    id: String(record.id || ""),
    addedAt: String(record.addedAt || ""),
    file: {
      name: String(record.name || "roster.xlsx"),
      size: Number(record.size || 0),
      lastModified: Number(record.lastModified || 0),
    },
    workbook,
  };
  const rosterDoctors = doctorOptions(sourceType === "mmc" ? [entry] : [], sourceType === "ddh" ? [entry] : [], sourceType === "casey" ? [entry] : [], sourceType === "mch" ? [entry] : [])
    .map((doctor) => ({
      key: doctor.key,
      displayName: doctor.displayName,
      sourceType,
    }));
  const providerDoctors = sourceType === "ddh" ? findmyshiftProviderStaffOptions([entry]) : [];
  return { sourceType, doctors: mergeMembershipDoctors(rosterDoctors, providerDoctors) };
}

export function findmyshiftProviderStaffOptions(entries = []) {
  const doctors = [];
  for (const entry of entries) {
    const sheet = entry?.workbook?.Sheets?.["FindMyShift staff"];
    if (!sheet) continue;
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:C1");
    for (let row = 1; row <= range.e.r; row += 1) {
      const name = cleanText(getCellValue(sheet, row + 1, 2));
      if (!name) continue;
      doctors.push({
        key: normalizeName(name),
        displayName: name,
        sourceType: "ddh",
        seniority: canonicalFindmyshiftProviderSeniority(getCellValue(sheet, row + 1, 3)),
        membershipSource: "provider",
      });
    }
  }
  return mergeMembershipDoctors([], doctors);
}

function canonicalFindmyshiftProviderSeniority(value) {
  const supplied = cleanText(value);
  if (/\b(?:SMS|SENIOR\s+MEDICAL|CONSULTANT|SPECIALIST)\b/i.test(supplied)) return "SMS";
  return sanitizeRuleSeniority(findmyshiftDdhSeniority(supplied, ""));
}

export function mergeMembershipDoctors(rosterDoctors = [], providerDoctors = []) {
  const byKey = new Map();
  for (const doctor of [...rosterDoctors, ...providerDoctors]) {
    const key = normalizeName(doctor?.key || doctor?.displayName || "");
    if (!key) continue;
    const previous = byKey.get(key);
    byKey.set(key, {
      key,
      displayName: String(doctor?.displayName || previous?.displayName || key).trim(),
      sourceType: String(doctor?.sourceType || previous?.sourceType || "").toLowerCase(),
      seniority: sanitizeRuleSeniority(doctor?.seniority || previous?.seniority || UNKNOWN_SENIORITY),
      membershipSource: doctor?.membershipSource === "provider" || previous?.membershipSource === "provider" ? "provider" : "roster",
    });
  }
  return [...byKey.values()];
}

export function normalizeRosterName(value) {
  return normalizeName(value);
}

export function previewSummary(events) {
  if (!events.length) {
    return { count: 0, date_range: "No events found" };
  }
  const first = asDateString(events[0].start);
  const lastEvent = events[events.length - 1];
  let last = asDateString(lastEvent.end);
  if (lastEvent.allDay) {
    last = addDays(last, -1);
  }
  return {
    count: events.length,
    date_range: `${first} to ${last}`,
  };
}

export function exportIcs(events, doctorDisplayName) {
  const dtstamp = new Date().toISOString().replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z");
  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Roster Converter//EN",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
    `X-WR-CALNAME:Roster - ${escapeIcsText(doctorDisplayName)}`,
    `X-WR-TIMEZONE:${TIMEZONE}`,
  ];

  for (const event of events) {
    const uid = `${hashString(`${event.source}|${event.title}|${event.start}|${event.end}|${event.location || ""}`)}@roster-converter`;
    lines.push("BEGIN:VEVENT");
    lines.push(`UID:${uid}`);
    lines.push(`DTSTAMP:${dtstamp}`);
    lines.push(`SUMMARY:${escapeIcsText(event.title)}`);
    if (event.allDay) {
      lines.push(`DTSTART;VALUE=DATE:${event.start.replace(/-/g, "")}`);
      lines.push(`DTEND;VALUE=DATE:${event.end.replace(/-/g, "")}`);
    } else {
      lines.push(`DTSTART;TZID=${TIMEZONE}:${toIcsDateTime(event.start)}`);
      lines.push(`DTEND;TZID=${TIMEZONE}:${toIcsDateTime(event.end)}`);
    }
    if (event.location) {
      lines.push(`LOCATION:${escapeIcsText(event.location)}`);
    }
    lines.push("END:VEVENT");
  }

  lines.push("END:VCALENDAR");
  return `${lines.join("\r\n")}\r\n`;
}

export function serializeEvent(event) {
  return {
    id: event.id,
    source: event.source,
    sources: Array.isArray(event.sources) ? event.sources : undefined,
    seniority: event.seniority || "",
    title: event.title,
    allDay: event.allDay,
    start: event.start,
    end: event.end,
    location: event.location || "",
    rawValue: event.rawValue,
    timeLabel: event.timeLabel,
    monthKey: event.monthKey,
  };
}

export function serializeConflict(conflict) {
  return conflict;
}

export function serializeReviewItem(item) {
  return {
    id: item.id,
    source: item.source,
    seniority: item.seniority || "",
    startDay: item.startDay,
    endDay: item.endDay,
    rawValue: item.rawValue,
    normalizedTitle: item.normalizedTitle,
    suggestedTitle: item.suggestedTitle,
    overrideTitle: item.overrideTitle,
    status: item.status,
    warnings: item.warnings,
    include: item.include,
    exportable: item.exportable,
    location: item.location || "",
    allDay: item.allDay,
    timeLabel: item.timeLabel,
  };
}

export function sourceNames(sources) {
  return {
    mmc: normalizeSourceEntries(sources.mmc).map((entry) => entry.file.name),
    ddh: normalizeSourceEntries(sources.ddh).map((entry) => entry.file.name),
    casey: normalizeSourceEntries(sources.casey).map((entry) => entry.file.name),
    mch: normalizeSourceEntries(sources.mch).map((entry) => entry.file.name),
  };
}

function hasAnyRosterSource(sources) {
  return Boolean(sources?.mmc?.length || sources?.ddh?.length || sources?.casey?.length || sources?.mch?.length);
}

function normalizeSourceEntries(value) {
  if (!value) return [];
  const items = Array.isArray(value) ? value : [value];
  return items.map((entry, index) => {
    if (entry?.workbook) return entry;
    return {
      id: `legacy-${index}`,
      addedAt: "",
      file: { name: `import-${index + 1}.xlsx`, size: 0, lastModified: 0 },
      workbook: entry,
    };
  });
}

function isRosterSourceType(value) {
  const source = String(value || "").toLowerCase();
  return source === "mmc" || source === "ddh" || source === "casey" || source === "mch";
}

function attachImportMeta(records, entry) {
  return records.map((record) => ({
    ...record,
    importId: entry.id,
    importName: entry.file.name,
    importAddedAt: entry.addedAt || "",
    weekKey: mondayOfDay(record.startDay),
    dedupKey: hashString(`${record.source}|${record.kind}|${record.rawValue}|${record.start}|${record.end}|${record.location}|${record.normalizedTitle}`),
  }));
}

function summarizeImports(entries) {
  return entries
    .map((entry) => ({
      id: entry.id,
      name: entry.file.name,
      sourceType: detectSourceType(entry.workbook, entry.file.name),
      addedAt: entry.addedAt || "",
      size: entry.file.size,
      lastModified: entry.file.lastModified,
    }))
    .sort((left, right) => (left.addedAt || "").localeCompare(right.addedAt || "") || left.name.localeCompare(right.name));
}

function mergeRecordsAcrossImports(records, rawSelections = {}) {
  const recordsByGroup = new Map();
  for (const record of records) {
    const key = `${record.source}|${record.weekKey}`;
    if (!recordsByGroup.has(key)) recordsByGroup.set(key, []);
    recordsByGroup.get(key).push(record);
  }

  const mergedRecords = [];
  const conflicts = [];
  for (const [groupKey, groupRecords] of recordsByGroup.entries()) {
    const imports = new Map();
    for (const record of groupRecords) {
      if (!imports.has(record.importId)) imports.set(record.importId, []);
      imports.get(record.importId).push(record);
    }

    if (imports.size === 1) {
      mergedRecords.push(...dedupeRecords(groupRecords));
      continue;
    }

    const importEntries = [...imports.entries()].map(([importId, importRecords]) => {
      const sample = importRecords[0];
      const signature = importRecords.map((record) => record.dedupKey).sort().join("|");
      return {
        importId,
        importName: sample.importName,
        importAddedAt: sample.importAddedAt,
        source: sample.source,
        weekKey: sample.weekKey,
        records: importRecords,
        signature,
      };
    }).sort(compareImportEntries);

    const uniqueSignatures = new Set(importEntries.map((entry) => entry.signature));
    if (uniqueSignatures.size === 1) {
      mergedRecords.push(...dedupeRecords(groupRecords));
      continue;
    }

    const winner = chooseWinningImport(importEntries, rawSelections[groupKey]);
    mergedRecords.push(...winner.records);
    conflicts.push({
      key: groupKey,
      source: winner.source,
      weekKey: winner.weekKey,
      selectedImportId: winner.importId,
      options: importEntries.map((entry) => ({
        importId: entry.importId,
        importName: entry.importName,
        addedAt: entry.importAddedAt,
        eventCount: entry.records.length,
      })),
    });
  }

  return {
    records: dedupeRecords(mergedRecords),
    conflicts: conflicts.sort((left, right) => left.weekKey.localeCompare(right.weekKey) || left.source.localeCompare(right.source)),
  };
}

function dedupeRecords(records) {
  const seen = new Set();
  const deduped = [];
  for (const record of records) {
    if (seen.has(record.dedupKey)) continue;
    seen.add(record.dedupKey);
    deduped.push(record);
  }
  return deduped;
}

function compareImportEntries(left, right) {
  const leftDate = left.importAddedAt || "";
  const rightDate = right.importAddedAt || "";
  if (leftDate !== rightDate) return rightDate.localeCompare(leftDate);
  return right.importName.localeCompare(left.importName);
}

function chooseWinningImport(importEntries, selectedImportId) {
  if (selectedImportId) {
    const explicit = importEntries.find((entry) => entry.importId === selectedImportId);
    if (explicit) return explicit;
  }
  return importEntries[0];
}

async function readWorkbook(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  if (isPdfFile(file.name, bytes)) {
    await ensurePdfDependencies();
    return readPdfWorkbook(bytes, file.name);
  }

  await ensureSpreadsheetDependency();
  return readSpreadsheetWorkbook(bytes, file.name);
}

async function readWorkbookDataUrl(dataUrl, filename) {
  const bytes = bytesFromDataUrl(dataUrl);
  if (isPdfFile(filename, bytes)) {
    await ensurePdfDependencies();
    return readPdfWorkbook(bytes, filename);
  }

  await ensureSpreadsheetDependency();
  return readSpreadsheetWorkbook(bytes, filename);
}

function readSpreadsheetWorkbook(bytes, filename) {
  const baseOptions = {
    type: "array",
    cellDates: true,
    cellNF: false,
    cellHTML: false,
    cellStyles: false,
  };

  try {
    const metadata = XLSX.read(bytes, { type: "array", bookSheets: true });
    const sheetNames = metadata.SheetNames || [];
    const weekSheets = sheetNames.filter((name) => name.startsWith("Week "));
    if (sheetNames.includes("Whole thing") && weekSheets.length) {
      const workbook = XLSX.read(bytes, { ...baseOptions, sheets: weekSheets });
      workbook.SheetNames = sheetNames;
      workbook.Sheets["Whole thing"] = workbook.Sheets["Whole thing"] || {};
      return workbook;
    }
    return XLSX.read(bytes, baseOptions);
  } catch {
    throw new Error(`${filename} is not a supported MMC workbook, MMC PDF, Dandenong Hospital FindMyShift export, Casey roster, or MCH roster.`);
  }
}

function bytesFromDataUrl(dataUrl) {
  const value = String(dataUrl || "");
  const [, payload = ""] = value.split(",", 2);
  const binary = atob(payload);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function detectSourceType(workbook, filename) {
  const sheetNames = workbook.SheetNames || [];
  if (sheetNames.includes("Whole thing") && sheetNames.some((name) => name.startsWith("Week "))) {
    return "mmc";
  }
  if (isMchWorkbook(workbook)) {
    return "mch";
  }
  const sheet = workbook.Sheets[sheetNames[0]];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  if (range.e.c + 1 >= 8 && [1, 2, 3, 4].some((row) => isDdhDateRow(sheet, row))) {
    return "ddh";
  }
  if (isCaseyWorkbook(workbook)) {
    return "casey";
  }
  throw new Error(`${filename} is not a supported MMC workbook, MMC PDF, Dandenong Hospital FindMyShift export, Casey roster, or MCH roster.`);
}

function validateRosterDatesWithinSingleTerm(workbook, sourceType, filename) {
  // FindMyShift reports are deliberately retained as one current, available
  // roster range.  Unlike a manually uploaded term workbook, that range can
  // cross a Victorian term boundary and must remain one coherent source.
  if (sourceType === "ddh" && isFindmyshiftDdhWorkbook(workbook)) return;
  const evidence = collectRosterDateEvidence(workbook, sourceType);
  if (!evidence.length) return;
  const groups = new Map();
  for (const item of evidence) {
    const term = australianTermForDate(parseDateOnly(item.date));
    const key = `${term.termNumber}-${term.year}`;
    if (!groups.has(key)) {
      groups.set(key, {
        ...term,
        label: formatAustralianTermLabel(term),
        items: [],
      });
    }
    groups.get(key).items.push(item);
  }
  if (groups.size <= 1) return;
  const dominant = [...groups.values()].sort((left, right) => (
    right.items.length - left.items.length
    || left.year - right.year
    || left.termNumber - right.termNumber
  ))[0];
  const conflicts = [...groups.values()]
    .filter((group) => group !== dominant)
    .flatMap((group) => group.items.map((item) => ({ ...item, termLabel: group.label })))
    .sort((left, right) => left.date.localeCompare(right.date) || left.sheetName.localeCompare(right.sheetName) || left.cell.localeCompare(right.cell));
  const shown = conflicts.slice(0, 3).map((item) => `${item.sheetName ? `${item.sheetName} ` : ""}${item.cell ? `cell ${item.cell} ` : ""}is ${item.date} (${item.termLabel})`);
  const remaining = conflicts.length > shown.length ? `, plus ${conflicts.length - shown.length} more conflicting date${conflicts.length - shown.length === 1 ? "" : "s"}` : "";
  throw new Error(`${filename} has dates from multiple terms. Most dates are ${dominant.label}, but ${shown.join("; ")}${remaining}. Fix the conflicting worksheet/date and upload again.`);
}

function isFindmyshiftDdhWorkbook(workbook) {
  const sheetName = workbook?.SheetNames?.[0] || "";
  const sheet = workbook?.Sheets?.[sheetName];
  return Boolean(sheet) && cleanText(getCellValue(sheet, 1, 1)) === "FindMyShift roster format";
}

function collectRosterDateEvidence(workbook, sourceType) {
  if (sourceType === "mmc") return collectMmcDateEvidence(workbook);
  if (sourceType === "ddh") return collectDdhDateEvidence(workbook);
  if (sourceType === "casey") return collectCaseyDateEvidence(workbook);
  if (sourceType === "mch") return collectMchDateEvidence(workbook);
  return [];
}

function collectMmcDateEvidence(workbook) {
  const evidence = [];
  for (const sheetName of workbook.SheetNames || []) {
    if (!sheetName.startsWith("Week ")) continue;
    const sheet = workbook.Sheets[sheetName];
    const layout = mmcWeekLayout(sheet);
    if (!layout) continue;
    layout.weekDates.forEach((date, index) => {
      evidence.push(dateEvidence(sheetName, layout.dateRow, layout.dayColumns[index], date));
    });
  }
  return evidence;
}

const MMC_WEEKDAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"];

function mmcWeekLayout(sheet) {
  const range = XLSX.utils.decode_range(sheet?.["!ref"] || "A1:A1");
  const lastRow = range.e.r + 1;
  const lastColumn = range.e.c + 1;

  // MMC occasionally inserts administrative columns (for example Pager No).
  // Locate the weekday and person headers instead of assuming their positions.
  for (let headerRow = 1; headerRow <= Math.min(lastRow, 12); headerRow += 1) {
    for (let firstDayColumn = 1; firstDayColumn <= lastColumn - 6; firstDayColumn += 1) {
      const dayColumns = MMC_WEEKDAYS.map((_, index) => firstDayColumn + index);
      if (!dayColumns.every((column, index) => mmcHeaderText(getCellValue(sheet, headerRow, column)) === MMC_WEEKDAYS[index])) continue;
      const markerColumn = mmcColumnWithHeader(sheet, headerRow, "cost centre")
        || mmcColumnWithHeader(sheet, headerRow, "cost center")
        || mmcColumnWithHeader(sheet, headerRow, "seniority")
        || mmcColumnWithHeader(sheet, headerRow, "role");
      const nameColumn = mmcColumnWithHeader(sheet, headerRow, "name");
      if (!nameColumn) continue;
      for (let dateRow = headerRow + 1; dateRow <= Math.min(lastRow, headerRow + 4); dateRow += 1) {
        const weekDates = dayColumns.map((column) => coerceDate(getCellValue(sheet, dateRow, column)));
        if (weekDates.every(Boolean)) return { dateRow, dayColumns, weekDates, markerColumn, nameColumn };
      }
    }
  }

  // Preserve compatibility with older MMC exports that predate column headings.
  // New and changed layouts always use the header-based path above.
  const legacyDates = [];
  for (let column = 6; column <= 12; column += 1) {
    const date = coerceDate(getCellValue(sheet, 4, column));
    if (!date) return null;
    legacyDates.push(date);
  }
  return { dateRow: 4, dayColumns: [6, 7, 8, 9, 10, 11, 12], weekDates: legacyDates, markerColumn: 3, nameColumn: 4 };
}

function mmcColumnWithHeader(sheet, row, header) {
  const range = XLSX.utils.decode_range(sheet?.["!ref"] || "A1:A1");
  for (let column = 1; column <= range.e.c + 1; column += 1) {
    const value = mmcHeaderText(getCellValue(sheet, row, column));
    if (value === header || value.startsWith(`${header} `)) return column;
  }
  return 0;
}

function mmcHeaderText(value) {
  return cleanText(value)
    .toLowerCase()
    .replace(/\([^)]*\)/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function collectDdhDateEvidence(workbook) {
  const sheetName = workbook.SheetNames?.[0] || "";
  const sheet = workbook.Sheets?.[sheetName];
  if (!sheet) return [];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const evidence = [];
  for (let row = 1; row <= range.e.r + 1; row += 1) {
    if (!isDdhDateRow(sheet, row)) continue;
    for (let col = 2; col <= 8; col += 1) {
      const date = parseDdhDate(getCellValue(sheet, row, col));
      if (date) evidence.push(dateEvidence(sheetName, row, col, date));
    }
  }
  return evidence;
}

function collectCaseyDateEvidence(workbook) {
  const evidence = [];
  for (const sheetName of workbook.SheetNames || []) {
    const sheet = workbook.Sheets[sheetName];
    if (!isCaseyWeekSheet(sheet, sheetName)) continue;
    const termYear = caseyTermYear(sheet);
    for (let col = 2; col <= 8; col += 1) {
      const date = parseCaseyWeekDate(sheet, 2, col, termYear);
      if (date) evidence.push(dateEvidence(sheetName, 2, col, date));
    }
  }
  return evidence;
}

function collectMchDateEvidence(workbook) {
  const evidence = [];
  for (const sheetName of workbook.SheetNames || []) {
    const sheet = workbook.Sheets[sheetName];
    if (!isMchWeekSheet(sheet, sheetName)) continue;
    for (let col = 6; col <= 12; col += 1) {
      const primaryDate = coerceDate(getCellValue(sheet, 19, col));
      const fallbackDate = primaryDate || coerceDate(getCellValue(sheet, 2, col));
      if (fallbackDate) evidence.push(dateEvidence(sheetName, primaryDate ? 19 : 2, col, fallbackDate));
    }
  }
  return evidence;
}

function dateEvidence(sheetName, row, col, date) {
  return {
    sheetName: String(sheetName || ""),
    cell: XLSX.utils.encode_cell({ r: row - 1, c: col - 1 }),
    date,
  };
}

function isPdfFile(filename, bytes) {
  if (String(filename || "").toLowerCase().endsWith(".pdf")) return true;
  return bytes?.[0] === 0x25 && bytes?.[1] === 0x50 && bytes?.[2] === 0x44 && bytes?.[3] === 0x46;
}

function readPdfWorkbook(bytes, filename) {
  const text = latin1FromBytes(bytes);
  if (!text.startsWith("%PDF")) {
    throw new Error(`${filename} is not a valid PDF roster.`);
  }

  const objects = parsePdfObjects(text);
  const fontMaps = parsePdfFontMaps(objects);
  const pages = parsePdfPages(objects, fontMaps);
  const workbook = mmcWorkbookFromPdfPages(pages);
  if (!workbook.SheetNames.length) {
    throw new Error(`${filename} does not look like an MMC roster PDF.`);
  }
  return workbook;
}

function parsePdfObjects(pdfText) {
  const objects = new Map();
  const pattern = /^(\d+)\s+0\s+obj\r?\n([\s\S]*?)\r?\nendobj/gm;
  let match;
  while ((match = pattern.exec(pdfText))) {
    objects.set(Number(match[1]), match[2]);
  }
  return objects;
}

function parsePdfFontMaps(objects) {
  const fontObjectByAlias = new Map();
  for (const body of objects.values()) {
    const fontBlock = body.match(/\/Font\s*<<([\s\S]*?)>>/);
    if (!fontBlock) continue;
    const fontPattern = /\/(TT\d+)\s+(\d+)\s+0\s+R/g;
    let match;
    while ((match = fontPattern.exec(fontBlock[1]))) {
      fontObjectByAlias.set(match[1], Number(match[2]));
    }
  }

  const maps = new Map();
  for (const [alias, objectId] of fontObjectByAlias.entries()) {
    const fontBody = objects.get(objectId) || "";
    const unicodeMatch = fontBody.match(/\/ToUnicode\s+(\d+)\s+0\s+R/);
    if (!unicodeMatch) continue;
    const cmapText = inflatePdfStream(objects.get(Number(unicodeMatch[1])) || "");
    maps.set(alias, parseToUnicodeCMap(cmapText));
  }
  return maps;
}

function parseToUnicodeCMap(cmapText) {
  const map = new Map();
  const charPattern = /^\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*$/gm;
  let charMatch;
  while ((charMatch = charPattern.exec(cmapText))) {
    map.set(Number.parseInt(charMatch[1], 16), String.fromCodePoint(Number.parseInt(charMatch[2], 16)));
  }

  const rangePattern = /<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>/g;
  let match;
  while ((match = rangePattern.exec(cmapText))) {
    const start = Number.parseInt(match[1], 16);
    const end = Number.parseInt(match[2], 16);
    const base = Number.parseInt(match[3], 16);
    for (let code = start; code <= end; code += 1) {
      map.set(code, String.fromCodePoint(base + code - start));
    }
  }
  return map;
}

function parsePdfPages(objects, fontMaps) {
  const pages = [];
  for (const [objectId, body] of objects.entries()) {
    if (!/\/Type\s*\/Page\b/.test(body)) continue;
    const contentMatch = body.match(/\/Contents\s+(\d+)\s+0\s+R/);
    if (!contentMatch) continue;
    const content = inflatePdfStream(objects.get(Number(contentMatch[1])) || "");
    const items = extractPdfTextItems(content, fontMaps);
    if (items.some((item) => item.text.includes("MMC ED ADULT ROSTER"))) {
      pages.push({ objectId, items });
    }
  }
  return pages.sort((left, right) => left.objectId - right.objectId);
}

function extractPdfTextItems(content, fontMaps) {
  const items = [];
  const blockPattern = /BT\s+([\s\S]*?)\s+ET/g;
  let match;
  while ((match = blockPattern.exec(content))) {
    const block = match[1];
    const transform = lastMatch(block, /([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+Tm/g);
    const font = lastMatch(block, /\/(TT\d+)\s+[\d.]+\s+Tf/g);
    if (!transform || !font) continue;
    const fontMap = fontMaps.get(font[1]);
    if (!fontMap) continue;
    const text = decodePdfTextBlock(block, fontMap);
    if (!text) continue;
    items.push({
      x: Number(transform[5]),
      y: Number(transform[6]),
      font: font[1],
      text,
    });
  }
  return items;
}

function lastMatch(value, pattern) {
  let result = null;
  let match;
  while ((match = pattern.exec(value))) {
    result = match;
  }
  return result;
}

function decodePdfTextBlock(block, fontMap) {
  const fragments = [];
  for (const bytes of pdfStringFragments(block)) {
    fragments.push([...bytes].map((byte) => fontMap.get(byte) || "").join(""));
  }
  return fragments.join("").replace(/\s+/g, " ").trim();
}

function pdfStringFragments(value) {
  const fragments = [];
  for (let index = 0; index < value.length; index += 1) {
    if (value[index] !== "(") continue;
    const bytes = [];
    index += 1;
    while (index < value.length) {
      const char = value.charCodeAt(index);
      if (char === 0x5c) {
        const parsed = parsePdfEscape(value, index);
        if (parsed.byte !== null) bytes.push(parsed.byte);
        index = parsed.nextIndex;
        continue;
      }
      if (char === 0x29) break;
      bytes.push(char & 0xff);
      index += 1;
    }
    fragments.push(bytes);
  }
  return fragments;
}

function parsePdfEscape(value, index) {
  const nextIndex = index + 1;
  if (nextIndex >= value.length) return { byte: null, nextIndex };
  const next = value[nextIndex];
  const simple = { n: 0x0a, r: 0x0d, t: 0x09, b: 0x08, f: 0x0c, "(": 0x28, ")": 0x29, "\\": 0x5c };
  if (Object.prototype.hasOwnProperty.call(simple, next)) {
    return { byte: simple[next], nextIndex: nextIndex + 1 };
  }
  if (/[0-7]/.test(next)) {
    let digits = next;
    let cursor = nextIndex + 1;
    while (cursor < value.length && digits.length < 3 && /[0-7]/.test(value[cursor])) {
      digits += value[cursor];
      cursor += 1;
    }
    return { byte: Number.parseInt(digits, 8) & 0xff, nextIndex: cursor };
  }
  return { byte: next.charCodeAt(0) & 0xff, nextIndex: nextIndex + 1 };
}

function mmcWorkbookFromPdfPages(pages) {
  const workbook = { SheetNames: [], Sheets: {} };
  pages.forEach((page) => {
    const sheet = mmcSheetFromPdfPage(page.items);
    if (!sheet) return;
    const sheetName = `Week ${workbook.SheetNames.length + 1}`;
    workbook.SheetNames.push(sheetName);
    workbook.Sheets[sheetName] = sheet;
  });
  if (workbook.SheetNames.length) {
    workbook.SheetNames.unshift("Whole thing");
    workbook.Sheets["Whole thing"] = {};
  }
  return workbook;
}

function mmcSheetFromPdfPage(items) {
  const dateItems = items
    .filter((item) => /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(item.text))
    .sort((left, right) => left.x - right.x)
    .slice(0, 7);
  if (dateItems.length !== 7) return null;

  const rowAnchors = extractMmcPdfRowAnchors(items, Math.min(...dateItems.map((item) => item.y)));
  if (!rowAnchors.length) return null;

  const rows = Array.from({ length: rowAnchors.length + 6 }, () => []);
  dateItems.forEach((item, index) => {
    rows[3][5 + index] = parseAustralianDate(item.text);
  });

  rowAnchors.forEach((anchor, index) => {
    const rowIndex = 6 + index;
    if (anchor.type === "section") {
      rows[rowIndex][2] = anchor.text.toUpperCase();
    } else {
      rows[rowIndex][3] = anchor.text;
    }
  });

  const dayCenters = dateItems.map((item) => item.x);
  for (const item of items) {
    const colIndex = nearestIndex(dayCenters, item.x, 38);
    if (colIndex < 0) continue;
    const rowIndex = nearestIndex(rowAnchors.map((row) => row.y), item.y, 5);
    if (rowIndex < 0) continue;
    const targetRow = 6 + rowIndex;
    const targetCol = 5 + colIndex;
    rows[targetRow][targetCol] = [rows[targetRow][targetCol], item.text].filter(Boolean).join(" ").trim();
  }

  return XLSX.utils.aoa_to_sheet(rows, { cellDates: true });
}

function extractMmcPdfRowAnchors(items, dateHeaderY) {
  const rowItems = items
    .filter((item) => item.x < 160 && item.y < dateHeaderY - 5)
    .filter((item) => looksLikePersonName(item.text) || isMmcSectionMarker(item.text))
    .sort((left, right) => right.y - left.y);
  const seen = new Set();
  const anchors = [];
  for (const item of rowItems) {
    const yKey = Math.round(item.y);
    if (seen.has(yKey)) continue;
    seen.add(yKey);
    anchors.push({
      y: item.y,
      text: item.text,
      type: isMmcSectionMarker(item.text) ? "section" : "name",
    });
  }
  return anchors;
}

function parseAustralianDate(value) {
  const match = String(value).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!match) return null;
  return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
}

function nearestIndex(values, target, tolerance) {
  let bestIndex = -1;
  let bestDistance = Infinity;
  values.forEach((value, index) => {
    const distance = Math.abs(value - target);
    if (distance < bestDistance) {
      bestDistance = distance;
      bestIndex = index;
    }
  });
  return bestDistance <= tolerance ? bestIndex : -1;
}

function inflatePdfStream(objectBody) {
  const match = objectBody.match(/stream\r?\n([\s\S]*?)\r?\nendstream/);
  if (!match) return "";
  const bytes = bytesFromLatin1(match[1]);
  try {
    return latin1FromBytes(decompressSync(bytes));
  } catch {
    return "";
  }
}

function latin1FromBytes(bytes) {
  let result = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    result += String.fromCharCode(...bytes.slice(index, index + chunkSize));
  }
  return result;
}

function bytesFromLatin1(value) {
  const bytes = new Uint8Array(value.length);
  for (let index = 0; index < value.length; index += 1) {
    bytes[index] = value.charCodeAt(index) & 0xff;
  }
  return bytes;
}

function sanitizeDoctorAliases(raw) {
  if (!Array.isArray(raw)) return [];
  return raw
    .map((item) => ({
      sourceType: String(item?.sourceType || "").toLowerCase(),
      key: normalizeName(item?.key || ""),
      displayName: String(item?.displayName || "").trim(),
    }))
    .filter((item) => isRosterSourceType(item.sourceType) && item.key);
}

function doctorKeysBySource(doctorKey, rawAliases = []) {
  const fallback = normalizeName(doctorKey || "");
  const aliases = sanitizeDoctorAliases(rawAliases);
  const mmc = new Set();
  const ddh = new Set();
  const casey = new Set();
  const mch = new Set();
  for (const alias of aliases) {
    if (alias.sourceType === "mmc") mmc.add(alias.key);
    if (alias.sourceType === "ddh") ddh.add(alias.key);
    if (alias.sourceType === "casey") casey.add(alias.key);
    if (alias.sourceType === "mch") mch.add(alias.key);
  }
  if (fallback) {
    if (!mmc.size) mmc.add(fallback);
    if (!ddh.size) ddh.add(fallback);
    if (!casey.size) casey.add(fallback);
    if (!mch.size) mch.add(fallback);
  }
  return { mmc: [...mmc], ddh: [...ddh], casey: [...casey], mch: [...mch] };
}

function parseJsonField(formData, fieldName, fallback) {
  const raw = formData.get(fieldName);
  if (!raw) return fallback;
  try {
    return JSON.parse(String(raw));
  } catch {
    return fallback;
  }
}

function sanitizeSettings(raw) {
  const input = typeof raw === "object" && raw ? raw : {};
  return {
    showSourcePrefix: input.showSourcePrefix !== false,
    showAmPm: input.showAmPm !== false,
    showTimes: input.showTimes !== false,
    showRawValues: input.showRawValues === true,
    showNormalizedTitles: input.showNormalizedTitles !== false,
    includeLocations: input.includeLocations !== false,
    includeAnnualLeave: input.includeAnnualLeave !== false,
    includeConferenceLeave: input.includeConferenceLeave !== false,
    includePublicHoliday: input.includePublicHoliday !== false,
    includeSickLeave: input.includeSickLeave !== false,
    defaultLocationMmc: sanitizeLocationSetting(input.defaultLocationMmc, DEFAULT_SETTINGS.defaultLocationMmc),
    defaultLocationDdh: sanitizeLocationSetting(input.defaultLocationDdh, DEFAULT_SETTINGS.defaultLocationDdh),
    defaultLocationCasey: sanitizeLocationSetting(input.defaultLocationCasey, DEFAULT_SETTINGS.defaultLocationCasey),
    defaultLocationMch: sanitizeLocationSetting(input.defaultLocationMch, DEFAULT_SETTINGS.defaultLocationMch),
    hospitalFilter: isRosterSourceType(input.hospitalFilter) ? input.hospitalFilter : "all",
    dateFrom: isDateString(input.dateFrom) ? input.dateFrom : "",
    dateTo: isDateString(input.dateTo) ? input.dateTo : "",
  };
}

function sanitizeLocationSetting(value, fallback) {
  const next = String(value || "").trim();
  return next || fallback;
}

function sanitizeOverrides(raw) {
  const overrides = {};
  if (!raw || typeof raw !== "object") return overrides;
  for (const [key, value] of Object.entries(raw)) {
    if (!value || typeof value !== "object") continue;
    const next = {
      title: typeof value.title === "string" ? value.title.trim() : "",
      include: typeof value.include === "boolean" ? value.include : undefined,
      start: typeof value.start === "string" ? value.start.trim() : "",
      end: typeof value.end === "string" ? value.end.trim() : "",
      allDay: typeof value.allDay === "boolean" ? value.allDay : undefined,
    };
    if (Object.prototype.hasOwnProperty.call(value, "location")) {
      next.location = typeof value.location === "string" ? value.location.trim() : "";
    }
    overrides[key] = next;
  }
  return overrides;
}

function sanitizeCustomEvents(raw) {
  if (!Array.isArray(raw)) return [];
  const events = [];
  for (const item of raw) {
    if (!item || typeof item !== "object") continue;
    const title = typeof item.title === "string" ? item.title.trim() : "";
    const startDate = isDateString(item.startDate) ? item.startDate : "";
    const endDate = isDateString(item.endDate) ? item.endDate : startDate;
    const allDay = item.allDay === true;
    const startTime = isClockString(item.startTime) ? item.startTime : "";
    const endTime = isClockString(item.endTime) ? item.endTime : "";
    if (!title || !startDate || !endDate) continue;
    if (!allDay && (!startTime || !endTime)) continue;
    events.push({
      id: typeof item.id === "string" && item.id ? item.id : hashString(`custom|${title}|${startDate}|${endDate}|${startTime}|${endTime}`),
      title,
      startDate,
      endDate,
      allDay,
      startTime,
      endTime,
      location: typeof item.location === "string" ? item.location.trim() : "",
      include: item.include !== false,
    });
  }
  return latestCustomEventsById(events);
}

function latestCustomEventsById(events) {
  const byId = new Map();
  for (const event of events || []) {
    byId.delete(event.id);
    byId.set(event.id, event);
  }
  return [...byId.values()];
}

function extractMmcNames(workbook) {
  const names = new Map();
  for (const sheetName of workbook.SheetNames) {
    if (!sheetName.startsWith("Week ")) continue;
    const sheet = workbook.Sheets[sheetName];
    for (const { name } of iterateMmcRosterPeople(sheet)) {
      if (!looksLikePersonName(name)) continue;
      const key = normalizeName(name);
      if (!names.has(key)) names.set(key, String(name).trim());
    }
  }
  return names;
}

function extractDdhNames(workbook) {
  const names = new Map();
  for (const entry of iterateDdhWeekEntries(workbook)) {
    const key = normalizeName(entry.rawName);
    if (!names.has(key)) names.set(key, entry.rawName);
  }
  return names;
}

function extractCaseyNames(workbook) {
  const names = new Map();
  for (const entry of iterateCaseyWeekEntries(workbook)) {
    const key = normalizeName(entry.rawName);
    if (!names.has(key)) names.set(key, entry.displayName || entry.rawName);
  }
  return names;
}

function extractMchNames(workbook) {
  const names = new Map();
  for (const entry of iterateMchWeekEntries(workbook)) {
    const key = normalizeName(entry.rawName);
    if (!names.has(key)) names.set(key, entry.displayName || entry.rawName);
  }
  return names;
}

function parseMmcRecords(workbook, doctorKey) {
  const records = [];
  for (const sheetName of workbook.SheetNames) {
    if (!sheetName.startsWith("Week ")) continue;
    const sheet = workbook.Sheets[sheetName];
    const layout = mmcWeekLayout(sheet);
    if (!layout) continue;
    const { weekDates } = layout;

    let currentSeniority = "SMS";
    const roleMap = mmcSeniorityMap();
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
    for (let row = 1; row <= range.e.r + 1; row += 1) {
      const marker = layout.markerColumn ? cleanText(getCellValue(sheet, row, layout.markerColumn)).replace(/\s+/g, " ").trim().toUpperCase() : "";
      if (isMmcStopSection(marker)) break;
      if (roleMap.has(marker)) currentSeniority = roleMap.get(marker);
      const name = cleanMmcRosterName(getCellValue(sheet, row, layout.nameColumn));
      if (!looksLikePersonName(name)) continue;
      if (normalizeName(name) !== doctorKey) continue;
      const weekValues = layout.dayColumns.map((column) => cleanText(getCellValue(sheet, row, column)));
      const weeklyLeave = mondayWeeklyLeave(weekValues);
      if (weeklyLeave && !hasNonLeaveMmcEntry(weekValues)) {
        records.push(createWeeklyLeaveRecord("MMC", weekDates[0], weeklyLeave, currentSeniority));
      } else {
        weekValues.forEach((raw, index) => {
          if (!raw) return;
          const record = parseMmcEntry(weekDates[index], raw, currentSeniority);
          if (record) records.push(record);
        });
      }
      break;
    }
  }
  return records;
}

function parseDdhRecords(workbook, doctorKey) {
  const findmyshiftRecords = parseFindmyshiftDdhRecords(workbook, doctorKey);
  if (findmyshiftRecords) return findmyshiftRecords;
  const records = [];
  for (const entry of iterateDdhWeekEntries(workbook)) {
    if (normalizeName(entry.rawName) !== doctorKey) continue;
    const weeklyLeave = entry.findmyshiftFormat ? null : mondayWeeklyLeave(entry.labels);
    if (weeklyLeave && !hasNonLeaveDdhEntry(entry.labels)) {
      records.push(createWeeklyLeaveRecord("DDH", entry.weekDates[0], weeklyLeave, entry.seniority));
      continue;
    }
    entry.weekDates.forEach((day, index) => {
      const record = parseDdhEntry(day, entry.labels[index], entry.times[index] || "", entry.seniority);
      if (record) records.push(record);
    });
  }
  return records;
}

// The generated grid keeps FindMyShift compatible with the established DDH
// worksheet parser, while this structured audit sheet remains the source of
// truth for an automated import.  It prevents facility and comment information
// from being flattened away before events are built.
function parseFindmyshiftDdhRecords(workbook, doctorKey) {
  if (!isFindmyshiftDdhWorkbook(workbook)) return null;
  const sheet = workbook?.Sheets?.["FindMyShift details"];
  if (!sheet) return null;
  const headers = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false })?.[0] || [];
  const expected = ["Staff ID", "Staff name", "Seniority/job title", "Date", "Shift label", "Start", "End", "Facility", "Comment"];
  if (expected.some((header, index) => cleanText(headers[index]) !== header)) return null;
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false }).slice(1);
  const records = [];
  const seen = new Set();
  for (const values of rows) {
    const [, rawName, rawSeniority, rawDate, rawLabel, rawStart, rawEnd, rawFacility, rawComment] = values;
    if (normalizeName(rawName) !== doctorKey) continue;
    const day = parseFindmyshiftDate(rawDate);
    const label = cleanText(rawLabel);
    const startHm = parseFindmyshiftTime(rawStart);
    const endHm = parseFindmyshiftTime(rawEnd);
    const seniority = findmyshiftDdhSeniority(rawSeniority, label);
    if (!day || !label || isDdhStructuralAnnotation(label, seniority)) continue;
    const facility = cleanText(rawFacility);
    const comment = cleanText(rawComment);
    let record;
    if (startHm && endHm) {
      const resolved = label.toUpperCase() === "SHIFT" ? null : resolveDdhShiftLabel(label, seniority);
      if (resolved?.normalized?.includeAsShift === false) continue;
      const titleParts = label.toUpperCase() === "SHIFT"
        ? findmyshiftTimedShiftTitleParts(facility, startHm, seniority)
        : resolved?.normalized?.titleParts || { base: label, period: "", suffix: "" };
      record = createTimedRecord("DDH", day, label, {
        kind: resolved?.normalized?.kind || "shift",
        titleParts,
        startHm,
        endHm,
        location: facility || DDH_LOCATION,
        status: resolved?.normalized?.status,
        warning: resolved?.normalized?.warning,
        seniority,
      });
    } else {
      record = parseDdhEntry(day, label, "", seniority);
      if (!record) continue;
      record = { ...record, location: facility || record.location || DDH_LOCATION };
    }
    const key = [record.kind, record.start, record.end, record.normalizedTitle, record.location].join("|");
    if (seen.has(key)) continue;
    seen.add(key);
    records.push({ ...record, comment, findmyshift: true });
  }
  return records;
}

function findmyshiftDdhSeniority(rawSeniority, label) {
  const supplied = cleanText(rawSeniority) || UNKNOWN_SENIORITY;
  if (supplied.toUpperCase() !== UNKNOWN_SENIORITY.toUpperCase()) return supplied;
  const upper = cleanText(label).replace(/\s+/g, " ").trim().toUpperCase();
  if (/\bINTERN\b/.test(upper)) return "Intern";
  if (/\bHMO\b/.test(upper)) return "HMO";
  if (/\bSMS\b/.test(upper)) return "SMS";
  if (/\bCMO\b/.test(upper)) return "CMO";
  if (/\b(?:SENIOR REGISTRAR|SENIOR REG|SR)\b/.test(upper)) return "Senior Registrar";
  if (/\b(?:TRANSITIONAL|INTERMEDIATE|IR|TR)\b/.test(upper)) return "Transitional/Intermediate Registrar";
  if (/\b(?:JUNIOR REGISTRAR|JUNIOR REG|JR)\b/.test(upper)) return "Junior Registrar";
  return supplied;
}

function findmyshiftTimedShiftTitleParts(facility, startHm, seniority) {
  const value = cleanText(facility);
  if (!value) return genericTimeOnlyShiftTitleParts(startHm, "DDH");
  const normalized = resolveDdhShiftLabel(value, seniority)?.normalized;
  if (normalized?.titleParts?.base) {
    return normalized.titleParts;
  }
  return { base: value, period: inferGenericTimeOnlyShiftPeriod(startHm, "DDH"), suffix: "" };
}

function parseFindmyshiftTime(value) {
  const match = cleanText(value).match(/^(\d{1,2}):(\d{2})$/);
  if (!match || Number(match[1]) > 23 || Number(match[2]) > 59) return null;
  return [Number(match[1]), Number(match[2])];
}

function parseFindmyshiftDate(value) {
  const date = cleanText(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(date)) return date;
  return coerceDate(value);
}

function parseCaseyRecords(workbook, doctorKey) {
  const records = [];
  for (const entry of iterateCaseyWeekEntries(workbook)) {
    if (normalizeName(entry.rawName) !== doctorKey) continue;
    entry.weekDates.forEach((day, index) => {
      const raw = entry.labels[index];
      if (!raw) return;
      const record = parseCaseyEntry(day, raw, entry.seniority);
      if (record) records.push(record);
    });
  }
  return records;
}

function parseMchRecords(workbook, doctorKey) {
  const records = [];
  for (const entry of iterateMchWeekEntries(workbook)) {
    if (normalizeName(entry.rawName) !== doctorKey) continue;
    entry.weekDates.forEach((day, index) => {
      const raw = entry.labels[index];
      if (!raw) return;
      const record = parseMchEntry(day, raw, entry.seniority);
      if (record) records.push(record);
    });
  }
  return records;
}

function parseMmcEntry(day, raw, seniority = UNKNOWN_SENIORITY) {
  const upper = raw.toUpperCase();
  // This is a clinical-support allocation, not exam leave.  It must be
  // checked before the generic leave matcher, which deliberately accepts
  // a trailing "Exam" in other roster annotations.
  const leaveRecord = isMmcClinicalSupportExam(raw) ? null : createRecognizedLeaveRecord("MMC", day, raw, seniority);
  if (leaveRecord) return leaveRecord;
  const orientation = createOrientationRecord("MMC", day, raw, seniority);
  if (orientation) return orientation;
  if (isOtherHospitalReference("MMC", raw)) return null;
  if (shouldIgnoreMmc(raw)) return null;
  if (upper === "PHNW") {
    return createAllDayRecord("MMC", day, raw, {
      kind: "public_holiday",
      titleParts: { base: "PHNW", period: "", suffix: "" },
      location: "",
      seniority,
    });
  }

  const explicit = extractMmcExplicitTime(raw);
  const label = explicit ? explicit.label : raw.trim();
  const normalized = findManualParserRule("MMC", seniority, label, explicit) || normalizeGenericMmcTimedLabel(label, explicit);
  if (!normalized) {
    return createUnknownRecord("MMC", day, raw, "MMC shift code not recognised.", seniority);
  }
  if (normalized.includeAsShift === false) {
    return createHiddenRecord("MMC", day, raw, normalized, seniority);
  }

  // Explicit roster-column hours describe the allocation column, but exam
  // events are intentionally represented as all-day calendar entries.
  if (normalized.allDay) {
    return createAllDayRecord("MMC", day, raw, {
      kind: normalized.kind,
      titleParts: normalized.titleParts,
      location: normalized.location || "",
      seniority,
    });
  }

  if (explicit) {
    return createTimedRecord("MMC", day, raw, {
      kind: normalized.kind,
      titleParts: normalized.titleParts,
      startHm: explicit.start,
      endHm: explicit.end,
      location: normalized.location || "",
      ambiguous: normalized.ambiguous,
      status: normalized.status,
      warning: normalized.warning,
      seniority,
    });
  }
  return createTimedRecord("MMC", day, raw, {
    kind: normalized.kind,
    titleParts: normalized.titleParts,
    startHm: normalized.defaultTimes[0],
    endHm: normalized.defaultTimes[1],
    location: normalized.location || "",
    seniority,
  });
}

function parseDdhEntry(day, label, timeText, seniority = UNKNOWN_SENIORITY) {
  if (!label) return null;
  if (isDdhStructuralAnnotation(label, seniority)) return null;
  const labelTime = parseDdhTimeRow(label);
  if (labelTime) {
    if (timeText) return parseDdhEntry(day, timeText, label, seniority);
    return createTimedRecord("DDH", day, label, {
      kind: "shift",
      titleParts: genericTimeOnlyShiftTitleParts(labelTime[0], "DDH"),
      startHm: labelTime[0],
      endHm: labelTime[1],
      location: DDH_LOCATION,
      seniority,
    });
  }
  if (parseDdhTimeRow(timeText) && (isOtherHospitalReference("DDH", label) || shouldIgnoreDdh(label, seniority))) {
    return null;
  }
  const upper = label.toUpperCase();
  const leaveRecord = createRecognizedLeaveRecord("DDH", day, label, seniority);
  if (leaveRecord) return leaveRecord;
  const orientation = createOrientationRecord("DDH", day, label, seniority);
  if (orientation) return orientation;
  if (isOtherHospitalReference("DDH", label)) return null;
  const extraRecord = createDdhExtraRecord(day, label, timeText, seniority);
  if (extraRecord !== undefined) return extraRecord;
  if (upper === "AM" || upper === "PM") return null;
  if (upper === "PHNW" || upper === "PHNW CLINICAL") {
    return createAllDayRecord("DDH", day, label, {
      kind: "public_holiday",
      titleParts: { base: "PHNW", period: "", suffix: "" },
      location: "",
      seniority,
    });
  }
  if (!isDdhClinicalSupportExam(label) && (shouldIgnoreDdh(label, seniority) || shouldIgnoreCommon(label))) return null;

  if (/\bCRISIS\s+LOCUM\b/i.test(label)) {
    const parsedTime = parseDdhTimeRow(timeText);
    if (parsedTime) {
      return createTimedRecord("DDH", day, label, {
        kind: "shift",
        titleParts: { base: "Crisis Locum", period: "", suffix: "" },
        startHm: parsedTime[0],
        endHm: parsedTime[1],
        location: DDH_LOCATION,
        seniority,
      });
    }
    return createAllDayRecord("DDH", day, label, {
      kind: "shift",
      titleParts: { base: "Crisis Locum", period: "", suffix: "" },
      location: DDH_LOCATION,
      seniority,
    });
  }

  const resolved = resolveDdhShiftLabel(label, seniority);
  const { mapped, normalized } = resolved;
  if (!normalized) {
    return createUnknownRecord("DDH", day, label, "DDH shift label not recognised.", seniority);
  }
  if (normalized.includeAsShift === false) {
    return createHiddenRecord("DDH", day, label, normalized, seniority);
  }
  const location = normalizeDdhLocation(mapped, normalized);

  const parsedTime = parseDdhTimeRow(timeText);
  if (parsedTime) {
    const titleParts = ddhSwingTitlePartsForStart(label, parsedTime[0], normalized.titleParts);
    return createTimedRecord("DDH", day, label, {
      kind: normalized.kind,
      titleParts,
      startHm: parsedTime[0],
      endHm: parsedTime[1],
      location,
      status: normalized.status,
      warning: normalized.warning,
      seniority,
    });
  }

  if (normalized.allDay !== true) {
    const defaultTimes = normalized.defaultTimes || inferDdhDefaultTimes(normalized, seniority);
    if (defaultTimes) {
      return createTimedRecord("DDH", day, label, {
        kind: normalized.kind,
        titleParts: normalized.titleParts,
        startHm: defaultTimes[0],
        endHm: defaultTimes[1],
        location,
        status: normalized.status,
        warning: normalized.warning,
        seniority,
      });
    }
  }

  return createAllDayRecord("DDH", day, label, {
    kind: normalized.kind,
    titleParts: normalized.titleParts,
    location,
    status: normalized.status,
    warning: normalized.warning,
    seniority,
  });
}

// DDH "Extra" annotations usually reconcile payment for a shift worked in a
// different week. They are not a shift on the displayed roster day unless the
// writer supplies a period or an explicit time. Extra Swing remains a genuine
// Swing allocation and is handled by the normal Swing rules below.
function createDdhExtraRecord(day, label, timeText, seniority) {
  const upper = String(label || "").replace(/\s+/g, " ").trim().toUpperCase();
  if (!/^EXTRA\b/.test(upper) || /^EXTRA\s+SWING$/.test(upper)) return undefined;
  const inlineTime = String(label || "").match(/\b(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})\b/);
  const parsedTime = parseDdhTimeRow(timeText) || (inlineTime
    ? [[Number(inlineTime[1]), Number(inlineTime[2])], [Number(inlineTime[3]), Number(inlineTime[4])]]
    : null);
  const explicitPeriod = extractDdhPeriod(upper);
  if (!parsedTime && !explicitPeriod) return null;
  const period = explicitPeriod || (Number(parsedTime[0][0]) < 14 ? "AM" : "PM");
  const titleParts = { base: "Extra", period, suffix: "" };
  if (parsedTime) {
    return createTimedRecord("DDH", day, label, {
      kind: "shift",
      titleParts,
      startHm: parsedTime[0],
      endHm: parsedTime[1],
      location: DDH_LOCATION,
      seniority,
    });
  }
  const defaultTimes = inferDdhDefaultTimes({ titleParts }, seniority);
  return defaultTimes
    ? createTimedRecord("DDH", day, label, { kind: "shift", titleParts, startHm: defaultTimes[0], endHm: defaultTimes[1], location: DDH_LOCATION, seniority })
    : createAllDayRecord("DDH", day, label, { kind: "shift", titleParts, location: DDH_LOCATION, seniority });
}

function inferDdhDefaultTimes(normalized, seniority = UNKNOWN_SENIORITY) {
  const base = String(normalized?.titleParts?.base || "").trim().toUpperCase();
  const explicitPeriod = String(normalized?.titleParts?.period || "").trim().toUpperCase();
  const period = explicitPeriod || (base === "SSU" ? "AM" : "");
  if (base === "NIGHT" || base === "NIGHT SSU") return [[23, 0], [8, 30]];
  if (period === "AM") {
    if (base === "SSU") return [[7, 30], [17, 30]];
    if (seniority === "SMS" && base === "AVAO") return [[7, 30], [17, 0]];
    return [[8, 0], [18, 0]];
  }
  if (period === "PM") {
    if (seniority === "SMS" && base !== "AVAO") return [[15, 0], [0, 0]];
    return [[14, 30], [0, 0]];
  }
  return null;
}

function parseCaseyEntry(day, raw, seniority = UNKNOWN_SENIORITY) {
  const label = cleanText(raw);
  if (!label) return null;
  const shiftLabel = stripCaseyOrientationMetadata(label);
  if (!shiftLabel) return null;
  const upper = label.toUpperCase();
  const leaveRecord = createRecognizedLeaveRecord("Casey", day, label, seniority);
  if (leaveRecord) return leaveRecord;
  const orientation = createOrientationRecord("Casey", day, label, seniority);
  if (orientation) return orientation;
  if (isOtherHospitalReference("Casey", shiftLabel)) return null;
  if (upper === "PHNW") {
    return createAllDayRecord("Casey", day, label, {
      kind: "public_holiday",
      titleParts: { base: "PHNW", period: "", suffix: "" },
      location: "",
      seniority,
    });
  }
  if (shouldIgnoreCasey(shiftLabel) || shouldIgnoreCommon(shiftLabel)) return null;

  const explicit = extractTimePrefix(shiftLabel);
  const normalized = findManualParserRule("Casey", seniority, shiftLabel, explicit) || normalizeCaseyLabel(shiftLabel, explicit);
  if (!normalized) {
    return createUnknownRecord("Casey", day, label, "Casey shift label not recognised.", seniority);
  }
  if (normalized.includeAsShift === false) {
    return createHiddenRecord("Casey", day, label, normalized, seniority);
  }
  if (normalized.allDay) {
    return createAllDayRecord("Casey", day, label, {
      kind: normalized.kind || "shift",
      titleParts: normalized.titleParts,
      location: normalized.location || CASEY_LOCATION,
      seniority,
    });
  }
  if (explicit) {
    if (!Array.isArray(explicit.start) || !Array.isArray(explicit.end)) {
      return createUnknownRecord("Casey", day, label, "Casey shift time could not be parsed.", seniority);
    }
    return createTimedRecord("Casey", day, label, {
      kind: normalized.kind,
      titleParts: normalized.titleParts,
      startHm: explicit.start,
      endHm: explicit.end,
      location: normalized.location || CASEY_LOCATION,
      status: normalized.status,
      warning: normalized.warning,
      seniority,
    });
  }
  if (!Array.isArray(normalized.defaultTimes) || normalized.defaultTimes.length < 2) {
    return createUnknownRecord("Casey", day, label, "Casey shift times are missing.", seniority);
  }
  return createTimedRecord("Casey", day, label, {
    kind: normalized.kind,
    titleParts: normalized.titleParts,
    startHm: normalized.defaultTimes[0],
    endHm: normalized.defaultTimes[1],
    location: normalized.location || CASEY_LOCATION,
    seniority,
  });
}

function parseMchEntry(day, raw, seniority = UNKNOWN_SENIORITY) {
  const label = cleanText(raw).replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
  if (!label) return null;
  const upper = label.toUpperCase();
  const leave = normalizeMchLeave(label);
  if (leave) {
    return createAllDayRecord("MCH", day, label, {
      kind: leave.kind,
      titleParts: { base: leave.title, period: "", suffix: "" },
      location: "",
      seniority,
    });
  }
  const orientation = createOrientationRecord("MCH", day, label, seniority);
  if (orientation) return orientation;
  if (isOtherHospitalReference("MCH", label)) return null;
  if (shouldIgnoreMch(label) || shouldIgnoreCommon(label)) return null;

  if (upper.includes("EDO")) return null;

  const explicit = extractTimeWithLabel(label);
  if (/^NIGHT$/i.test(label)) {
    return createTimedRecord("MCH", day, label, {
      kind: "shift",
      titleParts: { base: "Night shift", period: "", suffix: "" },
      startHm: [23, 0],
      endHm: [8, 30],
      location: MCH_LOCATION,
      seniority,
    });
  }
  if (/\bSIM\s+DAY\b/i.test(label) && explicit) {
    return createTimedRecord("MCH", day, label, {
      kind: "shift",
      titleParts: { base: "SIM day", period: "", suffix: "" },
      startHm: explicit.start,
      endHm: explicit.end,
      location: MCH_LOCATION,
      seniority,
    });
  }
  const manual = findManualParserRule("MCH", seniority, label, explicit);
  if (manual) {
    if (manual.includeAsShift === false) {
      return createHiddenRecord("MCH", day, label, manual, seniority);
    }
    if (explicit) {
      return createTimedRecord("MCH", day, label, {
        kind: manual.kind,
        titleParts: manual.titleParts,
        startHm: explicit.start,
        endHm: explicit.end,
        location: manual.location || MCH_LOCATION,
        seniority,
      });
    }
    if (manual.allDay) {
      return createAllDayRecord("MCH", day, label, {
        kind: manual.kind,
        titleParts: manual.titleParts,
        location: manual.location || "",
        seniority,
      });
    }
    return createTimedRecord("MCH", day, label, {
      kind: manual.kind,
      titleParts: manual.titleParts,
      startHm: manual.defaultTimes[0],
      endHm: manual.defaultTimes[1],
      location: manual.location || MCH_LOCATION,
      seniority,
    });
  }

  const timed = normalizeMchTimedLabel(label);
  if (!timed) {
    return createUnknownRecord("MCH", day, label, "MCH shift label not recognised.", seniority);
  }
  return createTimedRecord("MCH", day, label, {
    kind: "shift",
    titleParts: timed.titleParts,
    startHm: timed.start,
    endHm: timed.end,
    location: MCH_LOCATION,
    seniority,
  });
}

function normalizeMmcLabel(label) {
  const code = label.trim().toUpperCase();
  if (code === "CS") {
    return {
      kind: "shift",
      titleParts: { base: "CS", period: "", suffix: "" },
      location: "",
      allDay: true,
      defaultTimes: null,
    };
  }
  if (code === "CSO") {
    return {
      kind: "shift",
      titleParts: { base: "CSO", period: "", suffix: "" },
      location: MMC_LOCATION,
      allDay: true,
      defaultTimes: null,
    };
  }

  if (code.length === 4 && code[1] === "S" && code[2] === "S" && ["A", "P"].includes(code[0]) && ["C", "R"].includes(code[3])) {
    return {
      kind: "shift",
      titleParts: {
        base: "SSU",
        period: code[0] === "A" ? "AM" : "PM",
        suffix: code[3] === "R" ? "Float" : "",
      },
      location: MMC_LOCATION,
      allDay: false,
      defaultTimes: code[0] === "A" ? [[7, 30], [17, 30]] : [[14, 30], [0, 0]],
    };
  }

  if (code.length === 3 && ["A", "P"].includes(code[0]) && MMC_TEAM_MAP[code[1]] && ["C", "R"].includes(code[2])) {
    return {
      kind: "shift",
      titleParts: {
        base: MMC_TEAM_MAP[code[1]],
        period: code[0] === "A" ? "AM" : "PM",
        suffix: code[2] === "R" ? "Float" : "",
      },
      location: MMC_LOCATION,
      allDay: false,
      defaultTimes: code[0] === "A" ? [[8, 0], [17, 30]] : [[14, 30], [0, 0]],
    };
  }

  return null;
}

function normalizeGenericMmcTimedLabel(label, explicit) {
  if (!explicit) return null;
  const code = label.trim().toUpperCase();
  const titleParts = code
    ? { base: code, period: "", suffix: "" }
    : genericTimeOnlyShiftTitleParts(explicit.start, "MMC");
  const base = titleParts.base;
  if (!base) return null;
  return {
    kind: "shift",
    titleParts,
    location: MMC_LOCATION,
    allDay: false,
    defaultTimes: null,
    status: code ? "unknown" : "ok",
    warning: code ? "MMC shift code not recognised; using explicit roster time." : "",
  };
}

function extractMmcExplicitTime(value) {
  return extractTimeWithLabel(value);
}

function normalizeCaseyLabel(label, explicit = null) {
  const code = normalizeCaseyCode(explicit ? explicit.label : label);
  if (!code && explicit?.start) {
    return {
      kind: "shift",
      titleParts: genericTimeOnlyShiftTitleParts(explicit.start, "Casey"),
      location: CASEY_LOCATION,
      allDay: false,
      defaultTimes: null,
    };
  }
  if (!code) return null;
  const direct = caseyTimedRuleForCode(code);
  if (direct) {
    return {
      kind: "shift",
      titleParts: direct.titleParts,
      location: CASEY_LOCATION,
      allDay: false,
      defaultTimes: direct.defaultTimes,
    };
  }
  if (explicit) {
    return {
      kind: "shift",
      titleParts: caseyTitlePartsForCode(code, explicit.start),
      location: CASEY_LOCATION,
      allDay: false,
      defaultTimes: null,
    };
  }
  return null;
}

function caseyTimedRuleForCode(code) {
  const rules = {
    "CS": [["CS", "", ""], [[8, 0], [17, 30]]],
    "CLIN SUPP": [["CS", "", ""], [[8, 0], [17, 30]]],
    "CLINICAL SUPP": [["CS", "", ""], [[8, 0], [17, 30]]],
    "AM TL": [["TL", "AM", ""], [[8, 0], [17, 30]]],
    "AM T/L": [["TL", "AM", ""], [[8, 0], [17, 30]]],
    "AM UFD": [["UFD", "AM", ""], [[8, 0], [17, 30]]],
    "AM MIC": [["MIC", "AM", ""], [[8, 0], [17, 30]]],
    "PM TL": [["TL", "PM", ""], [[14, 30], [0, 0]]],
    "PM T/L": [["TL", "PM", ""], [[14, 30], [0, 0]]],
    "PM UFD": [["UFD", "PM", ""], [[14, 30], [0, 0]]],
    "PM MIC": [["MIC", "PM", ""], [[14, 30], [0, 0]]],
    "PM PAEDS": [["PAEDS", "PM", ""], [[14, 30], [0, 0]]],
    "AM PAEDS": [["PAEDS", "AM", ""], [[7, 30], [17, 0]]],
    "AM SSU": [["SSU", "AM", ""], [[7, 30], [17, 0]]],
    "PM": [["PM", "", ""], [[14, 30], [0, 0]]],
    "AM": [["AM", "", ""], [[8, 0], [17, 30]]],
  };
  const item = rules[code];
  if (!item) return null;
  return {
    titleParts: { base: item[0][0], period: item[0][1], suffix: item[0][2] },
    defaultTimes: item[1],
  };
}

function caseyTitlePartsForCode(code, startHm) {
  const period = extractCaseyPeriod(code) || caseyPeriodFromStart(startHm);
  const base = normalizeCaseyBase(code.replace(/\b(?:AM|PM)\b/g, " ").trim()) || inferCaseyTimeOnlyShiftLabel(startHm);
  if (base === "AM" || base === "PM" || base === "Night" || base === "Evening") {
    return { base, period: "", suffix: "" };
  }
  return { base, period, suffix: "" };
}

function normalizeCaseyCode(label) {
  return cleanText(label)
    .replace(/\([^)]*\)/g, " ")
    .replace(/\+\s*OC\b/gi, "")
    .replace(/\*/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function normalizeMchLeave(label) {
  const upper = normalizedLeaveLabel(label);
  if (upper === "ANNUAL & PARENTAL LEAVE") return { kind: "annual_parental_leave", title: "Annual & Parental Leave" };
  if (isAnnualLeaveLabel(upper)) return { kind: "annual_leave", title: "Annual Leave" };
  if (/^(?:SICK(?:\s+LEAVE)?|S\/L)(?:\s+.*)?$/.test(upper)) return { kind: "sick_leave", title: "Sick leave" };
  if (upper === "EXAM" || upper === "EXAM LEAVE" || upper === "ME/L" || upper === "EL") return { kind: "exam_leave", title: "Exam Leave" };
  if (upper === "EXAM/CONF LEAVE") return { kind: "exam_leave", title: "Exam / Conference Leave" };
  if (isConferenceLeaveLabel(upper)) return { kind: "conference_leave", title: "Conference Leave" };
  // DDH/MMC roster writers use SL (often with a second-site suffix such as
  // "SL MMC") for sabbatical leave. Keep this deliberately separate from
  // S/L, which is sick leave throughout the supported rosters.
  if (/^(?:SABBATICAL(?:\s+LEAVE)?|SAB\/L)(?:\s+.*)?$/.test(upper)
    || /^SL(?:\s+(?:MMC|DDH|CASEY|MCH|PAEDS|AM|PM|NIGHT|NS|SW))?$/.test(upper)) {
    return { kind: "sabbatical_leave", title: "Sabbatical" };
  }
  if (upper === "PAT/L" || upper === "PARENTAL/L" || upper === "PARENTAL LEAVE" || upper === "PATERNITY LEAVE" || /\bPARENTAL\b/.test(upper)) return { kind: "parental_leave", title: "Parental Leave" };
  if (/\b(?:LWP|LWOP)\b/.test(upper) || /\bLEAVE\s+WITHOUT\s+PAY\b/.test(upper)) return { kind: "leave", title: "Leave without pay" };
  if (upper === "LSL" || upper === "LONG SERVICE LEAVE") return { kind: "long_service_leave", title: "Long Service Leave" };
  if (upper === "C/L" || /^CARER'?S LEAVE$/.test(upper) || upper === "CARERS LEAVE") return { kind: "carers_leave", title: "Carer's Leave" };
  if (/^F\/L(?:\s+(?:AM|PM))?$/.test(upper) || upper === "FAM LEAVE" || upper === "FAMILY LEAVE") return { kind: "family_leave", title: "Family Leave" };
  if (upper === "SPECIAL LEAVE") return { kind: "special_leave", title: "Special Leave" };
  if (upper === "OTHER - MILITARY LEAVE" || upper === "MILITARY LEAVE") return { kind: "military_leave", title: "Military Leave" };
  if (upper === "LEAVE") return { kind: "leave", title: "Leave" };
  return null;
}

function createRecognizedLeaveRecord(source, day, rawValue, seniority = UNKNOWN_SENIORITY) {
  const leave = normalizeRecognizedLeave(rawValue);
  if (!leave) return null;
  return createAllDayRecord(source, day, rawValue, {
    kind: leave.kind,
    titleParts: { base: leave.title, period: "", suffix: "" },
    location: "",
    seniority,
  });
}

function normalizeRecognizedLeave(value) {
  for (const candidate of leaveLabelCandidates(value)) {
    const leave = normalizeMchLeave(candidate);
    if (leave) return leave;
    if (candidate === "ANNUAL") return { kind: "annual_leave", title: "Annual Leave" };
  }
  return null;
}

function leaveLabelCandidates(value) {
  const words = normalizedLeaveLabel(value).split(" ").filter(Boolean);
  return words.map((_, index) => words.slice(index).join(" "));
}

function normalizedLeaveLabel(value) {
  return cleanText(value)
    .replace(/S\s*[\\/.]+\s*L/gi, "S/L")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function isConferenceLeaveLabel(value) {
  const upper = String(value || "").replace(/\s+/g, " ").trim().toUpperCase();
  return /^(?:CONFERENCE(?:\s+LEAVE)?|CONF(?:\s+LEAVE)?|CME(?:\s+LEAVE)?)(?:\s+(?:-|X)?\s*\d+(?:\.\d+)?\s*(?:HRS?|HOURS?|SHIFTS?))?$/.test(upper)
    || /^(?:C\/L|CL|CME\/L)(?:\s+\d+(?:\.\d+)?\s*(?:HRS?|HOURS?))?$/.test(upper);
}

function isAnnualLeaveLabel(value) {
  const upper = cleanText(value).replace(/\s+/g, " ").trim().toUpperCase();
  if (/^ANNUAL(?:\s+LEAVE)?(?:\s+(?:-|X)?\s*\d+(?:\.\d+)?\s*(?:HRS?|HOURS?|SHIFTS?))?$/.test(upper)) return true;
  return /^(?:A\/L|AL)(?:\s+(?:\d+(?:\.\d+)?(?:\s*(?:HRS?|HOURS?))?|MMC|DDH|CASEY|MCH|PAEDS))?$/.test(upper);
}

function normalizeMchTimedLabel(label) {
  const text = cleanText(label)
    .replace(/\u00a0/g, " ")
    .replace(/(\d{4})\s*[-–]\s*(\d{4})/g, "$1-$2")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
  const match = text.match(/^(?:(PHNW)\s+)?(\d{4})-(\d{4})(?:\s*(.*))?$/);
  if (!match) return null;
  const start = [Number(match[2].slice(0, 2)), Number(match[2].slice(2))];
  const end = [Number(match[3].slice(0, 2)), Number(match[3].slice(2))];
  const suffix = normalizeMchShiftSuffix(match[1] || match[4] || "");
  if (suffix === "EDO") return null;
  const titleParts = suffix
    ? { base: suffix, period: "", suffix: "" }
    : genericTimeOnlyShiftTitleParts(start, "MCH");
  return { start, end, titleParts };
}

function normalizeMchShiftSuffix(value) {
  const suffix = String(value || "").replace(/[^A-Z0-9]/gi, "").toUpperCase();
  if (!suffix) return "";
  if (suffix === "OCS" || suffix === "0CS" || suffix === "CSOS") return "CS Office";
  if (suffix === "CS" || suffix === "DEMT" || suffix === "PHNW") return suffix;
  if (suffix === "EDO") return "EDO";
  return suffix;
}

function inferMchTimeOnlyShiftLabel(startHm) {
  const [hour] = startHm;
  if (hour >= 22 || hour < 6) return "Night";
  if (hour >= 12) return "PM";
  return "AM";
}

function genericTimeOnlyShiftTitleParts(startHm, source = "") {
  const period = inferGenericTimeOnlyShiftPeriod(startHm, source);
  return { base: `${period} shift`, period: "", suffix: "" };
}

function inferGenericTimeOnlyShiftPeriod(startHm, source = "") {
  if (source === "MCH") return inferMchTimeOnlyShiftLabel(startHm);
  return inferMmcTimeOnlyShiftLabel(startHm);
}

function normalizeCaseyBase(value) {
  const code = normalizeCaseyCode(value);
  if (!code) return "";
  if (code === "CLIN SUPP" || code === "CLINICAL SUPP") return "CS";
  if (code === "T/L") return "TL";
  if (code === "C/CARE") return "C/CARE";
  if (code.includes("PAEDS")) return "PAEDS";
  if (code.includes("SSU")) return "SSU";
  if (code.includes("MIC")) return "MIC";
  if (code.includes("UFD")) return "UFD";
  if (code.includes("TL") || code.includes("T/L")) return "TL";
  return code;
}

function extractCaseyPeriod(label) {
  const upper = String(label || "").toUpperCase();
  if (/\bAM\b/.test(upper)) return "AM";
  if (/\bPM\b/.test(upper)) return "PM";
  return "";
}

function caseyPeriodFromStart(startHm) {
  const [hour] = startHm;
  if (hour >= 14 && hour < 22) return "PM";
  if (hour >= 6 && hour < 14) return "AM";
  return "";
}

function inferCaseyTimeOnlyShiftLabel(startHm) {
  const [hour] = startHm;
  if (hour >= 22 || hour < 6) return "Night";
  if (hour >= 16) return "Evening";
  if (hour >= 12) return "PM";
  return "AM";
}

function sanitizeParserExtensions(value) {
  const input = value && typeof value === "object" ? value : {};
  const defaults = buildDefaultParserRules();
  const removed = sanitizeParserRuleRemovals(input._removed);
  return {
    mmc: applyParserRuleRemovals(mergeParserRuleLists(defaults.mmc, sanitizeParserExtensionRuleList(input.mmc, "MMC")), removed),
    ddh: applyParserRuleRemovals(mergeParserRuleLists(defaults.ddh, sanitizeParserExtensionRuleList(input.ddh, "DDH")), removed),
    casey: applyParserRuleRemovals(mergeParserRuleLists(defaults.casey, sanitizeParserExtensionRuleList(input.casey, "Casey")), removed),
    mch: applyParserRuleRemovals(mergeParserRuleLists(defaults.mch, sanitizeParserExtensionRuleList(input.mch, "MCH")), removed),
    _removed: removed,
  };
}

function sanitizeParserRuleRemovals(items) {
  if (!Array.isArray(items)) return [];
  const byKey = new Map();
  for (const item of items) {
    const source = sanitizeParserRuleSource(item?.source);
    const seniority = sanitizeRuleSeniority(item?.seniority);
    const code = String(item?.code || "").trim().toUpperCase();
    if (!source || !code) continue;
    byKey.set(`${source}|${seniority}|${code}`, { source, seniority, code });
  }
  return [...byKey.values()].sort(compareParserRules);
}

function applyParserRuleRemovals(rules, removals) {
  const removedKeys = new Set((removals || []).map(parserRuleKey));
  return (rules || []).filter((rule) => !removedKeys.has(parserRuleKey(rule)));
}

function mergeParserRuleLists(defaults, overrides) {
  const byKey = new Map();
  for (const rule of defaults) byKey.set(parserRuleKey(rule), rule);
  for (const rule of overrides) byKey.set(parserRuleKey(rule), rule);
  return [...byKey.values()].sort(compareParserRules);
}

function parserRuleKey(rule) {
  return `${rule.source}|${rule.seniority}|${rule.code}`;
}

function compareParserRules(left, right) {
  const rank = seniorityRank(left.seniority) - seniorityRank(right.seniority);
  if (rank) return rank;
  return left.code.localeCompare(right.code);
}

function sanitizeParserExtensionRuleList(items, source) {
  if (!Array.isArray(items)) return [];
  return items
    .map((item) => sanitizeParserExtensionRule(item, source))
    .filter(Boolean);
}

function sanitizeParserExtensionRule(item, forcedSource = "") {
  if (!item || typeof item !== "object") return null;
  const source = sanitizeParserRuleSource(forcedSource || item.source);
  const seniority = sanitizeRuleSeniority(item.seniority);
  const code = normalizeParserExtensionRuleCode(source, item.code || item.rawCode || "");
  const ignore = item.ignore === true || String(item.kind || "").trim().toLowerCase() === "ignore";
  const kind = ignore ? "ignore" : String(item.kind || "shift").trim().toLowerCase();
  const base = String(item.base || item.titleParts?.base || "").trim();
  const period = String(item.period || item.titleParts?.period || "").trim().toUpperCase();
  const suffix = String(item.suffix || item.titleParts?.suffix || "").trim();
  const location = String(item.location || "").trim();
  const allDay = item.allDay === true;
  const startTime = String(item.startTime || "").trim();
  const endTime = String(item.endTime || "").trim();
  if (!source || !code || (!ignore && !base)) return null;
  if (isRestrictedClinicalSupportRule({ seniority, code, base })) return null;
  if (!ignore && !allDay && (!isClockString(startTime) || !isClockString(endTime))) return null;
  return {
    source,
    seniority,
    code,
    kind,
    base,
    period,
    suffix,
    location,
    allDay: ignore ? true : allDay,
    startTime: ignore || allDay ? "" : startTime,
    endTime: ignore || allDay ? "" : endTime,
    includeAsShift: ignore ? false : item.includeAsShift !== false,
    ignore,
  };
}

function isRestrictedClinicalSupportRule(rule) {
  const seniority = sanitizeRuleSeniority(rule?.seniority);
  if (seniority === "SMS" || seniority === "CMO") return false;
  const code = String(rule?.code || "").trim().toUpperCase();
  const base = String(rule?.base || "").trim().toUpperCase();
  return code === "CS"
    || code === "CSO"
    || code === "CS ONSITE"
    || code === "CLIN SUPP"
    || code === "CLINICAL SUPP"
    || base === "CS"
    || base === "CSO"
    || base === "CS ONSITE";
}

function normalizeParserExtensionRuleCode(source, value) {
  const text = String(value || "").trim().toUpperCase();
  if (!text) return "";
  if (source === "MMC" || source === "Casey" || source === "MCH") {
    const explicit = extractTimeWithLabel(text);
    return (explicit?.label || text).trim().toUpperCase();
  }
  return text;
}

function sanitizeRuleSeniority(value) {
  const text = String(value || "").trim();
  const upper = text.toUpperCase();
  const aliases = new Map([
    ["SR", "Senior Registrar"],
    ["SENIOR REG", "Senior Registrar"],
    ["SENIOR REGISTRAR", "Senior Registrar"],
    ["IR", "Transitional/Intermediate Registrar"],
    ["TR", "Transitional/Intermediate Registrar"],
    ["INTERMEDIATE REG", "Transitional/Intermediate Registrar"],
    ["INTERMEDIATE REGISTRAR", "Transitional/Intermediate Registrar"],
    ["TRANSITIONAL REGISTRAR", "Transitional/Intermediate Registrar"],
    ["JR", "Junior Registrar"],
    ["JUNIOR REG", "Junior Registrar"],
    ["JUNIOR REGISTRAR", "Junior Registrar"],
    ["H", "HMO"],
    ["HMO", "HMO"],
    ["ED HMO", "HMO"],
    ["PAEDS HMO", "HMO"],
    ["GERI", "SMS"],
    ["I", "Intern"],
    ["INTERN", "Intern"],
  ]);
  if (aliases.has(upper)) return aliases.get(upper);
  return SENIORITY_LABELS.find((item) => item.toUpperCase() === upper) || UNKNOWN_SENIORITY;
}

function seniorityRank(value) {
  const index = SENIORITY_LABELS.indexOf(sanitizeRuleSeniority(value));
  return index >= 0 ? index : SENIORITY_LABELS.length;
}

function sanitizeParserRuleSource(value) {
  const source = String(value || "").trim().toUpperCase();
  if (source === "MMC" || source === "DDH") return source;
  if (source === "CASEY") return "Casey";
  if (source === "MCH") return "MCH";
  return "";
}

function findManualParserRule(source, seniority, label, explicit = null) {
  const normalizedSource = sanitizeParserRuleSource(source);
  const normalizedSeniority = sanitizeRuleSeniority(seniority);
  const code = normalizeParserRuleCode(normalizedSource, label);
  if (!normalizedSource || !code) return null;
  const rules = parserRulesForSource(normalizedSource);
  const rule = rules.find((item) => item.code === code && item.seniority === normalizedSeniority)
    || rules.find((item) => item.code === code && item.seniority === UNKNOWN_SENIORITY);
  return rule ? parserRuleToNormalized(rule, explicit) : null;
}

function normalizeParserRuleCode(source, label) {
  const text = String(label || "").trim().toUpperCase();
  if (!text) return "";
  if (source === "MMC") {
    const explicit = extractTimePrefix(text);
    return String(explicit?.label || text).trim().toUpperCase();
  }
  if (source === "Casey") {
    const explicit = extractTimePrefix(text);
    return normalizeCaseyCode(explicit?.label || text);
  }
  if (source === "MCH") {
    const explicit = extractTimeWithLabel(text);
    return (explicit?.label || text).replace(/\s+/g, " ").trim();
  }
  return text;
}

function parserRulesForSource(source) {
  if (source === "MMC") return MANUAL_PARSER_RULES.mmc || [];
  if (source === "DDH") return MANUAL_PARSER_RULES.ddh || [];
  if (source === "Casey") return MANUAL_PARSER_RULES.casey || [];
  if (source === "MCH") return MANUAL_PARSER_RULES.mch || [];
  return [];
}

function parserRuleToNormalized(rule, explicit = null) {
  const titleParts = {
    base: rule.base,
    period: rule.period || "",
    suffix: rule.suffix || "",
  };
  const location = rule.location || "";
  if (rule.allDay) {
    return {
      kind: rule.kind || "shift",
      titleParts,
      location,
      includeAsShift: rule.includeAsShift !== false,
      allDay: true,
      defaultTimes: null,
    };
  }
  const startHm = parseRuleTime(rule.startTime);
  const endHm = parseRuleTime(rule.endTime);
  if (!startHm || !endHm) return null;
  return {
    kind: rule.kind || "shift",
    titleParts,
    location,
    includeAsShift: rule.includeAsShift !== false,
    allDay: false,
    defaultTimes: [startHm, endHm],
    ambiguous: false,
    warning: "",
  };
}

function parseRuleTime(value) {
  const match = String(value || "").match(/^(\d{2}):(\d{2})$/);
  if (!match) return null;
  return [Number(match[1]), Number(match[2])];
}

function buildDefaultParserRules() {
  const rules = { mmc: [], ddh: [], casey: [], mch: [] };
  const activeSeniorities = SENIORITY_LABELS.filter((item) => item !== UNKNOWN_SENIORITY);
  const consultantSeniorities = ["SMS", "CMO"];
  // Senior Registrars can be rostered acting-up consultant allocations. A
  // consultant-coded MMC shift means the same work and hours for them.
  const actingConsultantSeniorities = [...consultantSeniorities, "Senior Registrar"];
  const nonConsultantMmcSeniorities = ["Senior Registrar", "Transitional/Intermediate Registrar", "Junior Registrar", "HMO", "Intern"];
  const add = (bucket, source, code, seniority, base, period, suffix, allDay, startTime, endTime, location = defaultParserRuleLocation(source)) => {
    bucket.push({
      source,
      seniority,
      code,
      kind: "shift",
      base,
      period,
      suffix,
      allDay,
      startTime: allDay ? "" : startTime,
      endTime: allDay ? "" : endTime,
      location: allDay && base === "CS" ? "" : location,
      includeAsShift: true,
    });
  };
  for (const seniority of consultantSeniorities) {
    add(rules.mmc, "MMC", "CS", seniority, "CS", "", "", true, "", "", "");
    add(rules.mmc, "MMC", "CSO", seniority, "CSO", "", "", true, "", "", MMC_LOCATION);
    add(rules.mmc, "MMC", "CS OS", seniority, "CS OS", "", "", false, "08:00", "17:30", MMC_LOCATION);
  }
  for (const seniority of actingConsultantSeniorities) {
    for (const periodPrefix of ["A", "P"]) {
      for (const [teamCode, teamName] of Object.entries(MMC_TEAM_MAP)) {
        for (const suffixCode of ["C", "R"]) {
          if (suffixCode === "R" && (teamCode === "C" || teamCode === "R")) continue;
          add(
            rules.mmc,
            "MMC",
            `${periodPrefix}${teamCode}${suffixCode}`,
            seniority,
            teamName,
            periodPrefix === "A" ? "AM" : "PM",
            suffixCode === "R" ? "Float" : "",
            false,
            periodPrefix === "A" ? "08:00" : "14:30",
            periodPrefix === "A" ? "17:30" : "00:00",
          );
        }
      }
      for (const suffixCode of ["C", "R"]) {
        if (suffixCode === "R") continue;
        add(
          rules.mmc,
          "MMC",
          `${periodPrefix}SS${suffixCode}`,
          seniority,
          "SSU",
          periodPrefix === "A" ? "AM" : "PM",
          suffixCode === "R" ? "Float" : "",
          false,
          periodPrefix === "A" ? "07:30" : "14:30",
          periodPrefix === "A" ? "17:30" : "00:00",
        );
      }
    }
  }
  for (const seniority of nonConsultantMmcSeniorities) {
    add(rules.mmc, "MMC", "AHJ", seniority, "Hub", "AM", "", false, "08:00", "17:30", MMC_LOCATION);
    add(rules.mmc, "MMC", "PHJ", seniority, "Hub", "PM", "", false, "14:30", "00:00", MMC_LOCATION);
    add(rules.mmc, "MMC", "ASSJ", seniority, "SSU", "AM", "", false, "07:30", "17:30", MMC_LOCATION);
    add(rules.mmc, "MMC", "PSSJ", seniority, "SSU", "PM", "", false, "14:30", "00:00", MMC_LOCATION);
    add(rules.mmc, "MMC", "NSSJ", seniority, "Night SSU", "", "", false, "23:00", "08:30", MMC_LOCATION);
  }
  for (const seniority of activeSeniorities) {
    const canWorkClinicalSupport = consultantSeniorities.includes(seniority);
    const isSms = seniority === "SMS";
    add(rules.mmc, "MMC", "SWA", seniority, "Swing", "AM", "", false, "08:00", "17:30", MMC_LOCATION);
    add(rules.mmc, "MMC", "SWP", seniority, "Swing", "PM", "", false, "14:30", "00:00", MMC_LOCATION);
    add(rules.mmc, "MMC", "PHNW", seniority, "PHNW", "", "", true, "", "", "");
    if (canWorkClinicalSupport) {
      add(rules.ddh, "DDH", "CS", seniority, "CS", "", "", true, "", "", "");
      // A Monday CS AM entry is paired with an external HITH PM allocation;
      // the DDH calendar should show the local Clinical Support component only.
      add(rules.ddh, "DDH", "CS AM", seniority, "CS", "", "", true, "", "", "");
      add(rules.ddh, "DDH", "CS ONSITE", seniority, "CS onsite", "", "", true, "", "", DDH_LOCATION);
      add(rules.ddh, "DDH", "CLINICAL SUPPORT ACEM OSCE", seniority, "CS Exam", "", "", true, "", "", DDH_LOCATION);
    }
    add(rules.ddh, "DDH", "PHNW", seniority, "PHNW", "", "", true, "", "", "");
    add(rules.ddh, "DDH", "SSU", seniority, "SSU", "", "", false, "07:30", "17:30", DDH_LOCATION);
    add(rules.ddh, "DDH", "ORANGE AM", seniority, "Orange", "AM", "", false, "08:00", "18:00", DDH_LOCATION);
    add(rules.ddh, "DDH", "ORANGE PM", seniority, "Orange", "PM", "", false, isSms ? "15:00" : "14:30", "00:00", DDH_LOCATION);
    add(rules.ddh, "DDH", "SILVER AM", seniority, "Silver", "AM", "", false, "08:00", "18:00", DDH_LOCATION);
    add(rules.ddh, "DDH", "SILVER PM", seniority, "Silver", "PM", "", false, isSms ? "15:00" : "14:30", "00:00", DDH_LOCATION);
    if (!isSms) add(rules.ddh, "DDH", "FAST AM", seniority, "FAST", "AM", "", false, "08:00", "18:00", DDH_LOCATION);
    add(rules.ddh, "DDH", "FAST PM", seniority, "FAST", "PM", "", false, isSms ? "15:00" : "14:30", "00:00", DDH_LOCATION);
    add(rules.ddh, "DDH", "AVAO AM", seniority, "AVAO", "AM", "", false, isSms ? "07:30" : "08:00", isSms ? "17:00" : "18:00", DDH_LOCATION);
    add(rules.ddh, "DDH", "AVAO PM", seniority, "AVAO", "PM", "", false, "14:30", "00:00", DDH_LOCATION);
    add(rules.ddh, "DDH", "ROVER AM", seniority, "Rover", "AM", "", false, "08:00", "18:00", DDH_LOCATION);
    add(rules.ddh, "DDH", "ROVER PM", seniority, "Rover", "PM", "", false, isSms ? "15:00" : "14:30", "00:00", DDH_LOCATION);
    add(rules.ddh, "DDH", "SWING AM", seniority, "Swing", "AM", "", false, "08:00", "18:00", DDH_LOCATION);
    add(rules.ddh, "DDH", "SWING PM", seniority, "Swing", "PM", "", false, isSms ? "15:00" : "14:30", "00:00", DDH_LOCATION);
    if (canWorkClinicalSupport) {
      add(rules.casey, "Casey", "CS", seniority, "CS", "", "", false, "08:00", "17:30", CASEY_LOCATION);
      add(rules.casey, "Casey", "CLIN SUPP", seniority, "CS", "", "", false, "08:00", "17:30", CASEY_LOCATION);
      add(rules.casey, "Casey", "CLINICAL SUPP", seniority, "CS", "", "", false, "08:00", "17:30", CASEY_LOCATION);
      add(rules.mch, "MCH", "CS", seniority, "CS", "", "", false, "08:00", "17:30", MCH_LOCATION);
      add(rules.mch, "MCH", "OCS", seniority, "CS Office", "", "", false, "08:00", "17:30", MCH_LOCATION);
      add(rules.mch, "MCH", "0CS", seniority, "CS Office", "", "", false, "08:00", "17:30", MCH_LOCATION);
      add(rules.mch, "MCH", "CSOS", seniority, "CS Office", "", "", false, "08:00", "17:30", MCH_LOCATION);
    }
    add(rules.casey, "Casey", "PHNW", seniority, "PHNW", "", "", true, "", "", "");
    add(rules.mch, "MCH", "PHNW", seniority, "PHNW", "", "", true, "", "", "");
    add(rules.casey, "Casey", "AM TL", seniority, "TL", "AM", "", false, "08:00", "17:30", CASEY_LOCATION);
    add(rules.casey, "Casey", "AM UFD", seniority, "UFD", "AM", "", false, "08:00", "17:30", CASEY_LOCATION);
    add(rules.casey, "Casey", "AM MIC", seniority, "MIC", "AM", "", false, "08:00", "17:30", CASEY_LOCATION);
    add(rules.casey, "Casey", "PM TL", seniority, "TL", "PM", "", false, "14:30", "00:00", CASEY_LOCATION);
    add(rules.casey, "Casey", "PM UFD", seniority, "UFD", "PM", "", false, "14:30", "00:00", CASEY_LOCATION);
    add(rules.casey, "Casey", "PM MIC", seniority, "MIC", "PM", "", false, "14:30", "00:00", CASEY_LOCATION);
    add(rules.casey, "Casey", "PM PAEDS", seniority, "PAEDS", "PM", "", false, "14:30", "00:00", CASEY_LOCATION);
    add(rules.casey, "Casey", "AM PAEDS", seniority, "PAEDS", "AM", "", false, "07:30", "17:00", CASEY_LOCATION);
    add(rules.casey, "Casey", "AM SSU", seniority, "SSU", "AM", "", false, "07:30", "17:00", CASEY_LOCATION);
    add(rules.casey, "Casey", "AM SWING", seniority, "Swing", "AM", "", false, "08:00", "17:30", CASEY_LOCATION);
    add(rules.casey, "Casey", "PM SWING", seniority, "Swing", "PM", "", false, "14:30", "00:00", CASEY_LOCATION);
  }
  // Some legacy DDH worksheet layouts do not retain the section seniority on
  // a staff row. Keep these two unambiguous Clinical Support labels usable in
  // those files as well.
  add(rules.ddh, "DDH", "CS AM", UNKNOWN_SENIORITY, "CS", "", "", true, "", "", "");
  add(rules.ddh, "DDH", "CLINICAL SUPPORT ACEM OSCE", UNKNOWN_SENIORITY, "CS Exam", "", "", true, "", "", DDH_LOCATION);
  for (const seniority of activeSeniorities) {
    add(rules.mmc, "MMC", "CSM", seniority, "CSM", "", "", false, "08:00", "17:30", MMC_LOCATION);
  }
  add(rules.mmc, "MMC", "CS EXAM", "SMS", "CS Exam", "", "", true, "", "", MMC_LOCATION);
  return rules;
}

function defaultParserRuleLocation(source) {
  if (source === "MMC") return MMC_LOCATION;
  if (source === "DDH") return DDH_LOCATION;
  if (source === "Casey") return CASEY_LOCATION;
  if (source === "MCH") return MCH_LOCATION;
  return "";
}


function inferMmcTimeOnlyShiftLabel(startHm) {
  if (!Array.isArray(startHm)) return "AM";
  const [hour] = startHm;
  if (hour >= 22 || hour < 6) return "Night";
  if (hour >= 14) return "PM";
  return "AM";
}

function resolveDdhShiftLabel(label, seniority = UNKNOWN_SENIORITY) {
  const mapped = DDH_LABEL_MAP[label] || label;
  const normalized = findManualParserRule("DDH", seniority, mapped) || normalizeDdhLabel(mapped) || normalizeGenericDdhLabel(mapped);
  return { mapped, normalized };
}

function normalizeDdhLabel(label) {
  if (!KNOWN_DDH_DIRECT_LABELS.has(label)) {
    return null;
  }

  if (label === "PHNW") {
    return { kind: "public_holiday", titleParts: { base: "PHNW", period: "", suffix: "" } };
  }
  if (label === "CS" || label === "CS onsite" || label === "SSU") {
    return { kind: "shift", titleParts: { base: label, period: "", suffix: "" } };
  }

  const parts = label.split(/\s+/);
  const last = parts.at(-1);
  if (last === "AM" || last === "PM") {
    return {
      kind: "shift",
      titleParts: {
        base: parts.slice(0, -1).join(" "),
        period: last,
        suffix: "",
      },
    };
  }

  return { kind: "shift", titleParts: { base: label, period: "", suffix: "" } };
}

function normalizeGenericDdhLabel(label) {
  const cleaned = cleanDdhLabel(label);
  if (!cleaned) return null;
  const recognisedSlot = normalizeDdhRosterSlotLabel(cleaned);
  if (recognisedSlot) return recognisedSlot;
  const period = extractDdhPeriod(cleaned);
  const upper = cleaned.toUpperCase();

  if (upper.includes("CLINICAL SUPPORT") || /^CS\b/.test(upper) || upper.includes(" OCS")) {
    const onsite = upper.includes("ONSITE");
    return genericUnknownDdhShift({ base: onsite ? "CS onsite" : "CS", period, suffix: "" });
  }
  if (upper.includes("SSU")) return genericUnknownDdhShift({ base: upper.includes("NIGHT") ? "Night SSU" : "SSU", period, suffix: "" });
  if (upper.includes("AVAO")) return genericUnknownDdhShift({ base: "AVAO", period, suffix: "" });
  if (upper.includes("ORANGE")) return genericUnknownDdhShift({ base: "Orange", period, suffix: "" });
  if (upper.includes("SILVER")) return genericUnknownDdhShift({ base: "Silver", period, suffix: "" });
  if (upper.includes("FAST")) return genericUnknownDdhShift({ base: "FAST", period, suffix: "" });
  if (upper.includes("VHH")) return genericUnknownDdhShift({ base: period ? "VHH" : cleaned.replace(/\bIC\b/gi, "").trim(), period, suffix: "" });
  if (upper.includes("ROVER")) return genericUnknownDdhShift({ base: "Rover", period, suffix: "" });
  if (upper.includes("HITH")) return genericUnknownDdhShift({ base: "HITH", period, suffix: "" });
  if (upper.includes("PAED")) return genericUnknownDdhShift({ base: "Paeds", period, suffix: "" });
  if (upper.includes("NIGHT")) return genericUnknownDdhShift({ base: cleaned.replace(/\bIC\b/gi, "").trim(), period: "", suffix: "" });
  if (upper.includes("AED")) return genericUnknownDdhShift({ base: "AED", period, suffix: "" });
  if (upper.includes("MED")) return genericUnknownDdhShift({ base: "MED", period, suffix: "" });
  if (upper.includes("GED")) return genericUnknownDdhShift({ base: cleaned.replace(/\bIC\b/gi, "").trim(), period: "", suffix: "" });
  if (upper.includes("EXTRA")) return genericUnknownDdhShift({ base: "Extra", period, suffix: "" });

  return null;
}

// DDH grids use a trailing number for parallel positions (for example,
// "Orange AM2" or "Night4") and sometimes prefix the role. Those markers do
// not alter the shift itself. Treat only these deliberately narrow forms as
// known shifts; free-text notes remain visible as unresolved review items.
function normalizeDdhRosterSlotLabel(value) {
  const cleaned = String(value || "").replace(/\s+/g, " ").trim();
  const upper = cleaned.toUpperCase();
  if (!upper) return null;
  if (/^NIGHT(?:\s+SSU)?$/.test(upper)) {
    return {
      kind: "shift",
      titleParts: { base: upper.includes("SSU") ? "Night SSU" : "Night", period: "", suffix: "" },
      status: "ok",
      warning: "",
    };
  }
  if (/^GED(?:\s+JUNIOR)?$/.test(upper)) {
    return {
      kind: "shift",
      titleParts: { base: "GED shift", period: "", suffix: "" },
      status: "ok",
      warning: "",
    };
  }
  if (/^(?:EXTRA\s+)?SWING$/.test(upper)) {
    return {
      kind: "shift",
      titleParts: { base: "Swing shift", period: "", suffix: "" },
      allDay: true,
      status: "ok",
      warning: "",
    };
  }
  const period = extractDdhPeriod(upper);
  const baseMatch = upper.match(/\b(ORANGE|SILVER|FAST|AVAO|ROVER|SSU|SWING)\b/);
  if (baseMatch && (period || baseMatch[1] === "SSU")) {
    const baseNames = {
      ORANGE: "Orange",
      SILVER: "Silver",
      FAST: "FAST",
      AVAO: "AVAO",
      ROVER: "Rover",
      SSU: "SSU",
      SWING: "Swing",
    };
    return {
      kind: "shift",
      titleParts: { base: baseNames[baseMatch[1]], period, suffix: "" },
      status: "ok",
      warning: "",
    };
  }
  const extra = upper.match(/^EXTRA(?:\s+(AVAO))?\s+(AM|PM)$/);
  if (extra) {
    return {
      kind: "shift",
      titleParts: { base: extra[1] ? "Extra AVAO" : "Extra", period: extra[2], suffix: "" },
      status: "ok",
      warning: "",
    };
  }
  return null;
}

function ddhSwingTitlePartsForStart(label, startHm, fallback) {
  if (!/^(?:EXTRA\s+)?SWING$/i.test(String(label || "").trim()) || !Array.isArray(startHm)) return fallback;
  const minutes = Number(startHm[0]) * 60 + Number(startHm[1]);
  if (minutes < 14 * 60) return { base: "Swing", period: "AM", suffix: "" };
  if (minutes > 15 * 60) return { base: "Swing", period: "PM", suffix: "" };
  return { base: "Swing shift", period: "", suffix: "" };
}

function genericUnknownDdhShift(titleParts) {
  return {
    kind: "shift",
    titleParts,
    status: "unknown",
    warning: "DDH shift label not recognised; using roster time.",
  };
}

function normalizeDdhLocation(label, normalized) {
  const upper = String(label || "").toUpperCase();
  if (normalized.titleParts.base === "PHNW" || normalized.titleParts.base === "CS") return "";
  if (upper.includes("OFFSITE") || upper.includes("NOT ONSITE") || upper.includes("CS/OFF")) return "";
  if (normalized.titleParts.base === "CS onsite") return DDH_LOCATION;
  return DDH_LOCATION;
}

function createOrientationRecord(source, day, rawValue, seniority = UNKNOWN_SENIORITY) {
  const label = String(rawValue || "").trim();
  // Casey uses the abbreviated form "Orient 09-1730". It is a real,
  // timed orientation shift, not metadata to be stripped from a roster cell.
  if (!/^(?:ORIENTATION|ORIENATION|ORIENT)\b/i.test(label)) return null;
  const range = label.match(/\b(\d{1,2})(?::?(\d{2}))?\s*(?:-|–|—|TO)\s*(\d{1,2})(?::?(\d{2}))?\b/i);
  const location = source === "MMC" ? MMC_LOCATION
    : source === "DDH" ? DDH_LOCATION
      : source === "Casey" ? CASEY_LOCATION
        : MCH_LOCATION;
  const details = {
    kind: "shift",
    titleParts: { base: "Orientation", period: "", suffix: "" },
    location,
    seniority,
  };
  if (!range) return createAllDayRecord(source, day, rawValue, details);
  const startHm = [Number(range[1]), Number(range[2] || 0)];
  const endHm = [Number(range[3]), Number(range[4] || 0)];
  if (startHm[0] > 23 || endHm[0] > 23 || startHm[1] > 59 || endHm[1] > 59) {
    return createAllDayRecord(source, day, rawValue, details);
  }
  return createTimedRecord(source, day, rawValue, { ...details, startHm, endHm });
}

function cleanDdhLabel(label) {
  return String(label || "")
    .replace(/\([^)]*\)/g, " ")
    .replace(/\bIC\b/gi, " ")
    .replace(/\d+\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function extractDdhPeriod(label) {
  const upper = String(label || "").toUpperCase();
  if (/\bAM\b/.test(upper)) return "AM";
  if (/\bPM\b/.test(upper)) return "PM";
  return "";
}

function shouldIgnoreDdh(value, seniority = UNKNOWN_SENIORITY) {
  const upper = String(value || "").trim().toUpperCase();
  if (!upper) return true;
  if (isDdhClinicalSupportExam(upper)) return false;
  if (isDdhStructuralAnnotation(upper, seniority)) return true;
  if (DDH_IGNORE_PREFIXES.some((prefix) => upper.startsWith(prefix))) return true;
  return DDH_IGNORE_CONTAINS.some((fragment) => upper.includes(fragment));
}

function isDdhClinicalSupportExam(value) {
  return /^CLINICAL\s+SUPPORT\s+ACEM\s+OSCE$/i.test(String(value || "").trim());
}

// FindMyShift carries staff headings, availability, and free-text rostering
// notes in the same Shift label column. These must not become calendar shifts
// or unresolved shift-code diagnostics.
function isDdhStructuralAnnotation(value, seniority = UNKNOWN_SENIORITY) {
  const upper = String(value || "").trim().toUpperCase();
  if ([
    "INTERN", "INTERNS", "UNAVAILABLE", "UNAVAILABE", "-", "--", "SEC", "N", "Y", "W",
    // Frank Soden's DDH cells are roster-writer annotations, not shifts.
    "MDC", "NA", "MDC NA", "NA PM", "COMM", "COMM FOR SAFETY",
  ].includes(upper)) return true;
  // The AMP section contains supervision and free-text notes (including
  // physiotherapist headings and staff names), not rostered AMP shifts.
  // Restrict this exclusion to AMP so identically shaped labels elsewhere
  // remain available for normal shift parsing.
  const ampSupervisionName = looksLikePersonName(value)
    && !/\b(?:AM|PM|NIGHT|SWING|FAST|SSU|ORANGE|SILVER|AVAO|ROVER|GED|ORIENTATION|ORIENATION|ORIENT)\b/.test(upper);
  if (sanitizeRuleSeniority(seniority) === "AMP" && (upper === "PHYSIOTHERAPIST" || upper === "PHYSIOTHERAPISTS" || ampSupervisionName)) return true;
  if (/^(?:AM|PM)\s*>\s*(?:AM|PM)$/.test(upper)) return true;
  if (/^\(?\d+\s+ED\s+SHIFTS?\s+THIS\s+WEEK\)?$/.test(upper)) return true;
  return false;
}

function shouldIgnoreCasey(value) {
  const upper = cleanText(value).replace(/\s+/g, " ").trim().toUpperCase();
  if (CASEY_IGNORED_EXACT.has(upper)) return true;
  if (/^[A-Z]{2,}(?:\s+[A-Z])?$/.test(upper) && upper.length <= 12 && !caseyTimedRuleForCode(upper)) return true;
  if (/^(MON|TUES|WED|THU|FRI|SAT|SUN)\b/.test(upper)) return true;
  return false;
}

function stripCaseyOrientationMetadata(value) {
  return cleanText(value).replace(/\bOrient(?:ation)?\b/gi, " ").replace(/\s+/g, " ").trim();
}

function shouldIgnoreMch(value) {
  const upper = cleanText(value).replace(/\s+/g, " ").trim().toUpperCase();
  if (!upper || upper === "0" || upper === "N/A" || upper === "NA") return true;
  if (/^(MONDAY|TUESDAY|WEDNESDAY|THURSDAY|FRIDAY|SATURDAY|SUNDAY)$/.test(upper)) return true;
  if (/^TRG_/.test(upper) || /^CON_/.test(upper) || upper === "59-OC-16") return true;
  if (upper.includes("TEACHING") || upper.includes("ORIENTATION")) return true;
  return false;
}

function iterateDdhWeekEntries(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const findmyshiftFormat = cleanText(getCellValue(sheet, 1, 1)) === "FindMyShift roster format";
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const dateRows = [];
  for (let row = 1; row <= range.e.r + 1; row += 1) {
    if (isDdhDateRow(sheet, row)) dateRows.push(row);
  }

  const entries = [];
  for (let index = 0; index < dateRows.length; index += 1) {
    const dateRow = dateRows[index];
    const nextDateRow = dateRows[index + 1] || range.e.r + 2;
    const weekDates = [];
    for (let col = 2; col <= 8; col += 1) {
      const value = parseDdhDate(getCellValue(sheet, dateRow, col));
      if (!value) {
        weekDates.length = 0;
        break;
      }
      weekDates.push(value);
    }
    if (!weekDates.length) continue;

    let currentSeniority = UNKNOWN_SENIORITY;
    const sectionMap = ddhSeniorityMap();
    for (let row = dateRow + 1; row < nextDateRow; row += 1) {
      const rawName = cleanText(getCellValue(sheet, row, 1));
      const upperName = rawName.replace(/\s+/g, " ").trim().toUpperCase();
      if (sectionMap.has(upperName)) {
        currentSeniority = sectionMap.get(upperName);
        continue;
      }
      if (/^ED HMO/i.test(upperName) || /^HMO\b/i.test(upperName)) {
        currentSeniority = "HMO";
        continue;
      }
      if (!rawName || !looksLikePersonName(rawName) || isDdhSectionMarker(rawName)) continue;
      const labels = [];
      for (let col = 2; col <= 8; col += 1) labels.push(cleanText(getCellValue(sheet, row, col)));
      const supplementaryRow = row + 1 < nextDateRow && isDdhSupplementaryRow(sheet, row + 1) ? row + 1 : 0;
      const times = [];
      for (let col = 2; col <= 8; col += 1) {
        times.push(supplementaryRow ? cleanText(getCellValue(sheet, supplementaryRow, col)) : "");
      }
      const seniority = currentSeniority === UNKNOWN_SENIORITY
        ? upcomingDdhSectionSeniority(sheet, row, nextDateRow, sectionMap)
        : currentSeniority;
      entries.push({ rawName, weekDates, labels, times, seniority, findmyshiftFormat });
      if (supplementaryRow) row = supplementaryRow;
    }
  }
  return entries;
}

function upcomingDdhSectionSeniority(sheet, row, nextDateRow, sectionMap) {
  for (let candidateRow = row + 1; candidateRow < nextDateRow; candidateRow += 1) {
    const rawName = cleanText(getCellValue(sheet, candidateRow, 1));
    if (!rawName) continue;
    const upperName = rawName.replace(/\s+/g, " ").trim().toUpperCase();
    if (sectionMap.has(upperName)) return sectionMap.get(upperName);
    if (isDdhSectionMarker(rawName) || looksLikePersonName(rawName)) break;
  }
  return UNKNOWN_SENIORITY;
}

function iterateMchWeekEntries(workbook) {
  const entries = [];
  for (const sheetName of workbook.SheetNames || []) {
    if (!isMchWeekSheet(workbook.Sheets[sheetName], sheetName)) continue;
    const sheet = workbook.Sheets[sheetName];
    const weekDates = mchWeekDates(sheet);
    if (!weekDates.length) continue;

    const ranges = [
      [21, 49],
      [54, 84],
      [86, 102],
    ];
    for (const [startRow, endRow] of ranges) {
      for (let row = startRow; row <= endRow; row += 1) {
        const rawName = cleanMchRosterName(getCellValue(sheet, row, 4));
        const role = cleanText(getCellValue(sheet, row, 1));
        if (!rawName || !looksLikePersonName(rawName) || isMchIgnoredName(rawName)) continue;
        const labels = [];
        for (let col = 6; col <= 12; col += 1) labels.push(cleanText(getCellValue(sheet, row, col)));
        entries.push({
          rawName,
          displayName: rawName,
          weekDates,
          labels,
          seniority: mchSeniorityForRole(role),
        });
      }
    }
  }
  return entries;
}

function iterateCaseyWeekEntries(workbook) {
  const entries = [];
  for (const sheetName of workbook.SheetNames || []) {
    if (!isCaseyWeekSheet(workbook.Sheets[sheetName], sheetName)) continue;
    const sheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
    const termYear = caseyTermYear(sheet);
    const weekDates = [];
    for (let col = 2; col <= 8; col += 1) {
      const value = parseCaseyWeekDate(sheet, 2, col, termYear);
      if (!value) {
        weekDates.length = 0;
        break;
      }
      weekDates.push(value);
    }
    if (!weekDates.length) continue;

    let currentSeniority = "SMS";
    const sectionMap = caseySeniorityMap();
    for (let row = 3; row <= range.e.r + 1; row += 1) {
      const rawName = cleanText(getCellValue(sheet, row, 1));
      const upperName = rawName.replace(/\s+/g, " ").trim().toUpperCase();
      if (!upperName) continue;
      if (CASEY_STOP_SECTIONS.has(upperName)) break;
      if (sectionMap.has(upperName)) {
        currentSeniority = sectionMap.get(upperName);
        continue;
      }
      if (upperName.startsWith("ONCALL")) continue;
      if (isCaseySectionMarker(rawName)) continue;
      const parsedName = parseCaseyRosterName(rawName);
      if (!parsedName || !looksLikePersonName(parsedName.name)) continue;
      const labels = [];
      for (let col = 2; col <= 8; col += 1) labels.push(getCaseyRosterCellText(sheet, row, col));
      entries.push({
        rawName: parsedName.name,
        displayName: parsedName.name,
        weekDates,
        labels,
        seniority: parsedName.seniority || currentSeniority,
      });
    }
  }
  return entries;
}

function getCaseyRosterCellText(sheet, row, col) {
  const direct = cleanText(getCellValue(sheet, row, col));
  if (direct || col < 2 || col > 8) return direct;
  const merged = caseyMergedRangeForCell(sheet, row, col);
  if (!merged) return "";
  return cleanText(getCellValue(sheet, merged.s.r + 1, merged.s.c + 1));
}

function caseyMergedRangeForCell(sheet, row, col) {
  const cellRow = row - 1;
  const cellCol = col - 1;
  for (const merge of sheet["!merges"] || []) {
    if (merge.s.r !== cellRow || merge.s.c < 1 || merge.e.c > 7) continue;
    if (merge.s.c === merge.e.c) continue;
    if (cellCol >= merge.s.c && cellCol <= merge.e.c) return merge;
  }
  return null;
}

function isCaseyWorkbook(workbook) {
  return (workbook.SheetNames || []).some((sheetName) => isCaseyWeekSheet(workbook.Sheets[sheetName], sheetName));
}

function isMchWorkbook(workbook) {
  return (workbook.SheetNames || []).some((sheetName) => isMchWeekSheet(workbook.Sheets[sheetName], sheetName));
}

function isMchWeekSheet(sheet, sheetName) {
  if (!sheet || !/^Week\s+\d+$/i.test(String(sheetName || "")) || String(sheetName).trim() === "Week 0") return false;
  const title = cleanText(getCellValue(sheet, 4, 4)).toUpperCase();
  if (!title.includes("PAEDIATRIC EMERGENCY DEPARTMENT ROSTER")) return false;
  const nameHeader = cleanText(getCellValue(sheet, 18, 4)).toUpperCase();
  const monday = cleanText(getCellValue(sheet, 18, 6)).toUpperCase();
  return nameHeader === "NAME" && monday === "MONDAY" && mchWeekDates(sheet).length === 7;
}

function mchWeekDates(sheet) {
  const dates = [];
  for (let col = 6; col <= 12; col += 1) {
    const value = coerceDate(getCellValue(sheet, 19, col)) || coerceDate(getCellValue(sheet, 2, col));
    if (!value) return [];
    dates.push(value);
  }
  return dates;
}

function isCaseyWeekSheet(sheet, sheetName) {
  if (!sheet || sheetName === "Personal Roster" || sheetName === "DayRoster") return false;
  const title = cleanText(getCellValue(sheet, 1, 1)).toUpperCase();
  if (!/^TERM\s+\d+,/.test(title)) return false;
  const weekdays = [];
  for (let col = 2; col <= 8; col += 1) weekdays.push(cleanText(getCellValue(sheet, 1, col)).toUpperCase());
  if (weekdays.join("|") !== "MONDAY|TUESDAY|WEDNESDAY|THURSDAY|FRIDAY|SATURDAY|SUNDAY") return false;
  const termYear = caseyTermYear(sheet);
  return Boolean(parseCaseyWeekDate(sheet, 2, 2, termYear) && parseCaseyWeekDate(sheet, 2, 8, termYear));
}

function caseyTermYear(sheet) {
  const match = cleanText(getCellValue(sheet, 1, 1)).match(/\b(20\d{2})\b/);
  return match ? Number(match[1]) : 0;
}

function parseCaseyWeekDate(sheet, row, col, termYear) {
  const address = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
  const cell = sheet[address];
  const formatted = cleanText(cell?.w || "");
  const text = formatted || cleanText(cell?.v || "");
  const match = text.match(/^(\d{1,2})[-/ ]([A-Za-z]{3,})$/);
  if (match && termYear) {
    const month = monthIndex(match[2]);
    if (month >= 0) return formatDateOnly(new Date(termYear, month, Number(match[1])));
  }
  const coerced = coerceDate(cell?.v);
  if (!coerced || !termYear) return coerced;
  const date = new Date(`${coerced}T00:00:00`);
  date.setFullYear(termYear);
  return formatDateOnly(date);
}

function monthIndex(value) {
  const index = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"].indexOf(String(value || "").slice(0, 3).toUpperCase());
  return index;
}

function parseCaseyRosterName(value) {
  const cleaned = cleanText(value)
    .replace(/\u00a0/g, " ")
    .replace(/\s+-\s+LOCUM(?:\s+SMS)?$/i, "")
    .replace(/\s+\(7\/fn\)$/i, "")
    .replace(/\s+/g, " ")
    .trim();
  if (!cleaned || cleaned.toUpperCase().includes("TBA")) return null;
  const match = cleaned.match(/^\(([^)]+)\)\s*(.+)$/);
  if (!match) return { name: cleaned, seniority: "" };
  return {
    name: match[2].trim(),
    seniority: sanitizeRuleSeniority(match[1]),
  };
}

function cleanMchRosterName(value) {
  return cleanText(value)
    .replace(/\u00a0/g, " ")
    .replace(/\s*\([^)]*\)\s*$/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function isMchIgnoredName(value) {
  const upper = cleanText(value).replace(/\s+/g, " ").trim().toUpperCase();
  if (!upper || upper === "LOCUM" || upper === "MEDICAL STUDENT" || upper === "NOT FOUND") return true;
  if (/^\d+(?:\.\d+)?$/.test(upper)) return true;
  return false;
}

function mchSeniorityForRole(value) {
  const upper = cleanText(value).replace(/\s+/g, " ").trim().toUpperCase();
  if (upper.includes("CONSULTANT") || upper.includes("STAFF SPECIALIST")) return "SMS";
  if (upper.includes("FELLOW")) return "Senior Registrar";
  if (upper.includes("REGISTRAR")) return "Junior Registrar";
  if (upper.includes("HMO")) return "HMO";
  if (upper.includes("INTERN")) return "Intern";
  return UNKNOWN_SENIORITY;
}

function isCaseySectionMarker(value) {
  const upper = cleanText(value).replace(/\s+/g, " ").trim().toUpperCase();
  return CASEY_SECTION_MARKERS.has(upper);
}

function mmcSeniorityMap() {
  return new Map([
    ["SMS", "SMS"],
    ["GERIATRICIAN", "SMS"],
    ["CMO", "CMO"],
    ["SENIOR REG", "Senior Registrar"],
    ["INTERMEDIATE REG", "Transitional/Intermediate Registrar"],
    ["JUNIOR REG", "Junior Registrar"],
    ["HMO", "HMO"],
    ["HMO MUST BE 111", "HMO"],
    ["HMO - MUST BE 111", "HMO"],
    ["ENP", "ENP"],
    ["AMP", "AMP"],
    ["EMERGENCY NURSE PRACTITIONER", "ENP"],
    ["AMBULATORY MUSCULOSKELETAL PHYSIOTHERAPIST", "AMP"],
    ["INTERN", "Intern"],
  ]);
}

function ddhSeniorityMap() {
  return new Map([
    ["SENIOR MEDICAL STAFF", "SMS"],
    ["SENIOR REGISTRARS", "Senior Registrar"],
    ["REGISTRAR", "Senior Registrar"],
    ["REGISTRARS", "Senior Registrar"],
    ["CMO'S", "CMO"],
    ["CMOS", "CMO"],
    ["JUNIOR REGISTRARS", "Junior Registrar"],
    ["ED HMO'S", "HMO"],
    ["ED HMOS", "HMO"],
    ["HMO'S", "HMO"],
    ["HMOS", "HMO"],
    ["INTERNS", "Intern"],
    ["ENP", "ENP"],
    ["NURSE PRACTITIONERS", "ENP"],
    ["NURSE PRAC. CANDIDATES", "ENP"],
    ["AMP", "AMP"],
    ["AMP'S", "AMP"],
    ["PHYSIOTHERAPIST", "AMP"],
    ["PHYSIOTHERAPISTS", "AMP"],
  ]);
}

function caseySeniorityMap() {
  return new Map([
    ["GERIATRICIAN", "SMS"],
    ["SNR CMOS", "CMO"],
    ["SENIOR CMOS", "CMO"],
    ["ED GPS", "CMO"],
    ["SNR REGS", "Senior Registrar"],
    ["SENIOR REGS", "Senior Registrar"],
    ["JNR REGS", "Junior Registrar"],
    ["JUNIOR REGS", "Junior Registrar"],
    ["PAEDS ED HMOS", "HMO"],
    ["HMOS", "HMO"],
    ["HMO", "HMO"],
    ["INTERNS", "Intern"],
    ["AMP COVER", "AMP"],
  ]);
}

function createTimedRecord(source, day, rawValue, details) {
  const start = buildDateTime(day, details.startHm);
  const plusDay = compareTimes(details.endHm, details.startHm) <= 0;
  const end = buildDateTime(day, details.endHm, plusDay);
  const normalizedTitle = formatTitle(source, details.titleParts, { ...DEFAULT_SETTINGS, showTimes: false, showRawValues: false }, details.kind);
  return {
    id: hashString(`${source}|${day}|${rawValue}|${normalizedTitle}|${start}|${end}`),
    source,
    seniority: sanitizeRuleSeniority(details.seniority),
    kind: details.kind,
    rawValue,
    startDay: day,
    endDay: asDateString(end),
    allDay: false,
    start,
    end,
    location: details.location || "",
    titleParts: details.titleParts,
    normalizedTitle,
    status: details.status || (details.ambiguous ? "ambiguous" : "ok"),
    warnings: details.warning ? [details.warning] : [],
    exportable: true,
    includeByDefault: true,
  };
}

function createAllDayRecord(source, day, rawValue, details) {
  const normalizedTitle = formatTitle(source, details.titleParts, { ...DEFAULT_SETTINGS, showTimes: false, showRawValues: false }, details.kind);
  return {
    id: hashString(`${source}|${day}|${rawValue}|${normalizedTitle}|all-day`),
    source,
    seniority: sanitizeRuleSeniority(details.seniority),
    kind: details.kind,
    rawValue,
    startDay: day,
    endDay: addDays(day, 1),
    allDay: true,
    start: day,
    end: addDays(day, 1),
    location: details.location || "",
    titleParts: details.titleParts,
    normalizedTitle,
    status: details.status || (details.ambiguous ? "ambiguous" : "ok"),
    warnings: details.warning ? [details.warning] : [],
    exportable: true,
    includeByDefault: true,
  };
}

function createWeeklyLeaveRecord(source, monday, rawValue, seniority = UNKNOWN_SENIORITY) {
  const leave = normalizeRecognizedLeave(rawValue) || { kind: "annual_leave", title: "Annual Leave" };
  const { kind } = leave;
  const normalizedTitle = leave.title;
  return {
    id: hashString(`${source}|${monday}|${rawValue}|week-leave`),
    source,
    seniority: sanitizeRuleSeniority(seniority),
    kind,
    rawValue,
    startDay: monday,
    endDay: addDays(monday, 7),
    allDay: true,
    start: monday,
    end: addDays(monday, 7),
    location: "",
    titleParts: { base: normalizedTitle, period: "", suffix: "" },
    normalizedTitle,
    status: "ok",
    warnings: [],
    exportable: true,
    includeByDefault: true,
  };
}

function createUnknownRecord(source, day, rawValue, warning, seniority = UNKNOWN_SENIORITY) {
  return {
    id: hashString(`${source}|${day}|${rawValue}|unknown`),
    source,
    seniority: sanitizeRuleSeniority(seniority),
    kind: "unknown",
    rawValue,
    startDay: day,
    endDay: addDays(day, 1),
    allDay: true,
    start: day,
    end: addDays(day, 1),
    location: "",
    titleParts: { base: "", period: "", suffix: "" },
    normalizedTitle: "",
    status: "unknown",
    warnings: [warning],
    exportable: true,
    includeByDefault: false,
  };
}

function createHiddenRecord(source, day, rawValue, normalized, seniority = UNKNOWN_SENIORITY) {
  const titleParts = normalized.titleParts || { base: rawValue, period: "", suffix: "" };
  const normalizedTitle = formatTitle(source, titleParts, { ...DEFAULT_SETTINGS, showTimes: false, showRawValues: false }, normalized.kind || "shift");
  return {
    id: hashString(`${source}|${day}|${rawValue}|${normalizedTitle}|hidden`),
    source,
    seniority: sanitizeRuleSeniority(seniority),
    kind: "hidden_shift",
    rawValue,
    startDay: day,
    endDay: addDays(day, 1),
    allDay: true,
    start: day,
    end: addDays(day, 1),
    location: "",
    titleParts,
    normalizedTitle,
    status: "ok",
    warnings: [],
    exportable: false,
    includeByDefault: false,
  };
}

function applySettings(records, settings, overrides) {
  const scopedRecords = records.filter((record) => matchesHospitalFilter(record, settings) && matchesDateFilter(record, settings));
  let events = [];
  const reviewItems = [];
  const issues = [];

  for (const record of scopedRecords) {
    if (record.kind === "hidden_shift") continue;
    const override = overrides[record.id] || {};
    const defaultInclude = record.includeByDefault && isKindEnabled(record.kind, settings);
    const include = typeof override.include === "boolean" ? override.include : defaultInclude;
    const suggestedTitle = formatTitle(record.source, record.titleParts, settings, record.kind);
    const overrideTitle = override.title || "";
    const finalTitle = overrideTitle || suggestedTitle;
    const timeLabel = record.allDay ? "All day" : formatTimeLabel(record.start, record.end);
    const location = settings.includeLocations ? resolveDefaultLocation(record.source, record.location, settings) : "";

    reviewItems.push({
      id: record.id,
      source: record.source,
      seniority: record.seniority || UNKNOWN_SENIORITY,
      startDay: record.startDay,
      endDay: record.endDay,
      rawValue: record.rawValue,
      normalizedTitle: record.normalizedTitle,
      suggestedTitle,
      overrideTitle,
      status: record.status,
      warnings: record.warnings,
      include,
      exportable: record.exportable,
      location,
      allDay: record.allDay,
      timeLabel,
    });

    if (record.status !== "ok" || record.warnings.length) {
      issues.push({
        id: record.id,
        source: record.source,
        seniority: record.seniority || UNKNOWN_SENIORITY,
        startDay: record.startDay,
        rawValue: record.rawValue,
        status: record.status,
        message: record.warnings[0] || "Review this roster entry before export.",
        resolutionType: record.status === "unknown" || record.warnings.some((warning) => /shift (code|label) not recognised/i.test(warning))
          ? "shift_code"
          : "review",
        suggestedTitle,
        timeLabel,
      });
    }

    if (!include || !record.exportable || !finalTitle) continue;

    events.push({
      id: record.id,
      source: record.source,
      seniority: record.seniority || UNKNOWN_SENIORITY,
      title: finalTitle,
      allDay: record.allDay,
      start: record.start,
      end: record.end,
      location,
      rawValue: record.rawValue,
      timeLabel,
      monthKey: record.startDay.slice(0, 7),
    });
  }

  events = mergeContiguousLeaveEvents(events);

  events.sort((left, right) => {
    const leftDate = asDateString(left.start);
    const rightDate = asDateString(right.start);
    if (leftDate !== rightDate) return leftDate.localeCompare(rightDate);
    if (left.allDay !== right.allDay) return left.allDay ? -1 : 1;
    return left.title.localeCompare(right.title);
  });

  reviewItems.sort((left, right) => {
    if (left.startDay !== right.startDay) return left.startDay.localeCompare(right.startDay);
    if (left.source !== right.source) return left.source.localeCompare(right.source);
    return left.rawValue.localeCompare(right.rawValue);
  });

  return { events, reviewItems, issues };
}

function mergeContiguousLeaveEvents(events) {
  const passthrough = [];
  const leaveEvents = [];
  for (const event of events) {
    if (!isMergeableLeaveEvent(event)) {
      passthrough.push(event);
      continue;
    }
    leaveEvents.push(event);
  }

  const merged = [];
  const ordered = [...leaveEvents].sort((left, right) => left.start.localeCompare(right.start) || left.end.localeCompare(right.end));
  for (const event of ordered) {
    const previous = merged.length ? merged[merged.length - 1] : null;
    const sameLeaveType = previous
      && preferredLeaveTitle(previous.title, "", previous.rawValue) === preferredLeaveTitle(event.title, "", event.rawValue);
    const overlaps = sameLeaveType && event.start < previous.end;
    const adjacentSameType = previous && event.start === previous.end
      && preferredLeaveTitle(previous.title, "", previous.rawValue) === preferredLeaveTitle(event.title, "", event.rawValue);
    if (previous && (overlaps || adjacentSameType)) {
      previous.end = event.end > previous.end ? event.end : previous.end;
      previous.rawValue = mergeRawLeaveValues(previous.rawValue, event.rawValue);
      previous.sources = mergeLeaveSources(previous.sources, event.sources, previous.source, event.source);
      previous.title = preferredLeaveTitle(previous.title, event.title, previous.rawValue);
      continue;
    }
    merged.push({
      ...event,
      title: preferredLeaveTitle(event.title, "", event.rawValue),
      sources: mergeLeaveSources(event.sources, null, event.source),
    });
  }

  return [...passthrough, ...merged];
}

function isMergeableLeaveEvent(event) {
  return event?.allDay === true && leaveTextMatches(`${event.title || ""} ${event.rawValue || ""}`);
}

function leaveTextMatches(value) {
  const text = String(value || "");
  // Clinical Support Exam is an all-day professional allocation, not leave.
  if (/\bCS\s+EXAM\b/i.test(text)) return false;
  return /\b(leave|conference|cme|study|annual|sick|personal|exam|sabbatical|parental|long service)\b/i.test(text);
}

function preferredLeaveTitle(leftTitle, rightTitle, rawValue = "") {
  const combined = `${leftTitle || ""} ${rightTitle || ""} ${rawValue || ""}`;
  if (/\bannual\s*&\s*parental\b/i.test(combined)) return "Annual & Parental Leave";
  if (/\b(conference|cme)\b/i.test(combined)) return "Conference Leave";
  if (/\bannual\b/i.test(combined)) return "Annual Leave";
  if (/\b(?:sick|s\/l)\b/i.test(combined)) return "Sick leave";
  if (/\bpersonal\b/i.test(combined)) return "Personal Leave";
  if (/\bstudy\b/i.test(combined)) return "Study Leave";
  if (/\bexam\b/i.test(combined)) return "Exam Leave";
  if (/\b(?:sabbatical|sab\/l)\b/i.test(combined)) return "Sabbatical";
  if (/\bparental\b/i.test(combined)) return "Parental Leave";
  if (/\blong service\b/i.test(combined)) return "Long Service Leave";
  if (/\bcarer'?s?\b/i.test(combined)) return "Carer's Leave";
  if (/\b(?:fam|family)\b/i.test(combined)) return "Family Leave";
  if (/\bmilitary\b/i.test(combined)) return "Military Leave";
  return String(leftTitle || rightTitle || "Leave").trim();
}

function mergeRawLeaveValues(left, right) {
  const values = [left, right]
    .flatMap((item) => String(item || "").split(" / "))
    .map((item) => item.trim())
    .filter(Boolean);
  return [...new Set(values)].join(" / ");
}

function mergeLeaveSources(leftSources, rightSources, leftSource, rightSource) {
  const values = [leftSources, rightSources, leftSource, rightSource]
    .flatMap((item) => Array.isArray(item) ? item : [item])
    .map((item) => String(item || "").trim())
    .filter((item) => /^(MMC|DDH|Casey|MCH)$/i.test(item));
  return [...new Set(values.map((item) => item.toUpperCase() === "CASEY" ? "Casey" : item.toUpperCase()))];
}

function resolveDefaultLocation(source, location, settings) {
  if (!location) return "";
  if (source === "MMC" && location.startsWith("MMC Car Park")) return settings.defaultLocationMmc;
  if (source === "DDH" && location.startsWith("DDH Car Park")) return settings.defaultLocationDdh;
  if (source === "Casey" && location.startsWith("Casey Hospital")) return settings.defaultLocationCasey;
  if (source === "MCH" && location.startsWith("Monash Children's Hospital")) return settings.defaultLocationMch;
  return location;
}

function formatTitle(source, titleParts, settings, kind = "shift") {
  const titleBits = [];
  if (titleParts.base) titleBits.push(titleParts.base);
  if (settings.showAmPm && titleParts.period) titleBits.push(formatTitlePeriod(titleParts.period));
  if (titleParts.suffix) titleBits.push(titleParts.suffix);
  const core = titleBits.join(" ").trim();
  if (!core) return "";
  if ([
    "leave", "annual_leave", "annual_parental_leave", "conference_leave", "sick_leave",
    "sabbatical_leave", "long_service_leave", "parental_leave", "carers_leave",
    "family_leave", "military_leave", "exam_leave", "special_leave",
  ].includes(kind)) {
    return core;
  }
  return settings.showSourcePrefix ? `${source}: ${core}` : core;
}

function formatTitlePeriod(value) {
  return String(value || "").trim().toUpperCase() === "NIGHT" ? "Night" : value;
}

export function customEventsToEvents(customEvents, settings = DEFAULT_SETTINGS, rosterEvents = []) {
  const events = customEvents
    .filter((item) => item.include !== false)
    .map((item) => {
      if (item.allDay) {
        return {
          id: item.id,
          ownerEmail: item.ownerEmail || "",
          source: "Custom",
          title: item.title,
          allDay: true,
          start: item.startDate,
          end: addDays(item.endDate, 1),
          location: settings.includeLocations ? item.location : "",
          rawValue: "Custom event",
          timeLabel: "All day",
          monthKey: item.startDate.slice(0, 7),
        };
      }

      const startHm = item.startTime.split(":").map(Number);
      const endHm = item.endTime.split(":").map(Number);
      const explicitEndDay = item.endDate !== item.startDate ? item.endDate : null;
      const endDate = explicitEndDay || (compareTimes(endHm, startHm) <= 0 ? addDays(item.startDate, 1) : item.startDate);
      return {
        id: item.id,
        ownerEmail: item.ownerEmail || "",
        source: "Custom",
        title: item.title,
        allDay: false,
        start: buildDateTime(item.startDate, startHm),
        end: buildDateTime(endDate, endHm),
        location: settings.includeLocations ? item.location : "",
        rawValue: "Custom event",
        timeLabel: `${item.startTime}-${item.endTime}`,
        monthKey: item.startDate.slice(0, 7),
      };
    });
  return suppressCoveredCustomLeaveEvents(events, rosterEvents);
}

// Manual leave is a fallback for a roster source that has not supplied it.
// When a synced roster later covers every day of that manual multi-day leave,
// prefer the roster data without deleting the custom event. If the source is
// subsequently corrected or removed, the saved manual fallback reappears.
function suppressCoveredCustomLeaveEvents(customEvents, rosterEvents) {
  const rosterLeaveEvents = (rosterEvents || []).filter(isSyncedAllDayLeaveEvent);
  if (!rosterLeaveEvents.length) return customEvents;
  return (customEvents || []).filter((event) => {
    if (!isMultiDayCustomLeaveEvent(event)) return true;
    return !isEveryCustomLeaveDayCovered(event, rosterLeaveEvents);
  });
}

function isMultiDayCustomLeaveEvent(event) {
  if (event?.allDay !== true || String(event?.source || "").toLowerCase() !== "custom") return false;
  if (!isManualLeaveLabel(event.title)) return false;
  const start = String(event.start || "").slice(0, 10);
  const end = String(event.end || "").slice(0, 10);
  return Boolean(start && end && end > addDays(start, 1));
}

function isSyncedAllDayLeaveEvent(event) {
  if (event?.allDay !== true || String(event?.source || "").toLowerCase() === "custom") return false;
  return isManualLeaveLabel(`${event?.title || ""} ${event?.rawValue || ""}`);
}

function isManualLeaveLabel(value) {
  const text = String(value || "").replace(/\s+/g, " ").trim();
  return /\bleave\b/i.test(text)
    || /(?:^|\s)(?:A\/L|AL|C\/L|CL|CME\/L|S\/L|SL|SAB\/L|LSL|PAT\/L|ME\/L|FAM)(?:\s|$)/i.test(text);
}

function isEveryCustomLeaveDayCovered(event, rosterLeaveEvents) {
  const end = String(event.end || "").slice(0, 10);
  let day = String(event.start || "").slice(0, 10);
  while (day && day < end) {
    if (!rosterLeaveEvents.some((rosterEvent) => {
      const start = String(rosterEvent.start || "").slice(0, 10);
      const finish = String(rosterEvent.end || "").slice(0, 10);
      return start <= day && day < finish;
    })) return false;
    day = addDays(day, 1);
  }
  return true;
}

export function applyEventOverrides(events, overrides) {
  const clean = sanitizeOverrides(overrides);
  return events.map((event) => {
    const override = clean[event.id];
    if (!override) return event;
    return {
      ...event,
      title: override.title || event.title,
      start: override.start || event.start,
      end: override.end || event.end,
      allDay: typeof override.allDay === "boolean" ? override.allDay : event.allDay,
      location: Object.prototype.hasOwnProperty.call(override, "location") ? override.location : event.location,
    };
  });
}

function isKindEnabled(kind, settings) {
  if (kind === "annual_leave") return settings.includeAnnualLeave;
  if (kind === "conference_leave") return settings.includeConferenceLeave;
  if (kind === "public_holiday") return settings.includePublicHoliday;
  if (kind === "sick_leave") return settings.includeSickLeave;
  return true;
}

function matchesHospitalFilter(record, settings) {
  if (settings.hospitalFilter === "all") return true;
  return settings.hospitalFilter.toLowerCase() === String(record.source || "").toLowerCase();
}

function matchesDateFilter(record, settings) {
  if (!settings.dateFrom && !settings.dateTo) return true;
  const eventStart = record.startDay;
  const eventEndInclusive = addDays(record.endDay, -1);
  if (settings.dateFrom && eventEndInclusive < settings.dateFrom) return false;
  if (settings.dateTo && eventStart > settings.dateTo) return false;
  return true;
}

function iterateMmcRosterPeople(sheet) {
  const layout = mmcWeekLayout(sheet);
  if (!layout) return [];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const entries = [];
  for (let row = layout.dateRow + 1; row <= range.e.r + 1; row += 1) {
    const marker = layout.markerColumn ? cleanText(getCellValue(sheet, row, layout.markerColumn)) : "";
    if (isMmcStopSection(marker)) break;
    const name = cleanMmcRosterName(getCellValue(sheet, row, layout.nameColumn));
    if (name && looksLikePersonName(name) && !isMmcSectionMarker(name)) {
      entries.push({ row, name });
    }
  }
  return entries;
}

function isDdhSupplementaryRow(sheet, row) {
  if (isDdhDateRow(sheet, row)) return false;
  if (cleanText(getCellValue(sheet, row, 1))) return false;
  for (let col = 2; col <= 8; col += 1) {
    if (cleanText(getCellValue(sheet, row, col))) return true;
  }
  return false;
}

function isMmcSectionMarker(value) {
  const upper = cleanText(value).replace(/\s+/g, " ").trim().toUpperCase();
  return MMC_SECTION_MARKERS.has(upper);
}

function isMmcStopSection(value) {
  const upper = cleanText(value).replace(/\s+/g, " ").trim().toUpperCase();
  return MMC_STOP_SECTIONS.has(upper);
}

function isDdhSectionMarker(value) {
  const upper = cleanText(value).replace(/\s+/g, " ").trim().toUpperCase();
  return DDH_SECTION_MARKERS.has(upper);
}

function cleanMmcRosterName(value) {
  return cleanText(value)
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .replace(/\s+\(?\d+(?:\.\d+)?\s*(?:EFT)?\)?\s*$/i, "")
    .trim();
}

function getCellValue(sheet, row, col) {
  const address = XLSX.utils.encode_cell({ r: row - 1, c: col - 1 });
  return sheet[address]?.v;
}

function isDdhDateRow(sheet, row) {
  const value = getCellValue(sheet, row, 2);
  return typeof value === "string" && WEEKDAY_PREFIXES.some((prefix) => value.startsWith(prefix));
}

function parseDdhDate(value) {
  if (typeof value !== "string") return null;
  const normalized = value.replace("Sept.", "Sep.").replace("June.", "Jun.");
  const date = new Date(normalized);
  return Number.isNaN(date.getTime()) ? null : formatDateOnly(date);
}

function parseDdhTimeRow(value) {
  if (!value) return null;
  const match = value.match(/^\s*(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})\s*$/);
  if (!match) return null;
  return [
    [Number(match[1]), Number(match[2])],
    [Number(match[3]), Number(match[4])],
  ];
}

function extractTimePrefix(value) {
  return extractTimeWithLabel(value, { prefixOnly: true });
}

function extractTimeWithLabel(value, options = {}) {
  const text = String(value || "").trim();
  const match = text.match(/^\s*(\d{1,2}):?(\d{2})-(\d{1,2}):?(\d{2})(?:\s*(.+?))?\s*$/);
  if (match) {
    return {
      start: [Number(match[1]), Number(match[2])],
      end: [Number(match[3]), Number(match[4])],
      label: (match[5] || "").trim(),
    };
  }
  if (options.prefixOnly) return null;
  const suffixMatch = text.match(/^\s*(.+?)\s+(\d{2})(\d{2})-(\d{2})(\d{2})\s*$/);
  if (!suffixMatch) return null;
  return {
    start: [Number(suffixMatch[2]), Number(suffixMatch[3])],
    end: [Number(suffixMatch[4]), Number(suffixMatch[5])],
    label: suffixMatch[1].trim(),
  };
}

function mondayWeeklyLeave(values) {
  const value = values?.[0];
  const leave = normalizeRecognizedLeave(value);
  if (["annual_leave", "annual_parental_leave", "conference_leave", "long_service_leave", "parental_leave"].includes(leave?.kind)) return value;
  if (leave?.kind === "sabbatical_leave" && !/\b(?:AM|PM|NIGHT|NS|SW)\b/i.test(String(value || ""))) return value;
  return null;
}

// A weekly leave marker means the whole roster week only when it is recorded
// on Monday and there is no substantive allocation anywhere in that week.
// A leave entry on any other day remains a single-day event. Some rows
// legitimately combine Monday leave with a later shift (for example Scott's
// CME leave and Friday PCC); do not let the leave marker suppress that shift.
function hasNonLeaveMmcEntry(values) {
  return (values || []).some((value) => {
    const text = cleanText(value);
    return text
      && !normalizeRecognizedLeave(text)
      && !shouldIgnoreMmc(text)
      && !isOtherHospitalReference("MMC", text);
  });
}

function hasNonLeaveDdhEntry(values) {
  return (values || []).some((value) => {
    const text = cleanText(value);
    return text
      && !normalizeRecognizedLeave(text)
      && !shouldIgnoreDdh(text)
      && !isOtherHospitalReference("DDH", text);
  });
}

function shouldIgnoreMmc(value) {
  const upper = value.trim().toUpperCase();
  if (upper.startsWith("DANDENONG")) return true;
  // "Exam" is normally a roster annotation, except for the explicit MMC
  // Clinical Support Exam allocation which is a real, timed shift.
  if (isMmcClinicalSupportExam(upper)) return false;
  return shouldIgnoreCommon(value);
}

function isMmcClinicalSupportExam(value) {
  return /^(?:(?:\d{4})-(?:\d{4})\s+)?CS\s+EXAM$/i.test(String(value || "").trim());
}

// Roster writers use references to another hospital as a safety annotation
// (for example, "PM MMC" on a DDH roster), not as a shift at this facility.
// Keep these labels out of calendars before any generic shift fallback sees them.
function isOtherHospitalReference(source, value) {
  const upper = cleanText(value).replace(/\s+/g, " ").trim().toUpperCase();
  if (/\b(?:TOX|HITH|VHH|ARV|WARRAGUL)\b/.test(upper)) return true;
  if (source !== "MCH" && /\bPAEDS\b/.test(upper)) return true;
  if (source === "DDH" && /\b(?:AED|PED)\b/.test(upper)) return true;
  if ((source === "DDH" || source === "MMC") && /\bCASEY\b/.test(upper)) return true;
  if ((source === "DDH" || source === "Casey") && /\bMMC\b/.test(upper)) return true;
  // In an MMC roster, DH means Dandenong Hospital. These are allocation
  // annotations for the DDH roster, including explicit-time variants.
  if (source === "MMC" && /(?:\bCS\s+(?:DH|DDH)\b|\b(?:DH|DDH)\s+CS\b|\bD\s+C(?:S)?\b|^D\s+\d{4}-\d{4}\s+CS$)/i.test(upper)) return true;
  return false;
}

function shouldIgnoreCommon(value) {
  const upper = value.trim().toUpperCase();
  if (upper === "OTHER" || /^V\s+[SC]$/.test(upper)) return true;
  if (upper.includes("SHIFTS MOVED TO FOLLOWING FORTNIGHT") || upper.includes("SWAPPED FOR SUNDAY NIGHT")) return true;
  if (IGNORED_EXACT.has(upper)) return true;
  return IGNORED_CONTAINS.some((fragment) => upper.includes(fragment));
}

function normalizeSickLeaveLabel(value) {
  const upper = value.trim().toUpperCase();
  const suffix = upper.replace(/^S\/L/, "").trim();
  return `Sick Leave ${suffix}`.trim();
}

function looksLikePersonName(value) {
  const cleaned = String(value).trim();
  if (cleaned.length < 5) return false;
  const upper = cleaned.toUpperCase();
  if (["NOT USED", "SMS", "DATE", "WEEK", "ROLE", "PAGER"].some((token) => upper.includes(token))) {
    return false;
  }
  if (/^\d/.test(cleaned) || /\bVS\b/i.test(cleaned)) return false;
  return /[A-Za-z]/.test(cleaned) && cleaned.includes(" ");
}

function normalizeName(value) {
  return String(value).replace(/[^A-Za-z0-9]+/g, " ").trim().replace(/\s+/g, " ").toUpperCase();
}

function rosterIdentityKey(value) {
  const stripped = String(value || "")
    .replace(/[^A-Za-z0-9,]+/g, " ")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase()
    .replace(/^(DR|DOCTOR|MR|MRS|MS|MISS|PROF|PROFESSOR|A PROF|ASSOC PROF)\s+/, "");
  const parts = stripped.split(/\s*,\s*/).filter(Boolean);
  return parts.length === 2 ? `${parts[1]} ${parts[0]}`.trim() : stripped.replace(/,/g, "");
}

function formatDoctorDisplayName(value) {
  const identity = rosterIdentityKey(value);
  const tokens = identity.split(" ").filter(Boolean);
  if (!tokens.length) return cleanText(value);
  return tokens.map((token, index) => index === tokens.length - 1 ? token : toDisplayNameToken(token)).join(" ");
}

function toDisplayNameToken(value) {
  return value.charAt(0).toUpperCase() + value.slice(1).toLowerCase();
}

function cleanText(value) {
  return value === undefined || value === null ? "" : String(value).trim();
}

function coerceDate(value) {
  if (value instanceof Date) return formatDateOnly(value);
  return null;
}

function formatDateOnly(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function parseDateOnly(value) {
  const [year, month, day] = String(value || "").slice(0, 10).split("-").map(Number);
  return new Date(year, month - 1, day);
}

function australianTermForDate(date) {
  const year = date.getFullYear();
  const candidates = [
    buildAustralianTerm(year, 1, 1),
    buildAustralianTerm(year, 2, 4),
    buildAustralianTerm(year, 3, 7),
    buildAustralianTerm(year, 4, 10),
    buildAustralianTerm(year - 1, 4, 10),
  ];
  return candidates.find((term) => date >= term.start && date < term.end) || buildAustralianTerm(year, 1, 1);
}

function buildAustralianTerm(year, termNumber, startMonthIndex) {
  const start = firstMondayOfMonth(year, startMonthIndex);
  const end = new Date(start);
  end.setDate(end.getDate() + 91);
  return { year, termNumber, start, end };
}

function firstMondayOfMonth(year, monthIndex) {
  const date = new Date(year, monthIndex, 1);
  const day = date.getDay();
  const delta = day === 0 ? 1 : day === 1 ? 0 : 8 - day;
  date.setDate(date.getDate() + delta);
  return date;
}

function formatAustralianTermLabel(term) {
  return `Term ${term.termNumber} ${term.year}`;
}

function buildDateTime(day, hm, plusDay = false) {
  const actualDay = plusDay ? addDays(day, 1) : day;
  const offset = offsetSuffixForDay(actualDay);
  return `${actualDay}T${String(hm[0]).padStart(2, "0")}:${String(hm[1]).padStart(2, "0")}:00${offset}`;
}

function offsetSuffixForDay(day) {
  const probe = new Date(`${day}T12:00:00Z`);
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: TIMEZONE,
    timeZoneName: "longOffset",
  }).formatToParts(probe);
  const token = parts.find((part) => part.type === "timeZoneName")?.value || "GMT+10:00";
  return token.replace("GMT", "");
}

function addDays(day, amount) {
  const date = typeof day === "string" ? new Date(`${day}T00:00:00`) : new Date(day);
  date.setDate(date.getDate() + amount);
  return formatDateOnly(date);
}

function mondayOfDay(day) {
  const date = new Date(`${day}T00:00:00`);
  const weekday = date.getDay();
  const delta = weekday === 0 ? -6 : 1 - weekday;
  date.setDate(date.getDate() + delta);
  return formatDateOnly(date);
}

function compareTimes(left, right) {
  return left[0] * 60 + left[1] - (right[0] * 60 + right[1]);
}

function asDateString(value) {
  return value.slice(0, 10);
}

function formatTimeLabel(start, end) {
  return `${extractClock(start)}-${extractClock(end)}`;
}

function extractClock(value) {
  const match = String(value).match(/T(\d{2}:\d{2})/);
  return match ? match[1] : "";
}

function toTitleCase(value) {
  return value
    .toLowerCase()
    .split(/\s+/)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(" ");
}

function toIcsDateTime(value) {
  return value.replace(/[-:]/g, "").replace(/\+(\d{2})(\d{2})$/, "");
}

function escapeIcsText(value) {
  return value.replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\n/g, "\\n");
}

function hashString(value) {
  let hash = 5381;
  for (let index = 0; index < value.length; index += 1) {
    hash = ((hash << 5) + hash) ^ value.charCodeAt(index);
  }
  return (hash >>> 0).toString(16);
}

function isDateString(value) {
  return typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value);
}

function isClockString(value) {
  return typeof value === "string" && /^\d{2}:\d{2}$/.test(value);
}
