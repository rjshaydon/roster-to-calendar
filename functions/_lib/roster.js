import * as XLSX from "xlsx";

const TIMEZONE = "Australia/Melbourne";

const MMC_LOCATION = "MMC Car Park, Tarella Road, Clayton VIC 3168, Australia";
const DDH_LOCATION = "DDH Car Park, 135 David St, Dandenong VIC 3175, Australia";

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
  "FAST PM",
  "AVAO AM",
  "AVAO PM",
  "PHNW",
]);

const WEEKDAY_PREFIXES = ["Mon.", "Tue.", "Wed.", "Thu.", "Fri.", "Sat.", "Sun."];
const WEEKLY_LEAVE_LABELS = new Set(["ANNUAL LEAVE", "CONFERENCE LEAVE"]);
const IGNORED_EXACT = new Set([
  "",
  "AL",
  "A/L",
  "EXAM",
  "EXAM LEAVE",
  "CME LEAVE",
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
const MMC_SECTION_BREAKS = new Set([
  "GERIATRICIAN",
  "CMO",
  "SENIOR REG",
  "INTERMEDIATE REG",
  "JUNIOR REG",
  "HMO",
]);

const DEFAULT_SETTINGS = {
  showSourcePrefix: true,
  showAmPm: true,
  showTimes: true,
  showLocations: false,
  showRawValues: false,
  showNormalizedTitles: true,
  includeLocations: true,
  includeAnnualLeave: true,
  includeConferenceLeave: true,
  includePublicHoliday: true,
  includeSickLeave: true,
  hospitalFilter: "all",
  dateFrom: "",
  dateTo: "",
};

export function defaultSettings() {
  return { ...DEFAULT_SETTINGS };
}

export async function parseUploadForm(request) {
  const formData = await request.formData();
  const uploads = formData.getAll("rosterFiles").filter((item) => item instanceof File);
  if (!uploads.length) {
    throw new Error("Upload at least one roster file.");
  }
  if (uploads.length > 2) {
    throw new Error("Upload no more than two roster files at a time.");
  }

  const sources = { mmc: null, ddh: null };
  for (const file of uploads) {
    const workbook = await readWorkbook(file);
    const sourceType = detectSourceType(workbook, file.name);
    if (sources[sourceType]) {
      throw new Error(`Upload at most one ${sourceType.toUpperCase()} roster at a time.`);
    }
    sources[sourceType] = { file, workbook };
  }

  if (!sources.mmc && !sources.ddh) {
    throw new Error("Upload at least one MMC or DDH roster.");
  }

  return {
    sources,
    doctorKey: String(formData.get("doctorKey") || ""),
    doctorDisplay: String(formData.get("doctorDisplay") || ""),
    settings: sanitizeSettings(parseJsonField(formData, "settings", DEFAULT_SETTINGS)),
    overrides: sanitizeOverrides(parseJsonField(formData, "overrides", {})),
    customEvents: sanitizeCustomEvents(parseJsonField(formData, "customEvents", [])),
  };
}

export function doctorOptions(mmcWorkbook, ddhWorkbook) {
  if (!mmcWorkbook && !ddhWorkbook) {
    throw new Error("Upload at least one MMC or DDH roster.");
  }
  const mmcNames = mmcWorkbook ? extractMmcNames(mmcWorkbook) : new Map();
  const ddhNames = ddhWorkbook ? extractDdhNames(ddhWorkbook) : new Map();
  const keys = mmcWorkbook && ddhWorkbook
    ? [...mmcNames.keys()].filter((key) => ddhNames.has(key)).sort()
    : [...(mmcWorkbook ? mmcNames.keys() : ddhNames.keys())].sort();

  return keys.map((key) => ({
    key,
    displayName: mmcNames.get(key) || ddhNames.get(key),
  }));
}

export function buildRosterView(mmcWorkbook, ddhWorkbook, doctorKey, settings = DEFAULT_SETTINGS, overrides = {}) {
  const records = [];
  if (mmcWorkbook) records.push(...parseMmcRecords(mmcWorkbook, doctorKey));
  if (ddhWorkbook) records.push(...parseDdhRecords(ddhWorkbook, doctorKey));

  records.sort((left, right) => {
    if (left.startDay !== right.startDay) return left.startDay.localeCompare(right.startDay);
    if (left.source !== right.source) return left.source.localeCompare(right.source);
    return left.rawValue.localeCompare(right.rawValue);
  });

  return applySettings(records, sanitizeSettings(settings), sanitizeOverrides(overrides));
}

export function generateEvents(mmcWorkbook, ddhWorkbook, doctorKey, settings = DEFAULT_SETTINGS, overrides = {}) {
  return buildRosterView(mmcWorkbook, ddhWorkbook, doctorKey, settings, overrides).events;
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

export function serializeReviewItem(item) {
  return {
    id: item.id,
    source: item.source,
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
    mmc: sources.mmc?.file.name || "",
    ddh: sources.ddh?.file.name || "",
  };
}

async function readWorkbook(file) {
  try {
    const bytes = await file.arrayBuffer();
    return XLSX.read(bytes, { type: "array", cellDates: true });
  } catch {
    throw new Error(`${file.name} is not a supported MMC workbook or Dandenong Hospital FindMyShift export.`);
  }
}

function detectSourceType(workbook, filename) {
  const sheetNames = workbook.SheetNames || [];
  if (sheetNames.includes("Whole thing") && sheetNames.some((name) => name.startsWith("Week "))) {
    return "mmc";
  }
  const sheet = workbook.Sheets[sheetNames[0]];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  if (range.e.c + 1 >= 8 && [1, 2, 3, 4].some((row) => isDdhDateRow(sheet, row))) {
    return "ddh";
  }
  throw new Error(`${filename} is not a supported MMC workbook or Dandenong Hospital FindMyShift export.`);
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
    showLocations: input.showLocations === true,
    showRawValues: input.showRawValues === true,
    showNormalizedTitles: input.showNormalizedTitles !== false,
    includeLocations: input.includeLocations !== false,
    includeAnnualLeave: input.includeAnnualLeave !== false,
    includeConferenceLeave: input.includeConferenceLeave !== false,
    includePublicHoliday: input.includePublicHoliday !== false,
    includeSickLeave: input.includeSickLeave !== false,
    hospitalFilter: input.hospitalFilter === "mmc" || input.hospitalFilter === "ddh" ? input.hospitalFilter : "all",
    dateFrom: isDateString(input.dateFrom) ? input.dateFrom : "",
    dateTo: isDateString(input.dateTo) ? input.dateTo : "",
  };
}

function sanitizeOverrides(raw) {
  const overrides = {};
  if (!raw || typeof raw !== "object") return overrides;
  for (const [key, value] of Object.entries(raw)) {
    if (!value || typeof value !== "object") continue;
    overrides[key] = {
      title: typeof value.title === "string" ? value.title.trim() : "",
      include: typeof value.include === "boolean" ? value.include : undefined,
      start: typeof value.start === "string" ? value.start.trim() : "",
      end: typeof value.end === "string" ? value.end.trim() : "",
      location: typeof value.location === "string" ? value.location.trim() : "",
      allDay: typeof value.allDay === "boolean" ? value.allDay : undefined,
    };
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
  return events;
}

function extractMmcNames(workbook) {
  const names = new Map();
  for (const sheetName of workbook.SheetNames) {
    if (!sheetName.startsWith("Week ")) continue;
    const sheet = workbook.Sheets[sheetName];
    for (const { name } of iterateMmcConsultantNames(sheet)) {
      if (!looksLikePersonName(name)) continue;
      const key = normalizeName(name);
      if (!names.has(key)) names.set(key, String(name).trim());
    }
  }
  return names;
}

function extractDdhNames(workbook) {
  const names = new Map();
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  for (let row = 1; row <= range.e.r + 1; row += 1) {
    const value = cleanText(getCellValue(sheet, row, 1));
    if (!value || isDdhDateRow(sheet, row) || !looksLikePersonName(value)) continue;
    const key = normalizeName(value);
    if (!names.has(key)) names.set(key, value);
  }
  return names;
}

function parseMmcRecords(workbook, doctorKey) {
  const records = [];
  for (const sheetName of workbook.SheetNames) {
    if (!sheetName.startsWith("Week ")) continue;
    const sheet = workbook.Sheets[sheetName];
    const weekDates = [];
    for (let col = 6; col <= 12; col += 1) {
      const value = coerceDate(getCellValue(sheet, 4, col));
      if (!value) {
        weekDates.length = 0;
        break;
      }
      weekDates.push(value);
    }
    if (!weekDates.length) continue;

    for (const { row, name } of iterateMmcConsultantNames(sheet)) {
      if (normalizeName(name) !== doctorKey) continue;
      const weekValues = [];
      for (let col = 6; col <= 12; col += 1) {
        weekValues.push(cleanText(getCellValue(sheet, row, col)));
      }
      const weeklyLeave = firstWeeklyLeave(weekValues);
      if (weeklyLeave) {
        records.push(createWeeklyLeaveRecord("MMC", weekDates[0], weeklyLeave));
      } else {
        weekValues.forEach((raw, index) => {
          if (!raw) return;
          const record = parseMmcEntry(weekDates[index], raw);
          if (record) records.push(record);
        });
      }
      break;
    }
  }
  return records;
}

function parseDdhRecords(workbook, doctorKey) {
  const records = [];
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  let row = 1;
  while (row <= range.e.r + 1) {
    if (!isDdhDateRow(sheet, row)) {
      row += 1;
      continue;
    }
    const weekDates = [];
    for (let col = 2; col <= 8; col += 1) {
      const value = parseDdhDate(getCellValue(sheet, row, col));
      if (!value) {
        weekDates.length = 0;
        break;
      }
      weekDates.push(value);
    }
    if (!weekDates.length) {
      row += 1;
      continue;
    }

    const nameRow = row + 1;
    const rawName = cleanText(getCellValue(sheet, nameRow, 1));
    const labels = [];
    const times = [];
    for (let col = 2; col <= 8; col += 1) {
      labels.push(cleanText(getCellValue(sheet, nameRow, col)));
      times.push(cleanText(getCellValue(sheet, nameRow + 1, col)));
    }
    const hasTimeRow = nameRow + 1 <= range.e.r + 1 && !isDdhDateRow(sheet, nameRow + 1) && times.some(Boolean);

    if (normalizeName(rawName) === doctorKey) {
      const weeklyLeave = firstWeeklyLeave(labels);
      if (weeklyLeave) {
        records.push(createWeeklyLeaveRecord("DDH", weekDates[0], weeklyLeave));
      } else {
        weekDates.forEach((day, index) => {
          const record = parseDdhEntry(day, labels[index], hasTimeRow ? times[index] : "");
          if (record) records.push(record);
        });
      }
    }
    row += hasTimeRow ? 3 : 2;
  }
  return records;
}

function parseMmcEntry(day, raw) {
  const upper = raw.toUpperCase();
  if (shouldIgnoreMmc(raw)) return null;
  if (upper === "PHNW") {
    return createAllDayRecord("MMC", day, raw, {
      kind: "public_holiday",
      titleParts: { base: "PHNW", period: "", suffix: "" },
      location: "",
    });
  }
  if (upper.startsWith("S/L")) {
    return createAllDayRecord("MMC", day, raw, {
      kind: "sick_leave",
      titleParts: { base: normalizeSickLeaveLabel(raw), period: "", suffix: "" },
      location: "",
    });
  }

  const explicit = extractTimePrefix(raw);
  const label = explicit ? explicit.label : raw.trim();
  const normalized = normalizeMmcLabel(label);
  if (!normalized) {
    return createUnknownRecord("MMC", day, raw, "MMC shift code not recognised.");
  }

  if (explicit) {
    return createTimedRecord("MMC", day, raw, {
      kind: normalized.kind,
      titleParts: normalized.titleParts,
      startHm: explicit.start,
      endHm: explicit.end,
      location: normalized.location || "",
    });
  }
  if (normalized.allDay) {
    return createAllDayRecord("MMC", day, raw, {
      kind: normalized.kind,
      titleParts: normalized.titleParts,
      location: normalized.location || "",
    });
  }
  return createTimedRecord("MMC", day, raw, {
    kind: normalized.kind,
    titleParts: normalized.titleParts,
    startHm: normalized.defaultTimes[0],
    endHm: normalized.defaultTimes[1],
    location: normalized.location || "",
  });
}

function parseDdhEntry(day, label, timeText) {
  if (!label) return null;
  if (parseDdhTimeRow(label)) {
    return timeText ? parseDdhEntry(day, timeText, label) : null;
  }
  const upper = label.toUpperCase();
  if (upper === "AM" || upper === "PM") return null;
  if (upper === "PHNW" || upper === "PHNW CLINICAL") {
    return createAllDayRecord("DDH", day, label, {
      kind: "public_holiday",
      titleParts: { base: "PHNW", period: "", suffix: "" },
      location: "",
    });
  }
  if (upper.startsWith("S/L")) {
    return createAllDayRecord("DDH", day, label, {
      kind: "sick_leave",
      titleParts: { base: normalizeSickLeaveLabel(label), period: "", suffix: "" },
      location: "",
    });
  }
  if (shouldIgnoreCommon(label)) return null;

  const mapped = DDH_LABEL_MAP[label] || label;
  const location = mapped === "CS" || mapped === "PHNW" ? "" : DDH_LOCATION;
  const normalized = normalizeDdhLabel(mapped);
  if (!normalized) {
    return createUnknownRecord("DDH", day, label, "DDH shift label not recognised.");
  }

  const parsedTime = parseDdhTimeRow(timeText);
  if (parsedTime) {
    return createTimedRecord("DDH", day, label, {
      kind: normalized.kind,
      titleParts: normalized.titleParts,
      startHm: parsedTime[0],
      endHm: parsedTime[1],
      location,
    });
  }

  return createAllDayRecord("DDH", day, label, {
    kind: normalized.kind,
    titleParts: normalized.titleParts,
    location,
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

function createTimedRecord(source, day, rawValue, details) {
  const start = buildDateTime(day, details.startHm);
  const plusDay = compareTimes(details.endHm, details.startHm) <= 0;
  const end = buildDateTime(day, details.endHm, plusDay);
  const normalizedTitle = formatTitle(source, details.titleParts, { ...DEFAULT_SETTINGS, showTimes: false, showRawValues: false, showLocations: false }, details.kind);
  return {
    id: hashString(`${source}|${day}|${rawValue}|${normalizedTitle}|${start}|${end}`),
    source,
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
    status: details.ambiguous ? "ambiguous" : "ok",
    warnings: details.warning ? [details.warning] : [],
    exportable: true,
    includeByDefault: true,
  };
}

function createAllDayRecord(source, day, rawValue, details) {
  const normalizedTitle = formatTitle(source, details.titleParts, { ...DEFAULT_SETTINGS, showTimes: false, showRawValues: false, showLocations: false }, details.kind);
  return {
    id: hashString(`${source}|${day}|${rawValue}|${normalizedTitle}|all-day`),
    source,
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
    status: details.ambiguous ? "ambiguous" : "ok",
    warnings: details.warning ? [details.warning] : [],
    exportable: true,
    includeByDefault: true,
  };
}

function createWeeklyLeaveRecord(source, monday, rawValue) {
  const label = toTitleCase(rawValue);
  const kind = rawValue.toUpperCase() === "CONFERENCE LEAVE" ? "conference_leave" : "annual_leave";
  const normalizedTitle = label;
  return {
    id: hashString(`${source}|${monday}|${rawValue}|week-leave`),
    source,
    kind,
    rawValue,
    startDay: monday,
    endDay: addDays(monday, 7),
    allDay: true,
    start: monday,
    end: addDays(monday, 7),
    location: "",
    titleParts: { base: label, period: "", suffix: "" },
    normalizedTitle,
    status: "ok",
    warnings: [],
    exportable: true,
    includeByDefault: true,
  };
}

function createUnknownRecord(source, day, rawValue, warning) {
  return {
    id: hashString(`${source}|${day}|${rawValue}|unknown`),
    source,
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

function applySettings(records, settings, overrides) {
  const scopedRecords = records.filter((record) => matchesHospitalFilter(record, settings) && matchesDateFilter(record, settings));
  const events = [];
  const reviewItems = [];
  const issues = [];

  for (const record of scopedRecords) {
    const override = overrides[record.id] || {};
    const defaultInclude = record.includeByDefault && isKindEnabled(record.kind, settings);
    const include = typeof override.include === "boolean" ? override.include : defaultInclude;
    const suggestedTitle = formatTitle(record.source, record.titleParts, settings, record.kind);
    const overrideTitle = override.title || "";
    const finalTitle = overrideTitle || suggestedTitle;
    const timeLabel = record.allDay ? "All day" : formatTimeLabel(record.start, record.end);
    const location = settings.includeLocations ? record.location : "";

    reviewItems.push({
      id: record.id,
      source: record.source,
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
        startDay: record.startDay,
        rawValue: record.rawValue,
        status: record.status,
        message: record.warnings[0] || "Review this roster entry before export.",
      });
    }

    if (!include || !record.exportable || !finalTitle) continue;

    events.push({
      id: record.id,
      source: record.source,
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

function formatTitle(source, titleParts, settings, kind = "shift") {
  const titleBits = [];
  if (titleParts.base) titleBits.push(titleParts.base);
  if (settings.showAmPm && titleParts.period) titleBits.push(titleParts.period);
  if (titleParts.suffix) titleBits.push(titleParts.suffix);
  const core = titleBits.join(" ").trim();
  if (!core) return "";
  if (kind === "annual_leave" || kind === "conference_leave") {
    return core;
  }
  return settings.showSourcePrefix ? `${source}: ${core}` : core;
}

export function customEventsToEvents(customEvents, settings = DEFAULT_SETTINGS) {
  return customEvents
    .filter((item) => item.include !== false)
    .map((item) => {
      if (item.allDay) {
        return {
          id: item.id,
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
      location: override.location || event.location,
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
  return settings.hospitalFilter.toUpperCase() === record.source;
}

function matchesDateFilter(record, settings) {
  if (!settings.dateFrom && !settings.dateTo) return true;
  const eventStart = record.startDay;
  const eventEndInclusive = addDays(record.endDay, -1);
  if (settings.dateFrom && eventEndInclusive < settings.dateFrom) return false;
  if (settings.dateTo && eventStart > settings.dateTo) return false;
  return true;
}

function iterateMmcConsultantNames(sheet) {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const entries = [];
  for (let row = 7; row <= range.e.r + 1; row += 1) {
    const marker = cleanText(getCellValue(sheet, row, 3)).toUpperCase();
    if (MMC_SECTION_BREAKS.has(marker)) break;
    const name = getCellValue(sheet, row, 4);
    if (name !== undefined && name !== null && String(name).trim()) {
      entries.push({ row, name: String(name).trim() });
    }
  }
  return entries;
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
  const match = value.match(/^\s*(\d{2})(\d{2})-(\d{2})(\d{2})\s+(.+?)\s*$/);
  if (!match) return null;
  return {
    start: [Number(match[1]), Number(match[2])],
    end: [Number(match[3]), Number(match[4])],
    label: match[5].trim(),
  };
}

function firstWeeklyLeave(values) {
  for (const value of values) {
    if (WEEKLY_LEAVE_LABELS.has(value.toUpperCase())) return value;
  }
  return null;
}

function shouldIgnoreMmc(value) {
  const upper = value.trim().toUpperCase();
  if (upper.startsWith("DANDENONG")) return true;
  return shouldIgnoreCommon(value);
}

function shouldIgnoreCommon(value) {
  const upper = value.trim().toUpperCase();
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
  return /[A-Za-z]/.test(cleaned) && cleaned.includes(" ");
}

function normalizeName(value) {
  return String(value).replace(/[^A-Za-z0-9]+/g, " ").trim().replace(/\s+/g, " ").toUpperCase();
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
