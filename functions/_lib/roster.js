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

export function generateEvents(mmcWorkbook, ddhWorkbook, doctorKey) {
  const events = [];
  if (mmcWorkbook) events.push(...parseMmcEvents(mmcWorkbook, doctorKey));
  if (ddhWorkbook) events.push(...parseDdhEvents(ddhWorkbook, doctorKey));
  return events.sort((left, right) => {
    const leftDate = asDateString(left.start);
    const rightDate = asDateString(right.start);
    if (leftDate !== rightDate) return leftDate.localeCompare(rightDate);
    if (left.allDay !== right.allDay) return left.allDay ? -1 : 1;
    return left.title.localeCompare(right.title);
  });
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
    source: event.source,
    title: event.title,
    allDay: event.allDay,
    start: event.start,
    end: event.end,
    location: event.location || "",
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

function parseMmcEvents(workbook, doctorKey) {
  const events = [];
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
        const monday = weekDates[0];
        events.push({
          source: "MMC",
          title: `MMC: ${toTitleCase(weeklyLeave)}`,
          start: monday,
          end: addDays(monday, 7),
          allDay: true,
          location: "",
        });
      } else {
        weekValues.forEach((raw, index) => {
          if (!raw) return;
          const event = parseMmcEntry(weekDates[index], raw);
          if (event) events.push(event);
        });
      }
      break;
    }
  }
  return events;
}

function parseDdhEvents(workbook, doctorKey) {
  const events = [];
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
        const monday = weekDates[0];
        events.push({
          source: "DDH",
          title: `DDH: ${toTitleCase(weeklyLeave)}`,
          start: monday,
          end: addDays(monday, 7),
          allDay: true,
          location: "",
        });
      } else {
        weekDates.forEach((day, index) => {
          const event = parseDdhEntry(day, labels[index], hasTimeRow ? times[index] : "");
          if (event) events.push(event);
        });
      }
    }
    row += hasTimeRow ? 3 : 2;
  }
  return events;
}

function parseMmcEntry(day, raw) {
  const upper = raw.toUpperCase();
  if (shouldIgnoreMmc(raw)) return null;
  if (upper === "PHNW") {
    return { source: "MMC", title: "MMC: PHNW", start: day, end: addDays(day, 1), allDay: true, location: "" };
  }
  if (upper.startsWith("S/L")) {
    return { source: "MMC", title: `MMC: ${normalizeSickLeaveLabel(raw)}`, start: day, end: addDays(day, 1), allDay: true, location: "" };
  }

  const explicit = extractTimePrefix(raw);
  const label = explicit ? explicit.label : raw.trim();
  const normalized = normalizeMmcLabel(label);
  if (!normalized) return null;

  if (explicit) {
    return {
      source: "MMC",
      title: `MMC: ${normalized.title}`,
      start: buildDateTime(day, explicit.start),
      end: buildDateTime(day, explicit.end, compareTimes(explicit.end, explicit.start) <= 0),
      allDay: false,
      location: normalized.location || "",
    };
  }
  if (normalized.allDay) {
    return {
      source: "MMC",
      title: `MMC: ${normalized.title}`,
      start: day,
      end: addDays(day, 1),
      allDay: true,
      location: normalized.location || "",
    };
  }
  return {
    source: "MMC",
    title: `MMC: ${normalized.title}`,
    start: buildDateTime(day, normalized.defaultTimes[0]),
    end: buildDateTime(day, normalized.defaultTimes[1], compareTimes(normalized.defaultTimes[1], normalized.defaultTimes[0]) <= 0),
    allDay: false,
    location: normalized.location || "",
  };
}

function parseDdhEntry(day, label, timeText) {
  if (!label) return null;
  if (parseDdhTimeRow(label)) {
    return timeText ? parseDdhEntry(day, timeText, label) : null;
  }
  const upper = label.toUpperCase();
  if (upper === "AM" || upper === "PM") return null;
  if (upper === "PHNW" || upper === "PHNW CLINICAL") {
    return { source: "DDH", title: "DDH: PHNW", start: day, end: addDays(day, 1), allDay: true, location: "" };
  }
  if (upper.startsWith("S/L")) {
    return { source: "DDH", title: `DDH: ${normalizeSickLeaveLabel(label)}`, start: day, end: addDays(day, 1), allDay: true, location: "" };
  }
  if (shouldIgnoreCommon(label)) return null;

  const mapped = DDH_LABEL_MAP[label] || label;
  const location = mapped === "CS" || mapped === "PHNW" ? "" : DDH_LOCATION;
  const parsedTime = parseDdhTimeRow(timeText);
  if (parsedTime) {
    return {
      source: "DDH",
      title: `DDH: ${mapped}`,
      start: buildDateTime(day, parsedTime[0]),
      end: buildDateTime(day, parsedTime[1], compareTimes(parsedTime[1], parsedTime[0]) <= 0),
      allDay: false,
      location,
    };
  }
  return {
    source: "DDH",
    title: `DDH: ${mapped}`,
    start: day,
    end: addDays(day, 1),
    allDay: true,
    location,
  };
}

function normalizeMmcLabel(label) {
  const code = label.trim().toUpperCase();
  if (code === "CS") return { title: "CS", location: "", allDay: true, defaultTimes: null };
  if (code === "CSO") return { title: "CSO", location: MMC_LOCATION, allDay: true, defaultTimes: null };

  if (code.length === 4 && code[1] === "S" && code[2] === "S" && ["A", "P"].includes(code[0]) && ["C", "R"].includes(code[3])) {
    const shift = code[0] === "A" ? "AM" : "PM";
    const floatSuffix = code[3] === "R" ? " Float" : "";
    return {
      title: `SSU ${shift}${floatSuffix}`,
      location: MMC_LOCATION,
      allDay: false,
      defaultTimes: code[0] === "A" ? [[7, 30], [17, 30]] : [[14, 30], [0, 0]],
    };
  }

  if (code.length === 3 && ["A", "P"].includes(code[0]) && MMC_TEAM_MAP[code[1]] && ["C", "R"].includes(code[2])) {
    const shift = code[0] === "A" ? "AM" : "PM";
    const floatSuffix = code[2] === "R" ? " Float" : "";
    return {
      title: `${MMC_TEAM_MAP[code[1]]} ${shift}${floatSuffix}`,
      location: MMC_LOCATION,
      allDay: false,
      defaultTimes: code[0] === "A" ? [[8, 0], [17, 30]] : [[14, 30], [0, 0]],
    };
  }
  return null;
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
