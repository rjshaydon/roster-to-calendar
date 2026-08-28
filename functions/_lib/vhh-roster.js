export const VHH_ROSTER_SOURCE_ID = "vhh-active-medical-roster";

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;
const MAX_BLOCKS = 24;
const MAX_ROWS_PER_BLOCK = 80;
const MAX_ASSIGNMENTS_PER_ROW = 14;
const VHH_LOCATION = "Victorian Heart Hospital, 631 Blackburn Road, Clayton VIC 3168, Australia";
const COMMA_NAME = /([A-Za-z][A-Za-z'’.-]*(?:[ \t]+[A-Za-z][A-Za-z'’.-]*)*),[ \t]*([A-Za-z][A-Za-z'’.-]*(?:[ \t]+[A-Za-z][A-Za-z'’.-]*)*)/g;

const SHIFT_DEFINITIONS = new Map([
  ["CST", shift("Clinical Support", "SMS", "", "", false)],
  ["T", shift("Clinical Support", "SMS", "", "", false)],
  ["AM SMS", shift("AM SMS", "SMS", "08:00", "17:30")],
  ["AM REG", shift("AM Reg", "Junior Registrar", "08:00", "17:30")],
  ["SSU HMO (8-4)", shift("SSU HMO", "HMO", "08:00", "16:00")],
  ["AM JMS", shift("AM JMS", "Junior Registrar", "08:00", "17:30")],
  ["SWING", shift("Swing SMS 1000", "SMS", "10:00", "19:30")],
  ["SWING 1230PM", shift("Swing SMS 1230", "SMS", "12:30", "22:00")],
  ["PM SMS", shift("PM SMS", "SMS", "14:30", "00:00")],
  ["PM REG", shift("PM Reg", "Junior Registrar", "14:30", "00:00")],
  ["PM JMS", shift("PM JMS", "Junior Registrar", "14:30", "00:00")],
  ["ON REG", shift("Night Reg", "Junior Registrar", "23:00", "08:30")],
  ["ON JMS", shift("Night JMS", "Junior Registrar", "23:00", "08:30")],
]);
const IGNORED_SHIFT_LABELS = new Set(["MED STUDENT"]);

// SharePoint remains authoritative for the workbook. The background worker
// extracts the roster structure, then this adapter applies the VHH-specific
// designation, timing, location and term-wide seniority rules.
export function normaliseVhhRosterExtract(payload) {
  if (!payload || typeof payload !== "object" || String(payload.sourceId || "").trim() !== VHH_ROSTER_SOURCE_ID) return null;
  const blocks = Array.isArray(payload.blocks) ? payload.blocks : [];
  if (!blocks.length || blocks.length > MAX_BLOCKS) return null;
  const normalisedBlocks = blocks.map(normaliseBlock).filter(Boolean);
  if (!normalisedBlocks.length) return null;
  return {
    schemaVersion: 1,
    sourceId: VHH_ROSTER_SOURCE_ID,
    fileName: safeRosterFileName(payload.fileName),
    providerModifiedAt: String(payload.providerModifiedAt || "").trim(),
    providerVersion: String(payload.providerVersion || "").trim(),
    blocks: normalisedBlocks,
  };
}

export function buildVhhDerivedRosterPayload({ extract, contentHash, fileId = "", providerVersion = "", fileSize = 0, lastModified = 0 } = {}) {
  const roster = normaliseVhhRosterExtract(extract);
  if (!roster) throw new Error("VHH roster JSON is invalid.");
  const resolvedFileId = String(fileId || `automation:${VHH_ROSTER_SOURCE_ID}:${String(contentHash || "").slice(0, 24)}`);
  const addedAt = new Date().toISOString();
  const doctorsByKey = new Map();
  const eventsByDoctor = {};
  const termSeniorities = buildTermSeniorities(roster.blocks);
  const unknownLabels = new Set();

  for (const block of roster.blocks) {
    for (const row of block.rows) {
      const definition = shiftDefinition(row.shiftLabel);
      if (!definition) {
        if (!isIgnoredShift(row.shiftLabel) && row.assignments.some((assignment) => vhhPeopleFromCell(assignment.namesText).length)) {
          unknownLabels.add(row.shiftLabel);
        }
        continue;
      }
      if (!block.visible) continue;
      for (const assignment of row.assignments) {
        const people = vhhPeopleFromCell(assignment.namesText);
        for (const person of people) {
          const seniority = termSeniorities.get(person.key) || definition.seniority;
          if (!doctorsByKey.has(person.key)) {
            doctorsByKey.set(person.key, {
              key: person.key,
              displayName: person.displayName,
              sourceType: "vhh",
              seniority,
              membershipSource: "roster",
            });
          }
          const events = eventsByDoctor[person.key] || (eventsByDoctor[person.key] = []);
          events.push(vhhEvent({ person, block, row, assignment, definition, seniority }));
        }
      }
    }
  }

  if (unknownLabels.size) {
    throw new Error(`Unsupported VHH shift designation${unknownLabels.size === 1 ? "" : "s"}: ${[...unknownLabels].sort().join(", ")}.`);
  }

  const doctors = [...doctorsByKey.values()].sort((left, right) => left.displayName.localeCompare(right.displayName));
  const eventCount = Object.values(eventsByDoctor).reduce((total, events) => total + events.length, 0);
  if (!doctors.length || !eventCount) throw new Error("VHH roster JSON contained no rostered staff.");
  return {
    file: {
      id: resolvedFileId,
      name: roster.fileName,
      sourceType: "vhh",
      sourceId: VHH_ROSTER_SOURCE_ID,
      size: Number(fileSize) > 0 ? Number(fileSize) : new TextEncoder().encode(JSON.stringify(roster)).byteLength,
      lastModified: Number(lastModified) > 0 ? Number(lastModified) : Date.parse(roster.providerModifiedAt) || Date.now(),
      addedAt,
      uploadedAt: addedAt,
      uploadedBy: `automation:${VHH_ROSTER_SOURCE_ID}`,
      providerVersion: String(providerVersion || roster.providerVersion || ""),
    },
    doctors,
    eventsByDoctor,
    issuesByDoctor: Object.fromEntries(doctors.map((doctor) => [doctor.key, []])),
    eventCount,
  };
}

function safeRosterFileName(value) {
  const name = String(value || "").trim().split(/[\\/]/).pop();
  return name && /\.(?:xlsx|json)$/i.test(name) ? name : "Active Medical Roster.json";
}

function normaliseBlock(value) {
  if (!value || typeof value !== "object") return null;
  const dates = (Array.isArray(value.dates) ? value.dates : [])
    .map((date) => ({
      date: String(date?.date || "").slice(0, 10),
      displayedDate: String(date?.displayedDate || "").trim(),
      sourceColumn: String(date?.sourceColumn || "").trim(),
    }))
    .filter((date) => ISO_DATE.test(date.date));
  if (!dates.length || dates.length > 14) return null;
  const allowedDates = new Set(dates.map((date) => date.date));
  const rows = (Array.isArray(value.rows) ? value.rows : []).slice(0, MAX_ROWS_PER_BLOCK).map((row) => {
    const shiftLabel = text(row?.shiftLabel);
    if (!shiftLabel || /jms\s+teaching\s+timetable/i.test(shiftLabel)) return null;
    const assignments = (Array.isArray(row?.assignments) ? row.assignments : []).slice(0, MAX_ASSIGNMENTS_PER_ROW)
      .map((assignment) => ({
        date: String(assignment?.date || "").slice(0, 10),
        displayedDate: text(assignment?.displayedDate),
        namesText: text(assignment?.namesText),
        sourceCell: text(assignment?.sourceCell),
      }))
      .filter((assignment) => allowedDates.has(assignment.date) && assignment.namesText);
    return assignments.length ? {
      sourceRow: Number(row?.sourceRow || 0),
      sourceShiftLabel: text(row?.sourceShiftLabel),
      shiftLabel,
      assignments,
    } : null;
  }).filter(Boolean);
  if (!rows.length) return null;
  return {
    sheetName: text(value.sheetName) || "VHH roster",
    blockIndex: Math.max(1, Number(value.blockIndex || 1)),
    visible: value.visible !== false,
    headerRow: Math.max(1, Number(value.headerRow || 1)),
    teachingTimetableRow: Math.max(0, Number(value.teachingTimetableRow || 0)),
    teachingTimetableEndRow: Math.max(0, Number(value.teachingTimetableEndRow || 0)),
    dates,
    rows,
  };
}

function vhhPeopleFromCell(value) {
  const raw = text(value);
  if (!raw || isNonPersonValue(raw)) return [];
  const byKey = new Map();
  for (const line of raw.split(/[\r\n;]+/)) {
    const timings = vhhTimingRanges(line);
    COMMA_NAME.lastIndex = 0;
    let match;
    while ((match = COMMA_NAME.exec(line))) {
      const person = personFromName(`${match[2]} ${match[1]}`);
      if (!person) continue;
      const timing = timings.length === 1 ? timings[0] : null;
      const existing = byKey.get(person.key);
      if (!existing || (!existing.timing && timing)) byKey.set(person.key, { ...person, timing });
    }
  }
  return [...byKey.values()];
}

function personFromName(value) {
  const displayName = titleCase(text(value).replace(/\s+/g, " "));
  const key = displayName.toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
  return key ? { key, displayName } : null;
}

function vhhTimingRanges(value) {
  const ranges = [];
  const expression = /(\d{1,2}:?\d{2})\s*[-–]\s*(\d{1,2}:?\d{2})/g;
  let match;
  while ((match = expression.exec(String(value || "")))) {
    const startTime = clock(match[1]);
    const endTime = clock(match[2]);
    if (startTime && endTime) ranges.push({ startTime, endTime });
  }
  return ranges;
}

function vhhEvent({ person, block, row, assignment, definition, seniority }) {
  const date = assignment.date;
  const timing = person.timing || definition.timing;
  const allDay = !timing;
  const endDate = timing && timing.endTime <= timing.startTime ? addDays(date, 1) : date;
  const start = allDay ? date : `${date}T${timing.startTime}:00`;
  const end = allDay ? addDays(date, 1) : `${endDate}T${timing.endTime}:00`;
  return {
    id: `vhh:${safeId(block.sheetName)}:${block.blockIndex}:${row.sourceRow}:${safeId(assignment.sourceCell)}:${person.key}`,
    source: "VHH",
    seniority,
    title: `VHH: ${definition.title}`,
    allDay,
    start,
    end,
    location: definition.hasLocation ? VHH_LOCATION : "",
    rawValue: assignment.namesText,
    timeLabel: allDay ? "All day" : `${timing.startTime}-${timing.endTime}`,
    monthKey: date.slice(0, 7),
  };
}

function buildTermSeniorities(blocks) {
  const evidence = new Map();
  for (const block of blocks) {
    for (const row of block.rows) {
      const definition = shiftDefinition(row.shiftLabel);
      if (!definition) continue;
      for (const assignment of row.assignments) {
        for (const person of vhhPeopleFromCell(assignment.namesText)) {
          const current = evidence.get(person.key) || "Unknown";
          if (seniorityRank(definition.seniority) > seniorityRank(current)) evidence.set(person.key, definition.seniority);
        }
      }
    }
  }
  return evidence;
}

function shift(title, seniority, startTime, endTime, hasLocation = true) {
  return {
    title,
    seniority,
    timing: startTime && endTime ? { startTime, endTime } : null,
    hasLocation,
  };
}

function shiftDefinition(label) {
  return SHIFT_DEFINITIONS.get(normaliseShiftLabel(label)) || null;
}

function isIgnoredShift(label) {
  return IGNORED_SHIFT_LABELS.has(normaliseShiftLabel(label));
}

function normaliseShiftLabel(label) {
  return text(label).toUpperCase().replace(/\s+/g, " ");
}

function seniorityRank(value) {
  if (value === "HMO") return 3;
  if (value === "SMS") return 2;
  if (value === "Junior Registrar") return 1;
  return 0;
}

function isNonPersonValue(value) {
  return /^(?:-|–|—|tba|vacant|leave|annual leave|sick leave|public holiday)$/i.test(text(value));
}

function clock(value) {
  const digits = String(value || "").replace(/\D/g, "");
  if (digits.length < 3 || digits.length > 4) return "";
  const padded = digits.padStart(4, "0");
  const hour = Number(padded.slice(0, 2));
  const minute = Number(padded.slice(2));
  if (hour === 24 && minute === 0) return "00:00";
  if (hour > 23 || minute > 59) return "";
  return `${String(hour).padStart(2, "0")}:${String(minute).padStart(2, "0")}`;
}

function addDays(date, days) {
  const value = new Date(`${date}T12:00:00Z`);
  value.setUTCDate(value.getUTCDate() + days);
  return value.toISOString().slice(0, 10);
}

function safeId(value) {
  return String(value || "").replace(/[^A-Za-z0-9]+/g, "-").replace(/^-|-$/g, "") || "cell";
}

function text(value) {
  return String(value ?? "").trim();
}

function titleCase(value) {
  return String(value || "").toLowerCase().replace(/\b([a-z])/g, (character) => character.toUpperCase());
}
