export const VHH_ROSTER_SOURCE_ID = "vhh-active-medical-roster";

const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;
const MAX_BLOCKS = 24;
const MAX_ROWS_PER_BLOCK = 80;
const MAX_ASSIGNMENTS_PER_ROW = 14;

// This is deliberately a JSON-to-calendar adapter rather than an Excel
// reader. SharePoint remains the authority for the workbook; Office Script
// extracts only the roster region and the GitHub worker converts that retained
// JSON to the ordinary D1 roster payload.
export function normaliseVhhRosterExtract(payload) {
  if (!payload || typeof payload !== "object" || String(payload.sourceId || "").trim() !== VHH_ROSTER_SOURCE_ID) return null;
  const blocks = Array.isArray(payload.blocks) ? payload.blocks : [];
  if (!blocks.length || blocks.length > MAX_BLOCKS) return null;
  const normalisedBlocks = blocks.map(normaliseBlock).filter(Boolean);
  if (!normalisedBlocks.length) return null;
  return {
    schemaVersion: 1,
    sourceId: VHH_ROSTER_SOURCE_ID,
    fileName: "Active Medical Roster.json",
    providerModifiedAt: String(payload.providerModifiedAt || "").trim(),
    providerVersion: String(payload.providerVersion || "").trim(),
    blocks: normalisedBlocks,
  };
}

export function buildVhhDerivedRosterPayload({ extract, contentHash, fileId = "", providerVersion = "" } = {}) {
  const roster = normaliseVhhRosterExtract(extract);
  if (!roster) throw new Error("VHH roster JSON is invalid.");
  const resolvedFileId = String(fileId || `automation:${VHH_ROSTER_SOURCE_ID}:${String(contentHash || "").slice(0, 24)}`);
  const addedAt = new Date().toISOString();
  const doctorsByKey = new Map();
  const eventsByDoctor = {};

  for (const block of roster.blocks) {
    for (const row of block.rows) {
      for (const assignment of row.assignments) {
        const people = vhhPeopleFromCell(assignment.namesText);
        if (!people.length) continue;
        const timings = vhhTimingRanges(assignment.namesText);
        for (const person of people) {
          if (!doctorsByKey.has(person.key)) {
            doctorsByKey.set(person.key, {
              key: person.key,
              displayName: person.displayName,
              sourceType: "vhh",
              seniority: seniorityForShift(row.shiftLabel),
              membershipSource: "roster",
            });
          }
          const events = eventsByDoctor[person.key] || (eventsByDoctor[person.key] = []);
          const ranges = timings.length ? timings : [null];
          ranges.forEach((timing, timingIndex) => {
            const event = vhhEvent({ person, block, row, assignment, timing, timingIndex });
            events.push(event);
          });
        }
      }
    }
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
      size: new TextEncoder().encode(JSON.stringify(roster)).byteLength,
      lastModified: Date.parse(roster.providerModifiedAt) || Date.now(),
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
    headerRow: Math.max(1, Number(value.headerRow || 1)),
    teachingTimetableRow: Math.max(1, Number(value.teachingTimetableRow || 1)),
    dates,
    rows,
  };
}

function vhhPeopleFromCell(value) {
  const raw = text(value);
  if (!raw || /^(?:-|tba|vacant|leave|annual leave|sick leave)$/i.test(raw)) return [];
  const people = [];
  const seen = new Set();
  // Keep a line break as a person boundary. A VHH cell can contain several
  // Last, First names on separate lines.
  const commaName = /([A-Za-z][A-Za-z'’.-]*(?:[ \t]+[A-Za-z][A-Za-z'’.-]*)*),[ \t]*([A-Za-z][A-Za-z'’.-]*(?:[ \t]+[A-Za-z][A-Za-z'’.-]*)*)/g;
  let match;
  while ((match = commaName.exec(raw))) addPerson(people, seen, `${match[2]} ${match[1]}`);
  if (!people.length) {
    for (const line of raw.split(/[\r\n;]+/)) {
      const candidate = line.replace(/\([^)]*\)/g, "").trim();
      if (/^[A-Za-z][A-Za-z'’.-]+(?:\s+[A-Za-z][A-Za-z'’.-]+){1,3}$/.test(candidate)) addPerson(people, seen, candidate);
    }
  }
  return people;
}

function addPerson(people, seen, value) {
  const displayName = titleCase(text(value).replace(/\s+/g, " "));
  const key = displayName.toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
  if (!key || seen.has(key)) return;
  seen.add(key);
  people.push({ key, displayName });
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

function vhhEvent({ person, block, row, assignment, timing, timingIndex }) {
  const date = assignment.date;
  const allDay = !timing;
  const endDate = timing && timing.endTime <= timing.startTime ? addDays(date, 1) : date;
  const start = allDay ? date : `${date}T${timing.startTime}:00`;
  const end = allDay ? addDays(date, 1) : `${endDate}T${timing.endTime}:00`;
  return {
    id: `vhh:${safeId(block.sheetName)}:${block.blockIndex}:${row.sourceRow}:${safeId(assignment.sourceCell)}:${person.key}:${timingIndex}`,
    source: "VHH",
    seniority: seniorityForShift(row.shiftLabel),
    title: `VHH: ${row.shiftLabel}`,
    allDay,
    start,
    end,
    location: "Victorian Heart Hospital, 631 Blackburn Road, Clayton VIC 3168, Australia",
    rawValue: assignment.namesText,
    timeLabel: allDay ? "All day (time not specified in roster)" : `${timing.startTime}-${timing.endTime}`,
    monthKey: date.slice(0, 7),
  };
}

function seniorityForShift(label) {
  const shift = String(label || "").toUpperCase();
  if (/\b(?:CST|SMS)\b/.test(shift)) return "SMS";
  if (/\bHMO\b/.test(shift)) return "HMO";
  if (/\bJMS\b/.test(shift)) return "Junior Registrar";
  return "Unknown";
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
