import * as XLSX from "xlsx";

const API_BASE = "https://www.findmyshift.com/api/1.4";
const NEXT_TERM_LOOKAHEAD_DAYS = 28;
const FINDMYSHIFT_MAX_RATE_LIMIT_RETRIES = 2;

export function findmyshiftConfiguredRosterRange(env = {}, now = new Date()) {
  // FindMyShift accepts the complete published roster window, but rejects an
  // open-ended multi-year request (HTTP 470).  The Diagnostic bounds are the
  // administrator-configurable published range; without them use the current
  // term until the next term is four weeks away. At that point, import the
  // upcoming term as its own source so it can appear in calendars before the
  // current term ends.
  return {
    from: validDateKey(env.FINDMYSHIFT_FROM) || validDateKey(env.FINDMYSHIFT_DIAGNOSTIC_FROM) || findmyshiftPublicationWindow(now).from,
    to: validDateKey(env.FINDMYSHIFT_TO) || validDateKey(env.FINDMYSHIFT_DIAGNOSTIC_TO) || findmyshiftPublicationWindow(now).to,
  };
}

function findmyshiftPublicationWindow(now) {
  const date = now instanceof Date && !Number.isNaN(now.getTime()) ? now : new Date();
  const today = date.toISOString().slice(0, 10);
  const terms = [];
  for (const year of [date.getUTCFullYear() - 1, date.getUTCFullYear(), date.getUTCFullYear() + 1]) {
    for (const monthIndex of [1, 4, 7, 10]) {
      const from = firstMondayDateKey(year, monthIndex);
      terms.push({ from, to: addDateKeyDays(from, 90) });
    }
  }
  const sorted = terms.sort((left, right) => left.from.localeCompare(right.from));
  const current = sorted.filter((term) => term.from <= today).at(-1)
    || { from: firstMondayDateKey(date.getUTCFullYear(), 1), to: addDateKeyDays(firstMondayDateKey(date.getUTCFullYear(), 1), 90) };
  const next = sorted.find((term) => term.from > today);
  return next && today >= addDateKeyDays(next.from, -NEXT_TERM_LOOKAHEAD_DAYS) ? next : current;
}

function firstMondayDateKey(year, monthIndex) {
  const date = new Date(Date.UTC(year, monthIndex, 1));
  const day = date.getUTCDay();
  date.setUTCDate(date.getUTCDate() + (day === 0 ? 1 : day === 1 ? 0 : 8 - day));
  return date.toISOString().slice(0, 10);
}

function addDateKeyDays(value, days) {
  const date = new Date(`${value}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

export async function findmyshiftLastModified(apiKey, teamId) {
  const payload = await findmyshiftRequest("teams/last-modified", apiKey, { teamId });
  const value = firstDateLikeValue(payload);
  if (!value) throw new Error("FindMyShift did not return a team modification time.");
  return value;
}

export async function findmyshiftRosterWorkbook(apiKey, teamId, range) {
  // FindMyShift allows only one concurrent request per API key.  Keep these
  // dependent lookups serial even though they are otherwise independent.
  const report = await findmyshiftShiftReport(apiKey, teamId, range);
  const staff = await findmyshiftStaffList(apiKey, teamId);
  const facilities = await findmyshiftFacilityList(apiKey, teamId);
  // FindMyShift's staff endpoint encodes DDH grades in ordered roster groups
  // (for example, "ED HMO's" followed by the doctors), rather than in each
  // person's jobTitle field. Preserve that source classification in both the
  // event rows and the retained staff sheet.
  const staffSeniorities = findmyshiftStaffSeniorityById(staff);
  const staffAssignments = findmyshiftStaffAssignmentById(staff);
  const shifts = extractShiftRows(report, { staff, facilities, staffSeniorities, staffAssignments });
  if (!shifts.length) throw noUsableFindmyshiftShiftsError();
  assertFindmyshiftDandenongAssignments(shifts);
  return findmyshiftRowsWorkbook(shifts, staff, { staffSeniorities });
}

function noUsableFindmyshiftShiftsError() {
  const error = new Error("FindMyShift returned no usable roster shifts for the configured date range.");
  error.code = "findmyshift-no-shifts";
  return error;
}

// A DDH shift must carry its stream in either its label or a facility value.
// The report API sometimes returns only a time range (for example 14:30-00:00)
// with no facility.  There is no reliable way to infer which DDH stream that
// represents, so importing it as a generic AM/PM/Night shift would silently
// replace a more precise manual roster with incorrect calendar entries.
export function findmyshiftDandenongAssignmentDiagnostics(rows = []) {
  const shifts = Array.isArray(rows) ? rows : [];
  const ambiguousRows = shifts.filter(isAmbiguousFindmyshiftTimedRow);
  const ambiguousTimedByLayout = {};
  for (const row of ambiguousRows) {
    const layout = String(row?.pairingIssue || "time-without-named-stream");
    ambiguousTimedByLayout[layout] = Number(ambiguousTimedByLayout[layout] || 0) + 1;
  }
  return {
    rows: shifts.length,
    timedRows: shifts.filter((row) => String(row?.start || "").trim() && String(row?.end || "").trim()).length,
    ambiguousTimed: ambiguousRows.length,
    ambiguousTimedByLayout,
    complete: ambiguousRows.length === 0,
  };
}

// This is intentionally limited to the creator-only review export.  It
// exposes the minimum information needed to reconcile a time-only row against
// the source roster: staff member, date, time range and structural reason.
export function findmyshiftDandenongAssignmentExceptions(rows = []) {
  return (Array.isArray(rows) ? rows : [])
    .filter(isAmbiguousFindmyshiftTimedRow)
    .map((row) => ({
      staffName: String(row.name || "").trim(),
      date: String(row.date || "").trim(),
      start: String(row.start || "").trim(),
      end: String(row.end || "").trim(),
      reason: String(row.pairingIssue || "time-without-named-stream").replace(/-/g, " "),
    }));
}

function isAmbiguousFindmyshiftTimedRow(row) {
  return String(row?.label || "").trim().toUpperCase() === "SHIFT"
    && String(row?.start || "").trim()
    && String(row?.end || "").trim()
    && !String(row?.facility || "").trim();
}

export function assertFindmyshiftDandenongAssignments(rows = []) {
  const diagnostics = findmyshiftDandenongAssignmentDiagnostics(rows);
  if (diagnostics.complete) return diagnostics;
  const layouts = Object.entries(diagnostics.ambiguousTimedByLayout)
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([layout, count]) => `${count} ${layout.replace(/-/g, " ")}`)
    .join(", ");
  const error = new Error(
    `FindMyShift did not include a stream or facility for ${diagnostics.ambiguousTimed} timed roster entries${layouts ? ` (${layouts})` : ""}. Automatic import was stopped so ambiguous Dandenong shifts cannot replace the detailed manual roster.`,
  );
  error.code = "findmyshift-incomplete-ddh-assignment";
  error.findmyshiftAssignmentExceptions = findmyshiftDandenongAssignmentExceptions(rows);
  throw error;
}

export async function findmyshiftShiftReport(apiKey, teamId, range) {
  // This is deliberately an unfiltered team report. In particular, do not send
  // `filters` or a facility id: doctors can be rostered across all facilities.
  // Keep it identical to the documented Developer API request which was
  // verified against this team. The returned flat rows already contain the
  // fields used by the importer; display/grouping options are not needed.
  return findmyshiftRequest("reports/shifts", apiKey, {
    teamId,
    from: range.from,
    to: range.to,
    // FindMyShift uses this to suppress free-text/availability rows while
    // retaining rostered time-and-stream assignment rows.
    comments: "no",
  });
}

export async function findmyshiftStaffList(apiKey, teamId) {
  return findmyshiftRequest("staff/list", apiKey, { teamId });
}

export async function findmyshiftFacilityList(apiKey, teamId) {
  return findmyshiftRequest("facilities/list", apiKey, { teamId });
}

export function findmyshiftReportDiagnostics(report, range = {}, options = {}) {
  const fieldCounts = {
    staff: 0,
    seniority: 0,
    date: 0,
    label: 0,
    timed: 0,
    allDay: 0,
    facility: 0,
    comment: 0,
  };
  const sourceDates = new Set();
  let objectCount = 0;
  let arrayCount = 0;
  let candidateCount = 0;
  const visit = (value, depth = 0) => {
    if (depth > 12 || value === null || value === undefined) return;
    if (Array.isArray(value)) {
      arrayCount += 1;
      for (const item of value) visit(item, depth + 1);
      return;
    }
    if (typeof value !== "object") return;
    objectCount += 1;
    const candidate = shiftCandidate(value);
    if (candidate) {
      candidateCount += 1;
      for (const field of candidateFields(value)) fieldCounts[field] += 1;
      const date = dateFrom(value);
      if (date) sourceDates.add(date);
    }
    for (const child of Object.values(value)) visit(child, depth + 1);
  };
  visit(report);
  const shifts = extractShiftRows(report, options);
  const dates = [...sourceDates].sort();
  return {
    responseFormat: Array.isArray(report) ? "array" : report && typeof report === "object" ? "object" : typeof report,
    objectCount,
    arrayCount,
    shiftCandidateCount: candidateCount,
    recognisedShiftCount: shifts.length,
    staffCount: new Set(shifts.map((shift) => shift.name)).size,
    dateRange: dates.length ? { from: dates[0], to: dates.at(-1) } : null,
    requestedDateRange: { from: String(range.from || ""), to: String(range.to || ""), label: String(range.label || "") },
    recognisedFields: fieldCounts,
    flatShiftRows: findmyshiftFlatShiftRowDiagnostics(report),
    responseShape: findmyshiftReportShape(report),
    facilityFilterApplied: false,
  };
}

export function findmyshiftStaffDiagnostics(staff) {
  const entries = Array.isArray(staff) ? staff : [];
  let id = 0;
  let name = 0;
  let seniority = 0;
  for (const entry of entries) {
    if (!entry || typeof entry !== "object") continue;
    if (entry.staffId || entry.id) id += 1;
    if (nameFrom(entry)) name += 1;
    if (entry.jobTitle || entry.department) seniority += 1;
  }
  return {
    responseFormat: Array.isArray(staff) ? "array" : staff && typeof staff === "object" ? "object" : typeof staff,
    itemCount: entries.length,
    recognised: { id, name, seniority },
  };
}

export function findmyshiftFacilityDiagnostics(facilities) {
  const entries = Array.isArray(facilities) ? facilities : [];
  let id = 0;
  let name = 0;
  for (const entry of entries) {
    if (!entry || typeof entry !== "object") continue;
    if (entry.facilityId || entry.id) id += 1;
    if (entry.facilityName || entry.name || entry.title) name += 1;
  }
  return {
    responseFormat: Array.isArray(facilities) ? "array" : facilities && typeof facilities === "object" ? "object" : typeof facilities,
    itemCount: entries.length,
    recognised: { id, name },
  };
}

function findmyshiftFlatShiftRowDiagnostics(report) {
  const rows = Array.isArray(report) ? report.filter((entry) => entry && typeof entry === "object" && !Array.isArray(entry)) : [];
  let staffId = 0;
  let date = 0;
  let shiftText = 0;
  let timeRangeLike = 0;
  let newline = 0;
  for (const row of rows) {
    if (row.staffId !== null && row.staffId !== undefined && String(row.staffId).trim()) staffId += 1;
    if (dateFrom(row)) date += 1;
    const shift = typeof row.shift === "string" ? row.shift.trim() : "";
    if (!shift) continue;
    shiftText += 1;
    if (/\b\d{1,2}(?::\d{2})?\s*(?:am|pm)?\s*(?:-|–|—|to)\s*\d{1,2}(?::\d{2})?\s*(?:am|pm)?\b/i.test(shift)) timeRangeLike += 1;
    if (/\r?\n/.test(shift)) newline += 1;
  }
  return {
    rowCount: rows.length,
    recognised: { staffId, date, shiftText },
    shiftTextCharacteristics: { timeRangeLike, newline },
  };
}

// The report diagnostic must be useful for adapting the parser, while never
// returning names, shift text, comments, facility names, IDs or arbitrary keys
// (the API may use staff names as object keys).  This intentionally reports
// only counts, JSON value types and a small allow-list of ordinary schema keys.
const SAFE_REPORT_SCHEMA_KEYS = new Set([
  "date", "dates", "day", "days", "staff", "staffid", "staffname", "staffmember",
  "employee", "employeename", "shifts", "shift", "shiftid", "shiftname", "shiftlabel",
  "start", "end", "starttime", "endtime", "time", "times", "duration", "hours",
  "facility", "facilities", "facilityname", "location", "locations", "comment", "comments",
  "note", "notes", "role", "title", "jobtitle", "department", "seniority", "leave",
  "timeoff", "absence", "published", "id", "name", "data", "items", "rows", "entries",
]);

function findmyshiftReportShape(report) {
  const levels = Array.from({ length: 4 }, () => ({
    arrays: 0,
    objects: 0,
    scalarValues: { string: 0, number: 0, boolean: 0, null: 0 },
    knownKeys: {},
    unknownKeyCount: 0,
  }));
  const visit = (value, depth = 0) => {
    if (depth >= levels.length) return;
    const level = levels[depth];
    if (value === null) {
      level.scalarValues.null += 1;
      return;
    }
    if (Array.isArray(value)) {
      level.arrays += 1;
      for (const child of value) visit(child, depth + 1);
      return;
    }
    if (typeof value === "object") {
      level.objects += 1;
      for (const [key, child] of Object.entries(value)) {
        const normalised = String(key).replace(/[^a-z0-9]/gi, "").toLowerCase();
        if (SAFE_REPORT_SCHEMA_KEYS.has(normalised)) {
          level.knownKeys[normalised] = Number(level.knownKeys[normalised] || 0) + 1;
        } else {
          level.unknownKeyCount += 1;
        }
        visit(child, depth + 1);
      }
      return;
    }
    if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
      level.scalarValues[typeof value] += 1;
    }
  };
  visit(report);
  return levels.map((level, depth) => ({
    depth,
    ...level,
  })).filter((level) => level.arrays || level.objects || Object.values(level.scalarValues).some(Boolean));
}

async function findmyshiftRequest(path, apiKey, params) {
  const body = new URLSearchParams({ apiKey: String(apiKey), ...Object.fromEntries(Object.entries(params).map(([key, value]) => [key, String(value)])) }).toString();
  let response;
  let text = "";
  for (let attempt = 0; attempt <= FINDMYSHIFT_MAX_RATE_LIMIT_RETRIES; attempt += 1) {
    response = await fetch(`${API_BASE}/${path}`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded", Accept: "application/json" },
      body,
    });
    text = await response.text();
    if (response.status !== 429 || attempt === FINDMYSHIFT_MAX_RATE_LIMIT_RETRIES) break;
    // FindMyShift permits only one request per key at a time. A manual Files
    // refresh can overlap the watchdog, so wait briefly and retry rather than
    // rejecting a perfectly valid automatic roster import.
    await waitFindmyshiftRetry(response.headers.get("retry-after"), attempt);
  }
  if (!response.ok) {
    throw findmyshiftRequestError(`FindMyShift ${path} returned HTTP ${response.status}.`, {
      code: "http-error",
      status: response.status,
      contentType: response.headers.get("content-type") || "",
      responseBytes: new TextEncoder().encode(text).byteLength,
    });
  }
  try {
    const parsed = JSON.parse(text);
    if (parsed?.error || parsed?.success === false) {
      throw findmyshiftRequestError("FindMyShift rejected the request.", {
        code: "provider-error",
        status: response.status,
        contentType: response.headers.get("content-type") || "",
        responseBytes: new TextEncoder().encode(text).byteLength,
      });
    }
    return parsed;
  } catch (error) {
    if (error instanceof SyntaxError) {
      throw findmyshiftRequestError("FindMyShift returned an unexpected response instead of JSON.", {
        code: "unexpected-format",
        status: response.status,
        contentType: response.headers.get("content-type") || "",
        responseBytes: new TextEncoder().encode(text).byteLength,
      });
    }
    throw error;
  }
}

async function waitFindmyshiftRetry(retryAfter, attempt) {
  const seconds = Number(retryAfter);
  const milliseconds = Number.isFinite(seconds) && seconds > 0
    ? Math.min(Math.round(seconds * 1000), 5_000)
    : 500 * (attempt + 1);
  await new Promise((resolve) => setTimeout(resolve, milliseconds));
}

function findmyshiftRequestError(message, details = {}) {
  const error = new Error(message);
  error.findmyshiftDiagnostic = {
    code: String(details.code || "request-failed"),
    status: Number(details.status || 0),
    contentType: String(details.contentType || "").split(";")[0].trim().slice(0, 80),
    responseBytes: Number(details.responseBytes || 0),
  };
  return error;
}

function validDateKey(value) {
  const date = String(value || "").trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(date) && !Number.isNaN(Date.parse(`${date}T00:00:00Z`)) ? date : "";
}

function firstDateLikeValue(value) {
  if (typeof value === "string" && !Number.isNaN(Date.parse(value))) return new Date(value).toISOString();
  if (value && typeof value === "object") {
    for (const child of Array.isArray(value) ? value : Object.values(value)) {
      const found = firstDateLikeValue(child);
      if (found) return found;
    }
  }
  return "";
}

export function extractShiftRows(report, options = {}) {
  const flatRows = extractFindmyshiftFlatRows(report, options);
  if (flatRows !== null) return dedupeShiftRows(flatRows);
  const rows = [];
  const visit = (value, inherited = {}) => {
    if (Array.isArray(value)) return value.forEach((item) => visit(item, inherited));
    if (!value || typeof value !== "object") return;
    const context = {
      name: nameFrom(value) || inherited.name || "",
      seniority: String(value.jobTitle || value.department || inherited.seniority || "Unknown"),
    };
    const date = dateFrom(value);
    const label = String(value.shiftName || value.name || value.title || value.label || value.role || "").trim();
    const start = timeFrom(value, ["startTime", "start", "timeStart", "from"]);
    const end = timeFrom(value, ["endTime", "end", "timeEnd", "to"]);
    if (context.name && date && (label || (start && end))) {
      rows.push({
        name: context.name,
        seniority: context.seniority,
        date,
        label: label || "Shift",
        start,
        end,
        facility: String(value.facilityName || value.facility || value.location || "").trim(),
        comment: String(value.comment || value.comments || value.notes || "").trim(),
      });
    }
    for (const child of Object.values(value)) if (child && typeof child === "object") visit(child, context);
  };
  visit(report);
  return dedupeShiftRows(rows);
}

// `reports/shifts` returns a flat array.  A row has staffId, firstName,
// lastName, date, shift, facilityId, payrollId and occurrences.  The API uses
// the `shift` text for either a named/all-day entry or a plain time range.
// Keep the branch explicit instead of sending it through the generic walker:
// a person's `firstName` would otherwise be mistaken for a shift label.
function extractFindmyshiftFlatRows(report, options) {
  if (!Array.isArray(report)) return null;
  const reportRows = report.filter((entry) => entry && typeof entry === "object" && !Array.isArray(entry));
  if (!reportRows.length || !reportRows.every((entry) => Object.hasOwn(entry, "staffId") && Object.hasOwn(entry, "date") && Object.hasOwn(entry, "shift"))) return null;
  const staffById = indexFindmyshiftStaff(options.staff);
  const staffSeniorities = options.staffSeniorities instanceof Map
    ? options.staffSeniorities
    : findmyshiftStaffSeniorityById(options.staff);
  const staffAssignments = options.staffAssignments instanceof Map
    ? options.staffAssignments
    : findmyshiftStaffAssignmentById(options.staff);
  const facilitiesById = indexFindmyshiftFacilities(options.facilities);
  const sourceRows = [];
  for (const entry of reportRows) {
    const staffId = String(entry.staffId || "").trim();
    const person = staffById.get(staffId) || {};
    const name = nameFrom(entry) || nameFrom(person);
    const date = dateFrom(entry);
    const shiftText = String(entry.shift || "").trim();
    if (!name || !date || !shiftText) continue;
    const time = timeRangeFromShiftText(shiftText);
    const facilityId = String(entry.facilityId || "").trim();
    sourceRows.push({
      sourceStaffId: staffId,
      name,
      seniority: findmyshiftSeniorityForStaff(person, entry, staffSeniorities.get(staffId)),
      date,
      label: time ? (shiftText === time.source ? "Shift" : shiftText) : shiftText,
      start: time?.start || "",
      end: time?.end || "",
      facility: facilitiesById.get(facilityId) || String(entry.facilityName || entry.facility || entry.location || facilityId || "").trim(),
      comment: String(entry.comment || entry.comments || entry.notes || "").trim(),
      sourceAssignment: staffAssignments.get(staffId) || "",
    });
  }
  return pairFindmyshiftTimeAndStreamRows(sourceRows);
}

// The FindMyShift shifts report represents a rostered shift as consecutive
// rows for the same person and day: a plain time range followed by its named
// stream.  The time row itself usually has no facility.  Treating those rows
// independently produces a generic AM/PM event plus a second all-day stream
// event, rather than the one stream-labelled timed shift shown in FindMyShift.
//
// Keep any extra named rows: they are genuine additional/all-day entries, not
// safe to discard just because they share a staff member and date.
function pairFindmyshiftTimeAndStreamRows(rows) {
  const paired = [];
  for (let index = 0; index < rows.length;) {
    const first = rows[index];
    const group = [first];
    index += 1;
    while (index < rows.length && sameFindmyshiftStaffDay(first, rows[index])) {
      group.push(rows[index]);
      index += 1;
    }
    paired.push(...pairFindmyshiftStaffDayRows(group));
  }
  return paired.map(applyKnownDandenongFindmyshiftAssignment);
}

// Kim Whelan is the Dandenong office worker. FindMyShift records his office
// shifts as time-only rows (the spreadsheet uses LSL for his non-working
// entries), so the one verified site-specific rule is that an otherwise
// unassigned timed row for him is Clinical Support. Shankar Thapaliya's
// corresponding entries are supported shifts: the second line names the
// clinician she is paired with, not a stream. The date-scoped entries below
// are separately verified corrections for the published Term 3 roster. Keep
// every rule narrow: none alters a named DDH stream nor guesses one from the
// paired clinician.
function applyKnownDandenongFindmyshiftAssignment(row) {
  if (!isAmbiguousFindmyshiftTimedRow(row)) return row;
  // FindMyShift represents Clinical Assistant shifts as a time range without
  // a stream. The authoritative ordered staff list identifies that role, so
  // preserve it as support work without guessing a clinical stream.
  if (row?.sourceAssignment) return { ...row, label: row.sourceAssignment, pairingIssue: "" };
  const name = normalizeFindmyshiftStaffName(row?.name);
  const key = `${name}|${String(row?.date || "").slice(0, 10)}`;
  const approvedLabels = {
    "LISETH JALABE|2026-08-04": "Paired AM",
    "LISETH JALABE|2026-08-05": "Paired AM",
    "LISETH JALABE|2026-08-06": "Paired AM",
    "LISETH JALABE|2026-08-07": "Paired AM",
    "STELLA TRAN|2026-08-04": "Paired AM",
    "STELLA TRAN|2026-08-05": "Paired AM",
    "STELLA TRAN|2026-08-06": "Paired AM",
    "STELLA TRAN|2026-08-07": "Paired AM",
    "DI FLOOD|2026-08-13": "S/L",
  };
  if (approvedLabels[key]) return { ...row, label: approvedLabels[key], pairingIssue: "" };
  if (name === "KIM WHELAN") return { ...row, label: "CS", pairingIssue: "" };
  if (name === "SHANKAR THAPALIYA") return { ...row, label: "Paired AM", pairingIssue: "" };
  return row;
}

function normalizeFindmyshiftStaffName(value) {
  return String(value || "").trim().replace(/\s+/g, " ").toUpperCase();
}

function sameFindmyshiftStaffDay(left, right) {
  return String(left?.sourceStaffId || left?.name || "") === String(right?.sourceStaffId || right?.name || "")
    && String(left?.date || "") === String(right?.date || "");
}

function pairFindmyshiftStaffDayRows(group) {
  const timedIndexes = group.map((row, index) => row.start && row.end ? index : -1).filter((index) => index >= 0);
  // The actual report's paired representation is strictly time first. Leave
  // unexpected layouts untouched so the safety check can reject an ambiguous
  // timed row rather than incorrectly attaching an unrelated named entry.
  const namedIndexes = group.map((row, index) => !row.start && !row.end && row.label ? index : -1).filter((index) => index >= 0);
  if (timedIndexes.length !== 1 || timedIndexes[0] !== 0 || !namedIndexes.length) {
    const pairingIssue = findmyshiftPairingIssue(timedIndexes, namedIndexes);
    return group.map((row) => row.start && row.end ? { ...row, pairingIssue } : row);
  }

  // When a staff/day has another named entry as well, the real report places
  // the stream/facility row last. Prefer a named row with a facility, then the
  // final named row, and preserve the other named row as its own entry.
  const assignmentIndex = namedIndexes.findLast((index) => String(group[index]?.facility || "").trim())
    ?? namedIndexes.at(-1);
  const timed = group[0];
  const assignment = group[assignmentIndex];
  const merged = {
    ...timed,
    label: assignment.label,
    facility: assignment.facility || timed.facility,
    comment: combineFindmyshiftComments(timed.comment, assignment.comment),
  };
  return group.flatMap((row, index) => {
    if (index === 0) return [merged];
    return index === assignmentIndex ? [] : [row];
  });
}

function findmyshiftPairingIssue(timedIndexes, namedIndexes) {
  if (timedIndexes.length > 1) return "multiple-time-rows";
  if (timedIndexes[0] > 0 && namedIndexes.length) return "named-stream-before-time";
  return "time-without-named-stream";
}

function combineFindmyshiftComments(...values) {
  return [...new Set(values.map((value) => String(value || "").trim()).filter(Boolean))].join("\n");
}

function indexFindmyshiftStaff(staff) {
  const index = new Map();
  for (const entry of Array.isArray(staff) ? staff : []) {
    if (!entry || typeof entry !== "object") continue;
    const id = String(entry.staffId || entry.id || "").trim();
    if (id) index.set(id, entry);
  }
  return index;
}

// The staff/list response is also the roster's ordered people list. DDH does
// not populate jobTitle for many clinicians, but places a coloured heading
// immediately before each grade's members. This is the authoritative, fully
// automatic source for those grades; do not turn it into a manual override.
export function findmyshiftStaffSeniorityById(staff) {
  const entries = (Array.isArray(staff) ? staff : [])
    .filter((entry) => entry && typeof entry === "object")
    .map((entry, index) => ({ entry, index }))
    .sort((left, right) => {
      const leftOrder = Number(left.entry.order);
      const rightOrder = Number(right.entry.order);
      const leftValid = Number.isFinite(leftOrder);
      const rightValid = Number.isFinite(rightOrder);
      if (leftValid && rightValid && leftOrder !== rightOrder) return leftOrder - rightOrder;
      if (leftValid !== rightValid) return leftValid ? -1 : 1;
      return left.index - right.index;
    });
  const grades = new Map();
  let currentGrade = "";
  for (const { entry } of entries) {
    const name = nameFrom(entry);
    const heading = findmyshiftStaffListHeading(name);
    if (heading) {
      // An unsupported section is still a boundary. Never let the preceding
      // grade bleed into a different FindMyShift staff group.
      currentGrade = heading.seniority;
      continue;
    }
    const id = String(entry.staffId || entry.id || "").trim();
    if (!id || isFindmyshiftSyntheticStaffName(name)) continue;
    const directGrade = recognisedFindmyshiftStaffSeniority(entry.jobTitle || entry.department);
    if (directGrade) grades.set(id, directGrade);
    else if (currentGrade) grades.set(id, currentGrade);
  }
  return grades;
}

// Some DDH support roles are rostered against a time only, not a clinical
// stream. Their ordered FindMyShift group is still sufficient to identify the
// work safely. Keep this distinct from seniority: it is an event assignment,
// not a grade and must never be inferred from a person's name or shift time.
export function findmyshiftStaffAssignmentById(staff) {
  const entries = findmyshiftOrderedStaffEntries(staff);
  const assignments = new Map();
  let currentAssignment = "";
  for (const { entry } of entries) {
    const name = nameFrom(entry);
    const heading = findmyshiftStaffListHeading(name);
    if (heading) {
      currentAssignment = heading.assignment || "";
      continue;
    }
    const id = String(entry.staffId || entry.id || "").trim();
    if (id && !isFindmyshiftSyntheticStaffName(name) && currentAssignment) assignments.set(id, currentAssignment);
  }
  return assignments;
}

function findmyshiftOrderedStaffEntries(staff) {
  return (Array.isArray(staff) ? staff : [])
    .filter((entry) => entry && typeof entry === "object")
    .map((entry, index) => ({ entry, index }))
    .sort((left, right) => {
      const leftOrder = Number(left.entry.order);
      const rightOrder = Number(right.entry.order);
      const leftValid = Number.isFinite(leftOrder);
      const rightValid = Number.isFinite(rightOrder);
      if (leftValid && rightValid && leftOrder !== rightOrder) return leftOrder - rightOrder;
      if (leftValid !== rightValid) return leftValid ? -1 : 1;
      return left.index - right.index;
    });
}

function findmyshiftSeniorityForStaff(person, entry, groupedGrade) {
  const direct = recognisedFindmyshiftStaffSeniority(
    person?.jobTitle || person?.department || entry?.jobTitle || entry?.department,
  );
  if (direct) return direct;
  if (groupedGrade) return groupedGrade;
  return String(person?.jobTitle || person?.department || entry?.jobTitle || entry?.department || "Unknown").trim() || "Unknown";
}

function findmyshiftStaffListHeading(value) {
  const upper = normaliseFindmyshiftSeniorityText(value);
  const seniority = /^(?:SENIOR MEDICAL STAFF|SMS)$/.test(upper) ? "SMS"
    : /^SENIOR REGISTRARS?$/.test(upper) ? "Senior Registrar"
      : /^(?:TRANSITIONAL|INTERMEDIATE) REGISTRARS?$/.test(upper) ? "Transitional/Intermediate Registrar"
        : /^JUNIOR REGISTRARS?$/.test(upper) ? "Junior Registrar"
          : /^(?:(?:ED|CRIT CARE) )?HMOS?$/.test(upper) ? "HMO"
            : /^CMOS?$/.test(upper) ? "CMO"
              : /^INTERNS?$/.test(upper) ? "Intern"
                : /^(?:NURSE PRACTITIONERS|NURSE PRAC\. CANDIDATES)$/.test(upper) ? "ENP"
                  : /^(?:AMPS?|ALLIED MEDICAL PRACTITIONERS?)$/.test(upper) ? "AMP"
                    : "";
  if (seniority) return { seniority, supported: true };
  // These are staff-list section rows, not people. They deliberately clear
  // group inheritance while keeping unsupported roles out of ED membership.
  if (/^(?:CLINICAL ASSISTANTS?|NURSE EDUCATORS?)$/.test(upper)) {
    return { seniority: "", supported: false, assignment: "Paired AM" };
  }
  return null;
}

function findmyshiftRosterGroupSeniority(value) {
  return findmyshiftStaffListHeading(value)?.seniority || "";
}

function isFindmyshiftSyntheticStaffName(value) {
  const upper = normaliseFindmyshiftSeniorityText(value);
  if (findmyshiftStaffListHeading(upper)) return true;
  return /^(?:(?:ED|SSU|CRIT CARE) )?(?:HMO|INTERN|SMS|CMO)(?: \d+)?$/.test(upper)
    || /^(?:(?:SENIOR|JUNIOR|TRANSITIONAL|INTERMEDIATE) )?REGISTRAR(?: \d+)?$/.test(upper);
}

function recognisedFindmyshiftStaffSeniority(value) {
  const upper = normaliseFindmyshiftSeniorityText(value);
  if (/\bINTERN\b/.test(upper)) return "Intern";
  if (/\b(?:ED |CRIT CARE )?HMOS?\b/.test(upper)) return "HMO";
  if (/\b(?:ENP|NP|NPC|NURSE PRACTITIONER)\b/.test(upper)) return "ENP";
  if (/\b(?:AMP|PHYSIO(?:THERAPIST)?|ALLIED MEDICAL PRACTITIONER)\b/.test(upper)) return "AMP";
  if (/\bSMS\b/.test(upper)) return "SMS";
  if (/\bCMO\b/.test(upper)) return "CMO";
  if (/\b(?:SENIOR REGISTRAR|SENIOR REG|SR)\b/.test(upper)) return "Senior Registrar";
  if (/\b(?:TRANSITIONAL|INTERMEDIATE|IR|TR)\b/.test(upper)) return "Transitional/Intermediate Registrar";
  if (/\b(?:JUNIOR REGISTRAR|JUNIOR REG|JR)\b/.test(upper)) return "Junior Registrar";
  return "";
}

function normaliseFindmyshiftSeniorityText(value) {
  return String(value || "").replace(/[’']/g, "").replace(/\s+/g, " ").trim().toUpperCase();
}

function indexFindmyshiftFacilities(facilities) {
  const index = new Map();
  for (const entry of Array.isArray(facilities) ? facilities : []) {
    if (!entry || typeof entry !== "object") continue;
    const id = String(entry.facilityId || entry.id || "").trim();
    const name = String(entry.facilityName || entry.name || entry.title || "").trim();
    if (id && name) index.set(id, name);
  }
  return index;
}

function timeRangeFromShiftText(value) {
  const match = String(value || "").trim().match(/^(\d{1,2}:\d{2})\s*(?:-|–|—|to)\s*(\d{1,2}:\d{2})$/i);
  if (!match) return null;
  return {
    source: String(value || "").trim(),
    start: normaliseTwentyFourHourTime(match[1]),
    end: normaliseTwentyFourHourTime(match[2]),
  };
}

function normaliseTwentyFourHourTime(value) {
  const [hour, minute] = String(value || "").split(":");
  return `${String(hour || "").padStart(2, "0")}:${minute}`;
}

function dedupeShiftRows(rows) {
  const unique = new Map();
  for (const row of rows) {
    const key = [row.sourceStaffId || row.name, row.date, row.label, row.start, row.end, row.facility, row.comment].join("|");
    unique.set(key, row);
  }
  return [...unique.values()];
}

function shiftCandidate(value) {
  if (!value || typeof value !== "object") return false;
  return Boolean(dateFrom(value) || timeFrom(value, ["startTime", "start", "timeStart", "from"]) || timeFrom(value, ["endTime", "end", "timeEnd", "to"]));
}

function candidateFields(value) {
  const result = [];
  if (nameFrom(value)) result.push("staff");
  if (value.jobTitle || value.department) result.push("seniority");
  if (dateFrom(value)) result.push("date");
  const shiftText = String(value.shift || value.shiftName || value.name || value.title || value.label || value.role || "").trim();
  const directTimed = timeFrom(value, ["startTime", "start", "timeStart", "from"]) && timeFrom(value, ["endTime", "end", "timeEnd", "to"]);
  const shiftTime = timeRangeFromShiftText(shiftText);
  if (shiftText) result.push("label");
  if (directTimed || shiftTime) result.push("timed");
  if (shiftText && !directTimed && !shiftTime) result.push("allDay");
  if (value.facilityId || value.facilityName || value.facility || value.location) result.push("facility");
  if (value.comment || value.comments || value.notes) result.push("comment");
  return result;
}

function nameFrom(value) {
  const direct = value.staffName || value.employeeName || value.displayName || value.fullName || "";
  if (direct) return String(direct).trim();
  const first = value.firstName || value.staffFirstName || "";
  const last = value.lastName || value.staffLastName || "";
  return `${first} ${last}`.trim();
}

function dateFrom(value) {
  for (const key of ["date", "shiftDate", "startDate", "startsAt", "startDateTime"]) {
    const raw = value[key];
    const date = raw && new Date(raw);
    if (date && !Number.isNaN(date.getTime())) return date.toISOString().slice(0, 10);
  }
  return "";
}

function timeFrom(value, keys) {
  for (const key of keys) {
    const raw = String(value[key] || "").trim();
    const match = raw.match(/(?:T|^)(\d{1,2}):(\d{2})/);
    if (match) return `${match[1].padStart(2, "0")}:${match[2]}`;
  }
  return "";
}

export function findmyshiftRowsWorkbook(rows, staff = [], options = {}) {
  const byWeek = new Map();
  for (const row of rows) {
    const monday = mondayFor(row.date);
    if (!byWeek.has(monday)) byWeek.set(monday, []);
    byWeek.get(monday).push(row);
  }
  const output = [];
  for (const [monday, weekRows] of [...byWeek.entries()].sort()) {
    const days = Array.from({ length: 7 }, (_, index) => addDays(monday, index));
    output.push(["", ...days.map(findmyshiftDateLabel)]);
    const bySeniority = new Map();
    for (const row of weekRows) {
      if (!bySeniority.has(row.seniority)) bySeniority.set(row.seniority, new Map());
      const people = bySeniority.get(row.seniority);
      if (!people.has(row.name)) people.set(row.name, []);
      people.get(row.name).push(row);
    }
    for (const [seniority, people] of bySeniority) {
      output.push([seniority]);
      for (const [name, personRows] of people) {
        const dayRows = (day) => personRows.filter((row) => row.date === day);
        const labels = days.map((day) => dayRows(day).map((row) => row.label === "Shift" && row.start && row.end ? `${row.start}-${row.end}` : row.label).join(" / "));
        const times = days.map((day) => dayRows(day)
          .filter((row) => row.label !== "Shift")
          .map((row) => row.start && row.end ? `${row.start}-${row.end}` : "")
          .filter(Boolean)
          .join(" / "));
        output.push([name, ...labels]);
        if (times.some(Boolean)) output.push(["", ...times]);
      }
    }
  }
  const workbook = XLSX.utils.book_new();
  // This marker makes the existing DDH parser treat single-day leave from the
  // API as a single day, rather than the legacy spreadsheet convention of one
  // leave label standing for a whole week.
  output.unshift(["FindMyShift roster format"]);
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(output), "FindMyShift");
  const details = [["Staff ID", "Staff name", "Seniority/job title", "Date", "Shift label", "Start", "End", "Facility", "Comment"]];
  for (const row of rows) {
    details.push([row.sourceStaffId || "", row.name, row.seniority, row.date, row.label, row.start, row.end, row.facility, row.comment]);
  }
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(details), "FindMyShift details");
  const staffRows = [["Staff ID", "Staff name", "Seniority/job title"]];
  const staffSeniorities = options.staffSeniorities instanceof Map
    ? options.staffSeniorities
    : findmyshiftStaffSeniorityById(staff);
  const seenStaff = new Set();
  for (const person of Array.isArray(staff) ? staff : []) {
    const id = String(person?.staffId || person?.id || "").trim();
    const name = nameFrom(person);
    if (!name) continue;
    if (findmyshiftRosterGroupSeniority(name) || isFindmyshiftSyntheticStaffName(name)) continue;
    const marker = id || name.toUpperCase();
    if (seenStaff.has(marker)) continue;
    seenStaff.add(marker);
    staffRows.push([id, name, findmyshiftSeniorityForStaff(person, {}, staffSeniorities.get(id))]);
  }
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(staffRows), "FindMyShift staff");
  return XLSX.write(workbook, { type: "array", bookType: "xlsx" });
}

function mondayFor(value) {
  const date = new Date(`${value}T00:00:00Z`);
  const day = (date.getUTCDay() + 6) % 7;
  date.setUTCDate(date.getUTCDate() - day);
  return date.toISOString().slice(0, 10);
}

function addDays(value, count) {
  const date = new Date(`${value}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() + count);
  return date.toISOString().slice(0, 10);
}

function findmyshiftDateLabel(value) {
  const date = new Date(`${value}T00:00:00Z`);
  const weekdays = ["Sun.", "Mon.", "Tue.", "Wed.", "Thu.", "Fri.", "Sat."];
  const months = ["Jan.", "Feb.", "Mar.", "Apr.", "May.", "Jun.", "Jul.", "Aug.", "Sep.", "Oct.", "Nov.", "Dec."];
  return `${weekdays[date.getUTCDay()]} ${months[date.getUTCMonth()]} ${String(date.getUTCDate()).padStart(2, "0")}, ${date.getUTCFullYear()}`;
}
