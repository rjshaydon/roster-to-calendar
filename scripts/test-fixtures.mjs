import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import XLSX from "xlsx";

import { onRequestPost as handleStatePost } from "../functions/api/state.js";
import { onRequestGet as handleFeedGet } from "../functions/api/feed.js";
import { assertFindmyshiftDandenongAssignments, extractShiftRows, findmyshiftConfiguredRosterRange, findmyshiftDandenongAssignmentDiagnostics, findmyshiftDandenongAssignmentExceptions, findmyshiftRowsWorkbook, findmyshiftStaffAssignmentById, findmyshiftStaffSeniorityById } from "../functions/_lib/findmyshift.js";
import { buildAutomatedDerivedRosterPayload } from "../functions/_lib/automation-import.js";
import { australianTermStartForDate, buildPreviewFromDerivedEvents, findRosterSyncByProviderVersion, isApprovedReparseOmission, sameRosterOccurrence, storeCachedSnapshot } from "../functions/_lib/d1-calendar.js";
import { recordRosterDispatchLifecycle, requestQueuedRosterProcessing } from "../functions/_lib/automation-dispatch.js";
import { applyRosterEventSeniorities, attachFindmyshiftStaffIds, buildRosterView, customEventsToEvents, doctorOptions, findmyshiftProviderStaffOptions, findmyshiftRosteredStaffOptions, mergeMembershipDoctors, parseUploadForm, parserRuleDefaults, previewSummary, setParserExtensions } from "../public/static/roster.js";
import { customEventsToEvents as serverCustomEventsToEvents } from "../functions/_lib/roster.js";
import { parserResultDelta, unresolvedCodeSummary } from "./parser-parity.mjs";

assert.deepEqual(
  parserResultDelta(
    { events: [{ title: "DDH" }], issues: [{ status: "unknown", source: "DDH", seniority: "HMO", rawValue: "X" }] },
    { events: [{ title: "DDH" }], issues: [{ status: "unknown", source: "DDH", seniority: "HMO", rawValue: "X" }] },
  ),
  {},
  "parser comparison helpers should report no delta for equivalent normalized output",
);
assert.equal(
  sameRosterOccurrence(
    { doctorKey: "EXAMPLE", source: "DDH", start: "2026-08-10T08:00:00+10:00", rawValue: "Orange AM" },
    { doctorKey: "EXAMPLE", source: "DDH", start: "2026-08-10T07:30:00+10:00", end: "2026-08-10T17:30:00+10:00", rawValue: "Orange AM" },
  ),
  true,
  "a corrected timed event must preserve its roster occurrence on the same date",
);
assert.equal(
  australianTermStartForDate("2026-08-17"),
  "2026-08-03",
  "effective staff seniorities should use the same Term 3 boundary as the At a glance editor",
);
assert.deepEqual(
  mergeMembershipDoctors(
    [{ key: "ROSTERED DOCTOR", displayName: "Rostered Doctor", sourceType: "ddh", seniority: "Unknown" }],
    [
      { key: "ROSTERED DOCTOR", displayName: "Rostered Doctor", sourceType: "ddh", seniority: "HMO", membershipSource: "provider" },
      { key: "DIRECTORY ONLY", displayName: "Directory Only", sourceType: "ddh", seniority: "Unknown", membershipSource: "provider" },
    ],
  ),
  [{ key: "ROSTERED DOCTOR", displayName: "Rostered Doctor", sourceType: "ddh", seniority: "HMO", membershipSource: "roster" }],
  "FindMyShift directory records must enrich rostered people without creating provider-only ED staff",
);
assert.deepEqual(
  applyRosterEventSeniorities(
    [{ key: "PROGRESSION", seniority: "Intern" }],
    { PROGRESSION: [{ seniority: "Intern", start: "2026-08-03" }, { seniority: "HMO", start: "2026-08-17" }, { seniority: "Unknown", start: "2026-08-24" }] },
  ),
  [{ key: "PROGRESSION", seniority: "HMO" }],
  "the most recent known rostered grade should become the persisted membership grade",
);
assert.equal(
  sameRosterOccurrence(
    { doctorKey: "EXAMPLE", source: "DDH", start: "2026-08-11", rawValue: "AL" },
    { doctorKey: "EXAMPLE", source: "DDH", start: "2026-08-10", end: "2026-08-13", allDay: true, rawValue: "AL / AL / AL" },
  ),
  true,
  "a merged all-day leave event must preserve each source-day occurrence",
);
assert.equal(
  isApprovedReparseOmission({
    doctorKey: "HWEE MIN LEE", source: "DDH", title: "Annual Leave", start: "2026-05-04", end: "2026-05-11", rawValue: "AL",
  }, "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007"),
  true,
  "the approved DDH weekly-leave replacement must be scoped to its exact retained source event",
);
assert.equal(
  isApprovedReparseOmission({
    doctorKey: "HWEE MIN LEE", source: "DDH", title: "Annual Leave", start: "2026-05-04", end: "2026-05-12", rawValue: "AL",
  }, "Dandenong_Emergency_Doctors'_Roster_04-05-2026_to_02-08-2026.xlsx:137815:1778982385007"),
  false,
  "a changed leave range must not inherit the one-off migration approval",
);
assert.equal(
  isApprovedReparseOmission({
    doctorKey: "AMY LEUTHAUSER", source: "MMC", title: "MMC: CS", start: "2026-05-28", end: "2026-05-28", rawValue: "0800-1730 CS DH",
  }),
  true,
  "a DDH clinical-support reference recorded in an MMC roster is an approved omission",
);
assert.equal(
  isApprovedReparseOmission({
    doctorKey: "MICKEY FERGUSON", source: "MMC", title: "Annual Leave", start: "2026-07-13", rawValue: "A/L",
  }, "automation:monash-adults:ac7f9d2e29c6bbb35e8a86df"),
  true,
  "the directly reviewed legacy MMC leave omission must be limited to its retained source",
);
assert.equal(
  isApprovedReparseOmission({
    doctorKey: "MICKEY FERGUSON", source: "MMC", title: "Annual Leave", start: "2026-07-13", rawValue: "A/L",
  }, "different-retained-file"),
  false,
  "an unsupported legacy leave omission must not apply to another source file",
);
assert.equal(
  isApprovedReparseOmission({
    doctorKey: "FRANK SODEN", source: "DDH", title: "DDH: CAN WORK 1 EXTRA THIS WEEK", start: "2026-04-06", rawValue: "Can work 1 extra this week",
  }),
  true,
  "a reviewed DDH roster-writer request should be safely omitted on reparse",
);
assert.equal(
  isApprovedReparseOmission({
    doctorKey: "FRANK SODEN", source: "DDH", title: "DDH: Can work", start: "2026-04-06", rawValue: "Can work extra AM",
  }),
  false,
  "a different DDH message-like value must not receive a broad reparse-removal approval",
);
assert.equal(isApprovedReparseOmission({ source: "DDH", rawValue: "CS - not onsite (standard setting)" }), true, "approved DDH Clinical Support requests should not block a safe reparse");
assert.equal(isApprovedReparseOmission({ source: "DDH", rawValue: "08H00" }), true, "a DDH VHH late-early time fragment omitted by the contextual parser should not block a safe reparse");
assert.equal(isApprovedReparseOmission({ source: "DDH", rawValue: "Shift in lieu" }), true, "approved DDH compensatory annotations should not block a safe reparse");
assert.deepEqual(
  unresolvedCodeSummary([
    { status: "unknown", source: "DDH", seniority: "HMO", rawValue: "X" },
    { status: "unknown", source: "DDH", seniority: "HMO", rawValue: "X" },
    { status: "resolved", source: "DDH", seniority: "HMO", rawValue: "Y" },
  ]),
  { occurrences: 2, distinctCodes: 1 },
  "unknown-code reporting must distinguish occurrences from distinct codes",
);

function cloneWorkbook(workbook) {
  return XLSX.read(XLSX.write(workbook, { type: "array", bookType: "xlsx" }), { type: "array", cellDates: true });
}

function workbookFile(workbook, name) {
  const bytes = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
  return new File([bytes], name, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

function workbookDataUrl(workbook) {
  const bytes = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
  return `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${Buffer.from(bytes).toString("base64")}`;
}

function eventDates(event) {
  const dates = [];
  let cursor = String(event.start_date || "");
  const end = String(event.end_date || event.start_date || "");
  while (cursor && end && cursor <= end) {
    dates.push(cursor);
    const [year, month, day] = cursor.split("-").map(Number);
    const next = new Date(Date.UTC(year || 1970, (month || 1) - 1, day || 1));
    next.setUTCDate(next.getUTCDate() + 1);
    cursor = next.toISOString().slice(0, 10);
  }
  return dates;
}

function sharedEventDate(left, right, start, end) {
  const rightDates = new Set(eventDates(right).filter((date) => date >= start && date <= end));
  return eventDates(left).some((date) => date >= start && date <= end && rightDates.has(date));
}

const syncedAnnualLeave = [{
  id: "synced-annual-leave",
  source: "DDH",
  title: "Annual Leave",
  rawValue: "Annual leave",
  allDay: true,
  start: "2026-08-03",
  end: "2026-08-10",
}];
const customLeaveFallbacks = [
  { id: "covered-manual-leave", title: "A/L", startDate: "2026-08-03", endDate: "2026-08-09", allDay: true, startTime: "", endTime: "", location: "", include: true },
  { id: "partially-covered-manual-leave", title: "Annual Leave", startDate: "2026-08-03", endDate: "2026-08-12", allDay: true, startTime: "", endTime: "", location: "", include: true },
  { id: "single-day-manual-leave", title: "S/L", startDate: "2026-08-04", endDate: "2026-08-04", allDay: true, startTime: "", endTime: "", location: "", include: true },
  { id: "multi-day-custom-nonleave", title: "Personal reminder", startDate: "2026-08-03", endDate: "2026-08-09", allDay: true, startTime: "", endTime: "", location: "", include: true },
];
const expectedVisibleCustomFallbackIds = ["partially-covered-manual-leave", "single-day-manual-leave", "multi-day-custom-nonleave"];
assert.deepEqual(
  customEventsToEvents(customLeaveFallbacks, undefined, syncedAnnualLeave).map((event) => event.id),
  expectedVisibleCustomFallbackIds,
  "a synced leave should suppress only a fully covered multi-day manual leave fallback",
);
assert.deepEqual(
  serverCustomEventsToEvents(customLeaveFallbacks, undefined, syncedAnnualLeave).map((event) => event.id),
  expectedVisibleCustomFallbackIds,
  "subscription and export leave precedence must match the browser preview",
);

function withWorkbookCell(workbook, sheetName, cell, value) {
  const copy = cloneWorkbook(workbook);
  copy.Sheets[sheetName][cell] = value;
  return copy;
}

function withWorkbookDate(workbook, sheetName, cell, date) {
  return withWorkbookCell(workbook, sheetName, cell, { t: "d", v: date, w: date.toLocaleDateString("en-AU") });
}

async function assertRejectsMixedTermUpload(label, workbook, filename, expectedParts) {
  const formData = new FormData();
  formData.append("rosterFiles", workbookFile(workbook, filename));
  await assert.rejects(
    () => parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: formData })),
    (error) => {
      assert.ok(expectedParts.every((part) => error.message.includes(part)), `${label}: ${error.message}`);
      return true;
    },
  );
}

const mmcWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/AdultTerm1.2026.xlsx", import.meta.url)), {
  cellDates: true,
});
const ddhWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/Dandenong_Emergency_Doctors_Roster_02-02-2026_to_03-05-2026.xlsx", import.meta.url)), {
  cellDates: true,
});
const findmyshiftFixture = JSON.parse(await readFile(new URL("../fixtures/findmyshift-reports-shifts.sanitised.json", import.meta.url), "utf8"));
assert.deepEqual(
  findmyshiftConfiguredRosterRange({ FINDMYSHIFT_FROM: "2026-08-03", FINDMYSHIFT_TO: "2026-11-01" }, new Date("2026-08-06T00:00:00Z")),
  { from: "2026-08-03", to: "2026-11-01" },
  "FindMyShift automation should use the configured full available roster range",
);
assert.deepEqual(
  findmyshiftConfiguredRosterRange({ FINDMYSHIFT_DIAGNOSTIC_FROM: "2026-08-03", FINDMYSHIFT_DIAGNOSTIC_TO: "2026-11-01" }, new Date("2026-08-06T00:00:00Z")),
  { from: "2026-08-03", to: "2026-11-01" },
  "FindMyShift automation should use the configured provider-compatible published roster window",
);
assert.deepEqual(
  findmyshiftConfiguredRosterRange({}, new Date("2026-08-06T00:00:00Z")),
  { from: "2026-08-03", to: "2026-11-01" },
  "FindMyShift automation should default to the current provider-compatible roster term",
);
assert.deepEqual(
  findmyshiftConfiguredRosterRange({}, new Date("2026-10-04T00:00:00Z")),
  { from: "2026-08-03", to: "2026-11-01" },
  "FindMyShift should retain the current term until the four-week publication window opens",
);
assert.deepEqual(
  findmyshiftConfiguredRosterRange({}, new Date("2026-10-05T00:00:00Z")),
  { from: "2026-11-02", to: "2027-01-31" },
  "FindMyShift should begin importing Term 4 four weeks before it starts",
);
assert.deepEqual(
  findmyshiftConfiguredRosterRange({}, new Date("2027-01-04T00:00:00Z")),
  { from: "2027-02-01", to: "2027-05-02" },
  "FindMyShift should roll into the following year using the same four-week lead time",
);
const findmyshiftRows = extractShiftRows(findmyshiftFixture.report, {
  staff: findmyshiftFixture.staff,
  facilities: findmyshiftFixture.facilities,
});
const findmyshiftPreliminaryRows = extractShiftRows(findmyshiftFixture.report);
const groupedFindmyshiftStaff = [
  { staffId: "sms-heading", displayName: "SENIOR MEDICAL STAFF", order: 3 },
  { staffId: "hmo-heading", displayName: "ED HMO's", order: 840 },
  { staffId: "gideo", firstName: "Gideon", lastName: "Charin", order: 923, jobTitle: null, department: null },
  { staffId: "hmo-heading-two", displayName: "HMO's", order: 937 },
  { staffId: "tea", firstName: "Tea", lastName: "Gunasena", order: 1211, jobTitle: null, department: null },
  { staffId: "amp-heading", displayName: "AMP's", order: 1899 },
  { staffId: "physio", firstName: "Pat", lastName: "Physio", order: 1900, jobTitle: null, department: null },
];
const groupedFindmyshiftSeniorities = findmyshiftStaffSeniorityById(groupedFindmyshiftStaff);
assert.deepEqual(
  Object.fromEntries(groupedFindmyshiftSeniorities),
  { gideo: "HMO", tea: "HMO", physio: "AMP" },
  "FindMyShift ordered roster headings should classify DDH staff without job titles",
);
const groupedFindmyshiftRows = extractShiftRows([
  { staffId: "gideo", date: "2026-08-03", firstName: "Gideon", lastName: "Charin", shift: "ED AM" },
  { staffId: "physio", date: "2026-08-03", firstName: "Pat", lastName: "Physio", shift: "Physiotherapist" },
], { staff: groupedFindmyshiftStaff });
assert.deepEqual(
  groupedFindmyshiftRows.map((row) => row.seniority),
  ["HMO", "AMP"],
  "FindMyShift shift rows should retain grades derived from their ordered staff groups",
);
const authoritativeFindmyshiftStaff = [
  { staffId: "hmo-heading", displayName: "ED HMO's", order: 10 },
  { staffId: "hmo-person", firstName: "Hmo", lastName: "Person", order: 11 },
  { staffId: "np-heading", displayName: "NURSE PRACTITIONERS", order: 20 },
  { staffId: "np-person", firstName: "Nurse", lastName: "Practitioner", order: 21, jobTitle: "NP" },
  { staffId: "candidate-heading", displayName: "NURSE PRAC. CANDIDATES", order: 30 },
  { staffId: "candidate", firstName: "Nurse", lastName: "Candidate", order: 31 },
  { staffId: "amp-heading", displayName: "AMP's", order: 40 },
  { staffId: "amp-person", firstName: "Amp", lastName: "Person", order: 41 },
  { staffId: "clinical-heading", displayName: "CLINICAL ASSISTANTS", order: 50 },
  { staffId: "clinical-person", firstName: "Clinical", lastName: "Person", order: 51 },
  { staffId: "educator-heading", displayName: "NURSE EDUCATORS", order: 60 },
  { staffId: "educator-person", firstName: "Fathima", lastName: "Support", order: 61 },
  { staffId: "synthetic", displayName: "HMO 1", order: 52 },
];
const authoritativeGrades = findmyshiftStaffSeniorityById(authoritativeFindmyshiftStaff);
assert.deepEqual(
  Object.fromEntries(authoritativeGrades),
  { "hmo-person": "HMO", "np-person": "ENP", candidate: "ENP", "amp-person": "AMP" },
  "FindMyShift groups must include NPs and stop at unsupported headings rather than leaking the preceding grade",
);
assert.deepEqual(
  Object.fromEntries(findmyshiftStaffAssignmentById(authoritativeFindmyshiftStaff)),
  { "clinical-person": "Paired AM", "educator-person": "Paired AM" },
  "source-defined DDH support groups must safely classify their time-only support shifts",
);
const clinicalAssistantRows = extractShiftRows([
  { staffId: "clinical-person", facilityId: null, date: "2026-08-24", firstName: "Clinical", lastName: "Person", shift: "08:00-17:30" },
  { staffId: "educator-person", facilityId: null, date: "2026-08-24", firstName: "Fathima", lastName: "Support", shift: "08:00-17:30" },
], { staff: authoritativeFindmyshiftStaff });
assert.deepEqual(
  clinicalAssistantRows.map((row) => ({ label: row.label, start: row.start, end: row.end, facility: row.facility, seniority: row.seniority })),
  [
    { label: "Paired AM", start: "08:00", end: "17:30", facility: "", seniority: "Unknown" },
    { label: "Paired AM", start: "08:00", end: "17:30", facility: "", seniority: "Unknown" },
  ],
  "support-role time-only rows must use their source-defined assignment rather than a guessed stream",
);
assert.doesNotThrow(
  () => assertFindmyshiftDandenongAssignments(clinicalAssistantRows),
  "source-defined Clinical Assistant support shifts should not block the automatic import",
);
const registrarHeadingWorkbook = XLSX.read(findmyshiftRowsWorkbook([
  { sourceStaffId: "junior-person", name: "Junior Person", seniority: "Junior Registrar", date: "2026-08-03", label: "Orange AM", start: "08:00", end: "17:30", facility: "Orange AM", comment: "" },
  { sourceStaffId: "senior-person", name: "Senior Person", seniority: "Senior Registrar", date: "2026-08-03", label: "Orange PM", start: "14:30", end: "00:00", facility: "Orange PM", comment: "" },
]), { type: "array", cellDates: true });
const registrarHeadingFormData = new FormData();
registrarHeadingFormData.append("rosterFiles", workbookFile(registrarHeadingWorkbook, "Dandenong-FindMyShift-registrar-headings.xlsx"));
const registrarHeadingUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: registrarHeadingFormData }));
assert.deepEqual(
  doctorOptions([], registrarHeadingUpload.sources.ddh).map((doctor) => doctor.key).sort(),
  ["JUNIOR PERSON", "SENIOR PERSON"],
  "singular FindMyShift Junior and Senior Registrar group headings must never become DDH staff",
);
const authoritativeWorkbook = XLSX.read(findmyshiftRowsWorkbook([
  { sourceStaffId: "hmo-person", name: "Hmo Person", seniority: "HMO", date: "2026-08-03", label: "Orange AM", start: "07:30", end: "17:00", facility: "Orange AM", comment: "" },
  { sourceStaffId: "amp-person", name: "Amp Person", seniority: "AMP", date: "2026-08-03", label: "Physiotherapist", start: "09:30", end: "18:00", facility: "", comment: "" },
], authoritativeFindmyshiftStaff), { type: "array", cellDates: true });
const authoritativeFormData = new FormData();
authoritativeFormData.append("rosterFiles", workbookFile(authoritativeWorkbook, "Dandenong-FindMyShift-authoritative.xlsx"));
const authoritativeUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: authoritativeFormData }));
const authoritativeProviders = findmyshiftProviderStaffOptions(authoritativeUpload.sources.ddh);
const authoritativeRosteredStaff = findmyshiftRosteredStaffOptions(authoritativeUpload.sources.ddh);
assert.deepEqual(
  authoritativeProviders.map((person) => ({ key: person.key, seniority: person.seniority, providerStaffId: person.providerStaffId })).sort((left, right) => left.key.localeCompare(right.key)),
  [
    { key: "AMP PERSON", seniority: "AMP", providerStaffId: "amp-person" },
    { key: "CLINICAL PERSON", seniority: "Unknown", providerStaffId: "clinical-person" },
    { key: "FATHIMA SUPPORT", seniority: "Unknown", providerStaffId: "educator-person" },
    { key: "HMO PERSON", seniority: "HMO", providerStaffId: "hmo-person" },
    { key: "NURSE CANDIDATE", seniority: "ENP", providerStaffId: "candidate" },
    { key: "NURSE PRACTITIONER", seniority: "ENP", providerStaffId: "np-person" },
  ],
  "the staff directory must return grade evidence by FindMyShift staff ID while excluding headings and synthetic slots",
);
assert.deepEqual(
  mergeMembershipDoctors(attachFindmyshiftStaffIds([{ key: "AMP PERSON", displayName: "Amp Person", sourceType: "ddh", seniority: "Unknown" }], authoritativeRosteredStaff), authoritativeProviders),
  [{ key: "AMP PERSON", displayName: "Amp Person", sourceType: "ddh", seniority: "AMP", membershipSource: "roster", providerStaffId: "amp-person" }],
  "only rostered staff should receive provider grades; directory-only people must not become ED members",
);
assert.equal(
  buildRosterView([], authoritativeUpload.sources.ddh, "AMP PERSON").events[0]?.seniority,
  "AMP",
  "a timed FindMyShift Physiotherapist assignment must remain a real AMP shift",
);
assert.equal(findmyshiftRows.length, 5, "FindMyShift time/stream pairs should become one timed event while extra named entries are preserved");
assert.deepEqual(
  findmyshiftDandenongAssignmentDiagnostics(findmyshiftPreliminaryRows),
  { rows: 5, timedRows: 3, ambiguousTimed: 0, ambiguousTimedByLayout: {}, complete: true },
  "FindMyShift paired time/stream rows can be quality-checked before optional staff and facility lookups",
);
assert.deepEqual(
  findmyshiftDandenongAssignmentDiagnostics(findmyshiftRows),
  { rows: 5, timedRows: 3, ambiguousTimed: 0, ambiguousTimedByLayout: {}, complete: true },
  "FindMyShift paired time/stream rows may be imported with meaningful DDH assignments",
);
assert.doesNotThrow(
  () => assertFindmyshiftDandenongAssignments(findmyshiftRows),
  "FindMyShift rows that retain a facility must remain importable",
);
assert.throws(
  () => assertFindmyshiftDandenongAssignments([{ name: "Example Doctor", date: "2026-08-06", label: "Shift", start: "14:30", end: "00:00", facility: "" }]),
  (error) => {
    assert.match(error.message, /did not include a stream or facility/i);
    assert.deepEqual(error.findmyshiftAssignmentExceptions, [{ staffName: "Example Doctor", date: "2026-08-06", start: "14:30", end: "00:00", reason: "time without named stream" }]);
    return true;
  },
  "FindMyShift rows without a stream or facility must be rejected rather than guessed and identify the rows for review",
);
assert.deepEqual(
  findmyshiftDandenongAssignmentDiagnostics([{ label: "Shift", start: "14:30", end: "00:00", facility: "", pairingIssue: "named-stream-before-time" }]),
  { rows: 1, timedRows: 1, ambiguousTimed: 1, ambiguousTimedByLayout: { "named-stream-before-time": 1 }, complete: false },
  "FindMyShift pairing diagnostics should describe only the safe structural cause of an incomplete timed row",
);
assert.deepEqual(
  findmyshiftDandenongAssignmentExceptions([{ name: "Example Doctor", date: "2026-08-06", label: "Shift", start: "14:30", end: "00:00", facility: "", pairingIssue: "time-without-named-stream" }]),
  [{ staffName: "Example Doctor", date: "2026-08-06", start: "14:30", end: "00:00", reason: "time without named stream" }],
  "creator exception exports should contain only the fields needed to cross-check ambiguous time rows",
);
const kimOfficeRows = extractShiftRows([
  { staffId: "office-worker", facilityId: null, date: "2026-08-06", firstName: "Kim", lastName: "Whelan", payrollId: null, occurrences: 1, shift: "08:00-17:30" },
]);
assert.deepEqual(
  kimOfficeRows.map((row) => ({ label: row.label, start: row.start, end: row.end, facility: row.facility })),
  [{ label: "CS", start: "08:00", end: "17:30", facility: "" }],
  "Kim Whelan's verified time-only office shifts should be classified as CS",
);
assert.doesNotThrow(
  () => assertFindmyshiftDandenongAssignments(kimOfficeRows),
  "Kim Whelan's verified CS rule should resolve his time-only FindMyShift rows",
);
const shankarSupportedRows = extractShiftRows([
  { staffId: "supported-worker", facilityId: null, date: "2026-08-12", firstName: "Shankar", lastName: "Thapaliya", payrollId: null, occurrences: 1, shift: "08:00-17:30" },
]);
assert.deepEqual(
  shankarSupportedRows.map((row) => ({ label: row.label, start: row.start, end: row.end, facility: row.facility })),
  [{ label: "Paired AM", start: "08:00", end: "17:30", facility: "" }],
  "Shankar Thapaliya's paired time-only shifts should not be assigned to the paired clinician's stream",
);
assert.doesNotThrow(
  () => assertFindmyshiftDandenongAssignments(shankarSupportedRows),
  "Shankar Thapaliya's verified paired AM allocation should resolve without guessing a DDH stream",
);
const approvedDdhExceptionRows = extractShiftRows([
  { staffId: "liseth", facilityId: null, date: "2026-08-04", firstName: "Liseth", lastName: "Jalabe", payrollId: null, occurrences: 1, shift: "08:00-17:30" },
  { staffId: "stella", facilityId: null, date: "2026-08-07", firstName: "Stella", lastName: "Tran", payrollId: null, occurrences: 1, shift: "08:00-17:30" },
  { staffId: "di", facilityId: null, date: "2026-08-13", firstName: "Di", lastName: "Flood", payrollId: null, occurrences: 1, shift: "14:30-00:00" },
]);
assert.deepEqual(
  approvedDdhExceptionRows.map((row) => ({ name: row.name, date: row.date, label: row.label, start: row.start, end: row.end, facility: row.facility })),
  [
    { name: "Liseth Jalabe", date: "2026-08-04", label: "Paired AM", start: "08:00", end: "17:30", facility: "" },
    { name: "Stella Tran", date: "2026-08-07", label: "Paired AM", start: "08:00", end: "17:30", facility: "" },
    { name: "Di Flood", date: "2026-08-13", label: "S/L", start: "14:30", end: "00:00", facility: "" },
  ],
  "approved date-scoped DDH exceptions should resolve without widening the stream-assignment rule",
);
assert.doesNotThrow(
  () => assertFindmyshiftDandenongAssignments(approvedDdhExceptionRows),
  "approved date-scoped DDH exceptions should satisfy the stream-completeness gate",
);
assert.deepEqual(
  findmyshiftRows.map((row) => ({ date: row.date, label: row.label, start: row.start, end: row.end, facility: row.facility, seniority: row.seniority, comment: row.comment })),
  [
    { date: "2026-08-03", label: "North AM", start: "07:00", end: "15:00", facility: "North Campus", seniority: "Senior", comment: "" },
    { date: "2026-08-04", label: "South Night", start: "19:00", end: "07:00", facility: "South Campus", seniority: "Senior", comment: "" },
    { date: "2026-08-05", label: "Annual leave", start: "", end: "", facility: "", seniority: "Registrar", comment: "Approved leave" },
    { date: "2026-08-06", label: "North CS", start: "08:00", end: "16:00", facility: "North Campus", seniority: "Senior", comment: "" },
    { date: "2026-08-06", label: "CS", start: "", end: "", facility: "", seniority: "Senior", comment: "" },
  ],
  "FindMyShift parser should pair timed stream rows and preserve overnight, all-day, multi-facility, comment and seniority data",
);
assert.equal(
  findmyshiftRows.filter((row) => row.label === "Shift" && row.start && row.end).length,
  0,
  "a paired FindMyShift stream must not leave a duplicate generic timed event behind",
);
const findmyshiftWorkbook = XLSX.read(findmyshiftRowsWorkbook(findmyshiftRows), { type: "array", cellDates: true });
const findmyshiftDetails = XLSX.utils.sheet_to_json(findmyshiftWorkbook.Sheets["FindMyShift details"], { header: 1, blankrows: false });
assert.deepEqual(
  findmyshiftDetails[0],
  ["Staff ID", "Staff name", "Seniority/job title", "Date", "Shift label", "Start", "End", "Facility", "Comment"],
  "FindMyShift retained workbook should include its complete structured audit sheet",
);
assert.equal(findmyshiftDetails.length, 6, "FindMyShift audit sheet should retain each valid unique paired source entry");
assert.equal(findmyshiftDetails[2][7], "South Campus", "FindMyShift audit sheet should preserve the resolved facility");
assert.equal(findmyshiftDetails[3][8], "Approved leave", "FindMyShift audit sheet should preserve comments when the API supplies them");
const findmyshiftFormData = new FormData();
findmyshiftFormData.append("rosterFiles", workbookFile(findmyshiftWorkbook, "Dandenong-FindMyShift-fixture.xlsx"));
const findmyshiftUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: findmyshiftFormData }));
const findmyshiftSource = findmyshiftUpload.sources.ddh[0];
const ddhCanonicalLabels = ["Orange PM (on-call)", "SSU SMS", "Orange IC", "Silver IC", "PM FAST IC", "Paired AM"];
const ddhCanonicalTitles = ["DDH: Orange PM", "DDH: SSU", "DDH: Orange", "DDH: Silver", "DDH: FAST PM", "DDH: Paired AM"];
const ddhCanonicalManualWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(ddhCanonicalManualWorkbook, XLSX.utils.aoa_to_sheet([
  ["", "Mon. Aug. 03, 2026", "Tue. Aug. 04, 2026", "Wed. Aug. 05, 2026", "Thu. Aug. 06, 2026", "Fri. Aug. 07, 2026", "Sat. Aug. 08, 2026", "Sun. Aug. 09, 2026"],
  ["SENIOR MEDICAL STAFF", "", "", "", "", "", "", ""],
  ["Canonical DDH Doctor", ...ddhCanonicalLabels, "", ""],
  ["", "08:00-17:00", "08:00-17:00", "08:00-17:00", "08:00-17:00", "08:00-17:00", "", ""],
]), "Sheet1");
const ddhCanonicalAutoWorkbook = XLSX.read(findmyshiftRowsWorkbook(ddhCanonicalLabels.map((label, index) => ({
  name: "Canonical DDH Doctor",
  seniority: "SMS",
  date: `2026-08-0${index + 3}`,
  label,
  start: "08:00",
  end: "17:00",
  facility: "",
  comment: "",
}))), { type: "array", cellDates: true });
const ddhCanonicalAutoFormData = new FormData();
ddhCanonicalAutoFormData.append("rosterFiles", workbookFile(ddhCanonicalAutoWorkbook, "Dandenong-FindMyShift-canonical-labels.xlsx"));
const ddhCanonicalAutoUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: ddhCanonicalAutoFormData }));
assert.deepEqual(
  buildRosterView([], ddhCanonicalManualWorkbook, "CANONICAL DDH DOCTOR").events.map((event) => event.title),
  ddhCanonicalTitles,
  "manual DDH shift-code normalization should retain the approved canonical labels",
);
assert.deepEqual(
  buildRosterView([], ddhCanonicalAutoUpload.sources.ddh, "CANONICAL DDH DOCTOR").events.map((event) => event.title),
  ddhCanonicalTitles,
  "timed FindMyShift shifts should use the same DDH shift-code normalization as manual rosters",
);
const ddhWeeklyLeaveWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(ddhWeeklyLeaveWorkbook, XLSX.utils.aoa_to_sheet([
  ["", "Mon. Aug. 03, 2026", "Tue. Aug. 04, 2026", "Wed. Aug. 05, 2026", "Thu. Aug. 06, 2026", "Fri. Aug. 07, 2026", "Sat. Aug. 08, 2026", "Sun. Aug. 09, 2026"],
  ["SENIOR MEDICAL STAFF", "", "", "", "", "", "", ""],
  ["Monday Leave Doctor", "Annual Leave", "", "", "", "", "", ""],
  ["Friday Leave Doctor", "", "", "", "", "Annual Leave", "", ""],
  ["Mixed Leave Doctor", "Annual Leave", "", "CS", "", "", "", ""],
]), "Sheet1");
const mondayLeaveEvent = buildRosterView([], ddhWeeklyLeaveWorkbook, "MONDAY LEAVE DOCTOR").events.find((event) => event.title === "Annual Leave");
assert.ok(mondayLeaveEvent);
assert.deepEqual([mondayLeaveEvent.start, mondayLeaveEvent.end], ["2026-08-03", "2026-08-10"], "a Monday-only leave marker should cover the full roster week");
const fridayLeaveEvent = buildRosterView([], ddhWeeklyLeaveWorkbook, "FRIDAY LEAVE DOCTOR").events.find((event) => event.title === "Annual Leave");
assert.ok(fridayLeaveEvent);
assert.deepEqual([fridayLeaveEvent.start, fridayLeaveEvent.end], ["2026-08-07", "2026-08-08"], "a non-Monday leave marker must remain a one-day event");
const mixedLeaveEvents = buildRosterView([], ddhWeeklyLeaveWorkbook, "MIXED LEAVE DOCTOR").events;
assert.ok(mixedLeaveEvents.some((event) => event.title === "Annual Leave" && event.start === "2026-08-03" && event.end === "2026-08-04"), "Monday leave with another shift must remain a one-day event");
assert.ok(mixedLeaveEvents.some((event) => event.title === "DDH: CS" && event.start === "2026-08-05"), "a substantive allocation must remain alongside Monday leave");
const ddhCanonicalAutomatedPayload = await buildAutomatedDerivedRosterPayload({
  file: workbookFile(ddhCanonicalAutoWorkbook, "Dandenong-FindMyShift-canonical-labels.xlsx"),
  sourceId: "dandenong-findmyshift",
  contentHash: "findmyshift-canonical-labels-content-hash",
  providerVersion: "2026-08-11T00:00:00.000Z",
});
assert.deepEqual(
  ddhCanonicalAutomatedPayload.eventsByDoctor["CANONICAL DDH DOCTOR"].map((event) => event.title),
  ddhCanonicalTitles,
  "automated FindMyShift processing should retain the canonical DDH shift-code labels",
);
const unknownInternWorkbook = XLSX.read(findmyshiftRowsWorkbook([{
  name: "Pranay Pius",
  seniority: "Unknown",
  date: "2026-08-07",
  label: "INTERN SSU AM",
  start: "07:30",
  end: "17:00",
  facility: "INTERN SSU AM",
  comment: "",
}]), { type: "array", cellDates: true });
const unknownInternFormData = new FormData();
unknownInternFormData.append("rosterFiles", workbookFile(unknownInternWorkbook, "Dandenong-FindMyShift-unknown-intern.xlsx"));
const unknownInternUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: unknownInternFormData }));
const unknownInternDoctor = doctorOptions([], unknownInternUpload.sources.ddh).find((doctor) => doctor.key === "PRANAY PIUS");
assert.ok(unknownInternDoctor, "FindMyShift fixture should expose the unknown-seniority intern");
const unknownInternEvent = buildRosterView([], unknownInternUpload.sources.ddh, unknownInternDoctor.key).events[0];
assert.equal(unknownInternEvent.seniority, "Intern", "FindMyShift labels should supply seniority when staff metadata says Unknown");
const edHmosWorkbook = XLSX.read(findmyshiftRowsWorkbook([{
  name: "Gideon Charin",
  seniority: "ED HMO's",
  date: "2026-08-18",
  label: "Night",
  start: "22:00",
  end: "08:00",
  facility: "Night",
  comment: "",
}]), { type: "array", cellDates: true });
const edHmosFormData = new FormData();
edHmosFormData.append("rosterFiles", workbookFile(edHmosWorkbook, "Dandenong-FindMyShift-ed-hmos.xlsx"));
const edHmosUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: edHmosFormData }));
const edHmosEvent = buildRosterView([], edHmosUpload.sources.ddh, "GIDEON CHARIN").events[0];
assert.equal(edHmosEvent.seniority, "HMO", "FindMyShift ED HMO's and HMO's roster group labels should normalise to HMO");
const findmyshiftLeaveWorkbook = XLSX.read(findmyshiftRowsWorkbook([
  { sourceStaffId: "leave-doctor", name: "Ananth Sundaralingam", seniority: "SMS", date: "2026-08-10", label: "SL MMC", start: "", end: "", facility: "", comment: "" },
  { sourceStaffId: "leave-doctor", name: "Ananth Sundaralingam", seniority: "SMS", date: "2026-08-11", label: "S/L", start: "14:30", end: "00:00", facility: "", comment: "" },
  { sourceStaffId: "leave-doctor", name: "Ananth Sundaralingam", seniority: "SMS", date: "2026-08-12", label: "Annual leave 19hrs", start: "", end: "", facility: "", comment: "" },
]), { type: "array", cellDates: true });
const findmyshiftLeaveFormData = new FormData();
findmyshiftLeaveFormData.append("rosterFiles", workbookFile(findmyshiftLeaveWorkbook, "Dandenong-FindMyShift-leave.xlsx"));
const findmyshiftLeaveUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: findmyshiftLeaveFormData }));
const findmyshiftLeaveEvents = buildRosterView([], findmyshiftLeaveUpload.sources.ddh, "ANANTH SUNDARALINGAM").events;
assert.deepEqual(
  findmyshiftLeaveEvents.map((event) => [event.title, event.start, event.end]),
  [
    ["Sabbatical", "2026-08-10", "2026-08-11"],
    ["Sick leave", "2026-08-11", "2026-08-12"],
    ["Annual Leave", "2026-08-12", "2026-08-13"],
  ],
  "structured FindMyShift imports must preserve SL sabbatical, timed S/L sick leave, and annotated annual leave as distinct one-day entries",
);
const findmyshiftLeaveAutomatedPayload = await buildAutomatedDerivedRosterPayload({
  file: workbookFile(findmyshiftLeaveWorkbook, "Dandenong-FindMyShift-leave.xlsx"),
  sourceId: "dandenong-findmyshift",
  contentHash: "findmyshift-leave-content-hash",
  providerVersion: "2026-08-12T00:00:00.000Z",
});
assert.deepEqual(
  findmyshiftLeaveAutomatedPayload.eventsByDoctor["ANANTH SUNDARALINGAM"].map((event) => event.title),
  ["Sabbatical", "Sick leave", "Annual Leave"],
  "automated server-side parsing must use the same leave labels as manual/browser parsing",
);
const findmyshiftCrossTermRows = extractShiftRows([
  ...findmyshiftFixture.report,
  { "staffId": "staff-001", "facilityId": "facility-north", "date": "2026-10-05", "firstName": "Alex", "lastName": "Example", "payrollId": null, "occurrences": 1, "shift": "07:00-15:00" },
], { staff: findmyshiftFixture.staff, facilities: findmyshiftFixture.facilities });
const findmyshiftCrossTermFormData = new FormData();
findmyshiftCrossTermFormData.append("rosterFiles", workbookFile(XLSX.read(findmyshiftRowsWorkbook(findmyshiftCrossTermRows), { type: "array", cellDates: true }), "Dandenong-FindMyShift-full-range.xlsx"));
const findmyshiftCrossTermUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: findmyshiftCrossTermFormData }));
assert.equal(findmyshiftCrossTermUpload.sources.ddh.length, 1, "a FindMyShift full available roster may safely cross term boundaries");
const findmyshiftDoctors = doctorOptions([], [findmyshiftSource]);
const findmyshiftAlex = findmyshiftDoctors.find((doctor) => doctor.key === "ALEX EXAMPLE");
assert.ok(findmyshiftAlex, "FindMyShift workbook should expose its doctor names to the DDH parser");
const findmyshiftAlexEvents = buildRosterView([], [findmyshiftSource], findmyshiftAlex.key).events;
assert.ok(findmyshiftAlexEvents.some((event) => event.start === "2026-08-03T07:00:00+10:00" && event.end === "2026-08-03T15:00:00+10:00"), "FindMyShift timed shifts should reach the DDH calendar parser");
assert.ok(findmyshiftAlexEvents.some((event) => event.start === "2026-08-04T19:00:00+10:00" && event.end === "2026-08-05T07:00:00+10:00"), "FindMyShift overnight shifts should retain their next-day end time");
assert.ok(findmyshiftAlexEvents.some((event) => event.allDay && event.title === "DDH: CS"), "an unstreamed CS row should remain one all-day DDH event");
const findmyshiftBlair = findmyshiftDoctors.find((doctor) => doctor.key === "BLAIR EXAMPLE");
const findmyshiftBlairEvents = buildRosterView([], [findmyshiftSource], findmyshiftBlair.key).events;
assert.ok(findmyshiftBlairEvents.some((event) => event.title === "Annual Leave" && event.start === "2026-08-05" && event.end === "2026-08-06"), "FindMyShift one-day leave must not expand to the whole week");
assert.ok(findmyshiftAlexEvents.some((event) => event.location === "North Campus"), "FindMyShift facilities should reach the calendar event location");
assert.ok(findmyshiftAlexEvents.some((event) => event.title.includes("North AM")), "FindMyShift timed rows should preserve the paired stream in the title");
const findmyshiftAutomatedPayload = await buildAutomatedDerivedRosterPayload({
  file: workbookFile(findmyshiftWorkbook, "Dandenong-FindMyShift-fixture.xlsx"),
  sourceId: "dandenong-findmyshift",
  contentHash: "findmyshift-fixture-content-hash",
  providerVersion: "2026-08-06T06:39:00.000Z",
});
assert.equal(findmyshiftAutomatedPayload.eventCount, 5, "automated FindMyShift processing should use the structured Dandenong parser");
assert.ok(
  Object.values(findmyshiftAutomatedPayload.eventsByDoctor).flat().some((event) => event.location === "South Campus"),
  "automated FindMyShift processing should retain facility-specific calendar locations",
);

const ddhVariableLineStaff = [
  ["shemma", "Shemma", "Hasanovic", "SMS"],
  ["rajan", "Rajan", "Kailainathan", "SMS"],
  ["nagendran", "Nagendran", "Mathavan", "SMS"],
  ["mina", "Mina", "Nessim", "SMS"],
  ["igor", "Igor", "Tulchinsky", "SMS"],
  ["aditya", "Aditya", "Mehta", "Junior Registrar"],
  ["buthpitiya", "Buthpitiya", "Buthpitiya", "Intern"],
  ["khue", "Khue Dong Huynh", "Le", "Intern"],
].map(([staffId, firstName, lastName, jobTitle], order) => ({ staffId, firstName, lastName, jobTitle, order }));
const ddhVariableLineRows = [
  ["shemma", "Shemma Hasanovic", "SMS", "Orange AM IC", "08:00", "18:00"],
  ["rajan", "Rajan Kailainathan", "SMS", "SSU SMS", "07:30", "17:30"],
  ["nagendran", "Nagendran Mathavan", "SMS", "Silver AM IC", "08:00", "18:00"],
  ["mina", "Mina Nessim", "SMS", "Rover AM", "08:00", "18:00"],
  ["igor", "Igor Tulchinsky", "SMS", "AM Fast (3)", "10:00", "17:00"],
  ["aditya", "Aditya Mehta", "Junior Registrar", "AM Fast IC", "08:00", "17:30"],
  ["buthpitiya", "Buthpitiya Buthpitiya", "Intern", "INTERN SSU AM", "07:30", "17:00"],
  ["khue", "Khue Dong Huynh Le", "Intern", "Orange AM4", "08:00", "17:30"],
].map(([sourceStaffId, name, seniority, label, start, end]) => ({
  sourceStaffId,
  name,
  seniority,
  date: "2026-08-25",
  label,
  start,
  end,
  facility: label,
  comment: "",
}));
const ddhVariableLineWorkbook = XLSX.read(findmyshiftRowsWorkbook(ddhVariableLineRows, ddhVariableLineStaff), { type: "array", cellDates: true });
const ddhVariableLinePayload = await buildAutomatedDerivedRosterPayload({
  file: workbookFile(ddhVariableLineWorkbook, "Dandenong-FindMyShift-variable-lines.xlsx"),
  sourceId: "dandenong-findmyshift",
  contentHash: "findmyshift-variable-lines-content-hash",
  providerVersion: "2026-08-25T00:00:00.000Z",
  parserExtensions: { mmc: [], ddh: [], casey: [], mch: [] },
});
assert.deepEqual(
  ddhVariableLineRows.map((expected) => {
    const key = expected.name.toUpperCase();
    const doctor = ddhVariableLinePayload.doctors.find((candidate) => candidate.key === key);
    const event = ddhVariableLinePayload.eventsByDoctor[key]?.find((candidate) => candidate.start.startsWith("2026-08-25"));
    return {
      name: expected.name,
      seniority: doctor?.seniority,
      eventSeniority: event?.seniority,
      providerStaffId: doctor?.providerStaffId,
      allocation: event?.rawValue,
    };
  }),
  ddhVariableLineRows.map((expected) => ({
    name: expected.name,
    seniority: expected.seniority,
    eventSeniority: expected.seniority,
    providerStaffId: expected.sourceStaffId,
    allocation: expected.label,
  })),
  "structured FindMyShift details must preserve DDH grades and allocations when the visual roster uses variable one-, two-, and three-line cells",
);

const mmcTypoWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(mmcTypoWorkbook, XLSX.utils.aoa_to_sheet([
  ["Seniority", "Name", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
  ["", "", new Date("2026-03-23T00:00:00"), new Date("2026-03-24T00:00:00"), new Date("2026-03-25T00:00:00"), new Date("2026-03-26T00:00:00"), new Date("2026-03-27T00:00:00"), new Date("2026-03-28T00:00:00"), new Date("2026-03-29T00:00:00")],
  ["SMS", "MMC Typo Doctor", "1000-17300 SWA", "1430-0000 SWP PH", "1500-0000 PH", "", "", "", ""],
  ["SMS", "Christina Hatton", "", "", "", "i", "", "", ""],
], { cellDates: true }), "Week 1");
XLSX.utils.book_append_sheet(mmcTypoWorkbook, XLSX.utils.aoa_to_sheet([["MMC test roster"]]), "Whole thing");
const mmcTypoFormData = new FormData();
mmcTypoFormData.append("rosterFiles", workbookFile(mmcTypoWorkbook, "AdultTerm1.2026.xlsx"));
const mmcTypoSource = (await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: mmcTypoFormData }))).sources.mmc;
const mmcTypoDoctor = doctorOptions(mmcTypoSource, []).find((doctor) => doctor.key === "MMC TYPO DOCTOR");
const christinaHatton = doctorOptions(mmcTypoSource, []).find((doctor) => doctor.key === "CHRISTINA HATTON");
assert.ok(mmcTypoDoctor && christinaHatton, "MMC typo fixtures should expose both clinicians");
assert.deepEqual(
  buildRosterView(mmcTypoSource, [], mmcTypoDoctor.key).events.map((event) => [event.title, event.rawValue, event.timeLabel]),
  [
    ["MMC: Swing AM", "1000-17300 SWA", "10:00-17:30"],
    ["MMC: Swing PM Hub", "1430-0000 SWP PH", "14:30-00:00"],
    ["MMC: Hub PM", "1500-0000 PH", "15:00-00:00"],
  ],
  "MMC swing and Hub typo variants should retain their supplied times and normalised titles",
);
assert.deepEqual(buildRosterView(mmcTypoSource, [], christinaHatton.key).events, [], "a stray MMC i annotation should not become an event");
assert.deepEqual(buildRosterView(mmcTypoSource, [], christinaHatton.key).issues, [], "a stray MMC i annotation should not remain unresolved");

const ddhMessageWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(ddhMessageWorkbook, XLSX.utils.aoa_to_sheet([
  ["", "Mon. Sep. 14, 2026", "Tue. Sep. 15, 2026", "Wed. Sep. 16, 2026", "Thu. Sep. 17, 2026", "Fri. Sep. 18, 2026", "Sat. Sep. 19, 2026", "Sun. Sep. 20, 2026"],
  ["Shawn Test", "", "", "", "", "", "08H00", "08H00"],
  ["", "", "", "", "", "", "VHH", "VHH"],
  ["Message Test", "AM (AVOID IF POSSIBLE)", "AM OK", "C/S for 27/4", "Can work", "Can work 1 extra this week", "Cant do this weekend, sorry!", "4 shifts this week to make up for next week pls"],
]), "Sheet1");
const ddhMessageFormData = new FormData();
ddhMessageFormData.append("rosterFiles", workbookFile(ddhMessageWorkbook, "Dandenong messages.xlsx"));
const ddhMessageSource = (await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: ddhMessageFormData }))).sources.ddh;
const shawnTest = doctorOptions([], ddhMessageSource).find((doctor) => doctor.key === "SHAWN TEST");
const messageTest = doctorOptions([], ddhMessageSource).find((doctor) => doctor.key === "MESSAGE TEST");
assert.ok(shawnTest && messageTest, "DDH annotation fixtures should expose their clinicians");
assert.deepEqual(buildRosterView([], ddhMessageSource, shawnTest.key).events, [], "two-line 08H00/VHH late–early warnings should not become DDH shifts");
assert.deepEqual(buildRosterView([], ddhMessageSource, shawnTest.key).issues, [], "two-line 08H00/VHH late–early warnings should not remain unresolved");
assert.deepEqual(buildRosterView([], ddhMessageSource, messageTest.key).events, [], "DDH roster-writer messages should not become calendar shifts");
assert.deepEqual(buildRosterView([], ddhMessageSource, messageTest.key).issues, [], "DDH roster-writer messages should not remain unresolved");

const caseyWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/Casey_Term_2_2026_DRAFT.xlsm", import.meta.url)), {
  cellDates: true,
});
const mchWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/Paeds_Term_2_2026.xlsx", import.meta.url)), {
  cellDates: true,
});
const caseyBytes = await readFile(fileURLToPath(new URL("../fixtures/Casey_Term_2_2026_DRAFT.xlsm", import.meta.url)));
const mchBytes = await readFile(fileURLToPath(new URL("../fixtures/Paeds_Term_2_2026.xlsx", import.meta.url)));
const appSource = await readFile(new URL("../public/static/app.js", import.meta.url), "utf8");
assert.doesNotMatch(appSource, /Multiple grades recorded/, "ED staff should show one effective grade rather than exposing conflicting source grades");
assert.match(appSource, /eventSeniorityRoleCode\(event\)[\s\S]*facilityOverviewDetectedSeniority/, "Who should infer a staff grade from an explicit roster title when the provider omits seniority");
const rosterSource = await readFile(new URL("../public/static/roster.js", import.meta.url), "utf8");
const stateSource = await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8");
const findmyshiftModuleSource = await readFile(new URL("../functions/_lib/findmyshift.js", import.meta.url), "utf8");
const styleSource = await readFile(new URL("../public/static/styles.css", import.meta.url), "utf8");
const calendarMigrationSource = await readFile(new URL("../migrations/0001_calendar_store.sql", import.meta.url), "utf8");
const insightIndexMigrationSource = await readFile(new URL("../migrations/0005_roster_insight_index.sql", import.meta.url), "utf8");
const facilityAccessMigrationSource = await readFile(new URL("../migrations/0011_facility_overview_access.sql", import.meta.url), "utf8");
const facilityOptInRepairMigrationSource = await readFile(new URL("../migrations/0017_restore_facility_overview_opt_in.sql", import.meta.url), "utf8");
const safeStagedActivationMigrationSource = await readFile(new URL("../migrations/0019_safe_staged_roster_activation.sql", import.meta.url), "utf8");
const d1CalendarSource = await readFile(new URL("../functions/_lib/d1-calendar.js", import.meta.url), "utf8");
const stateApiSource = await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8");
const automationIngestSource = await readFile(new URL("../functions/api/automation/ingest.js", import.meta.url), "utf8");
const automationDerivedSource = await readFile(new URL("../functions/api/automation/derived.js", import.meta.url), "utf8");
const findmyshiftCheckSource = await readFile(new URL("../functions/api/automation/findmyshift-check.js", import.meta.url), "utf8");
const automationDispatchSource = await readFile(new URL("../functions/_lib/automation-dispatch.js", import.meta.url), "utf8");
const automationDispatchEndpointSource = await readFile(new URL("../functions/api/automation/dispatch.js", import.meta.url), "utf8");
const automationWorkflowSource = await readFile(new URL("../.github/workflows/monash-roster-sync.yml", import.meta.url), "utf8");
const indexSource = await readFile(new URL("../public/index.html", import.meta.url), "utf8");
assert.match(indexSource, /id="stayLoggedIn"[^>]*checked/, "Stay logged in should be selected by default");
assert.equal((appSource.match(/data-test-findmyshift/g) || []).length, 0, "FindMyShift diagnostics should not remain exposed as a UI control");
assert.equal((appSource.match(/data-sync-findmyshift/g) || []).length, 0, "FindMyShift is automated and should not expose a manual sync control");
assert.equal((appSource.match(/data-download-findmyshift-exceptions/g) || []).length, 0, "FindMyShift exception review is no longer exposed as a UI control");
assert.match(findmyshiftModuleSource, /findmyshiftRequest\("reports\/shifts",[\s\S]*comments:\s*"no"/, "FindMyShift imports should request comments=no so free-text roster comments are excluded upstream");
assert.match(findmyshiftCheckSource, /!force && current\?\.providerVersion === providerVersion && isIncompleteDandenongAssignmentError/, "a creator-forced FindMyShift refresh must retry a cached incomplete provider response");
assert.match(facilityAccessMigrationSource, /facility_overview_enabled INTEGER NOT NULL DEFAULT 0/, "At a glance database access should default to opt-in");
assert.match(facilityOptInRepairMigrationSource, /WHEN role IN \('creator', 'owner'\) THEN 1[\s\S]*ELSE 0/, "the At a glance repair should retain Creator access and revoke unintended standard-user access");
assert.match(indexSource, /data-facility-overview-tab="on-shift">On shift[\s\S]*data-facility-overview-tab="staff">ED staff[\s\S]*data-facility-overview-tab="together">Working together[\s\S]*data-facility-overview-tab="by-stream">By stream/, "By stream should be the final At a glance tab");
assert.match(appSource, /activeTab === "by-stream"[\s\S]*From[\s\S]*To[\s\S]*function renderFacilityOverviewByStream[\s\S]*Add another stream[\s\S]*Hide dates without assignments/, "By stream should render date controls, selectable stream rows, and the sparse-range control");
assert.match(appSource, /data-facility-overview-by-stream-date="from"[\s\S]*data-facility-overview-by-stream-date="to"[\s\S]*preview-today-button[\s\S]*setFacilityOverviewByStreamRange/, "By stream banner should expose the shared date range and Today action");
assert.match(appSource, /facility-overview-by-stream-today[\s\S]*renderFacilityOverviewDateNavigation\("range"\)/, "By stream's At a glance toolbar should include its own Previous and Next controls beside Today");
assert.match(styleSource, /facility-overview-controls \.facility-overview-date-actions \{\s*display: inline-flex;[\s\S]*@media \(max-width: 1200px\)[\s\S]*facility-overview-controls \.facility-overview-date-navigation-label \{\s*display: none;[\s\S]*facility-overview-controls \.facility-overview-date-navigation-chevron \{\s*display: inline;/, "At a glance toolbar navigation should become chevrons when the viewport narrows");
assert.match(appSource, /function setFacilityOverviewByStreamRange[\s\S]*nextFrom > nextTo[\s\S]*facilityOverviewState\.byStreamFrom = nextFrom;[\s\S]*facilityOverviewState\.byStreamTo = nextTo;[\s\S]*loadFacilityOverviewByStream/, "By stream dates should clamp and refresh through one shared mutation path");
assert.match(appSource, /facilityOverviewState\.tab === "by-stream"[\s\S]*setFacilityOverviewByStreamRange\(\{ from: today, to: today \}\)/, "By stream Today should reset both range bounds through the shared mutation path");
assert.match(appSource, /FACILITY_OVERVIEW_SENIORITY_ORDER = \["SMS", "Senior Registrar", "CMO", "Transitional\/Intermediate Registrar", "Junior Registrar", "HMO", "NP", "Physio", "Intern", "Unknown"\]/, "At a glance should use the approved seniority hierarchy");
assert.match(appSource, /function compareFacilityOverviewSeniorities[\s\S]*function compareFacilityOverviewPeople[\s\S]*function compareFacilityOverviewAssignmentsBySeniority/, "At a glance should use shared seniority comparators");
assert.match(appSource, /renderFacilityOverviewOnShiftNames[\s\S]*compareFacilityOverviewPeople/, "On shift stream tiles should sort staff by hierarchy before name");
assert.match(appSource, /function newFacilityOverviewByStreamRow[\s\S]*seniority: options\.seniority \|\| "ALL"[\s\S]*if \(!row\.seniority\) row\.seniority = "ALL"/, "By stream should default new and repaired selections to All team");
assert.match(appSource, /function facilityOverviewByStreamShiftBlocks[\s\S]*normalStartMinutes[\s\S]*isExceptional[\s\S]*compareFacilityOverviewAssignmentsBySeniority/, "By stream should order shift blocks by time and people within them by seniority");
assert.match(appSource, /facility-overview-by-stream-grade[\s\S]*function facilityOverviewSeniorityAbbreviation[\s\S]*"Senior Registrar": "SR"[\s\S]*"Junior Registrar": "JR"/, "By stream should show abbreviated seniority beside each name");
assert.match(appSource, /function facilityOverviewAssignmentForRangeRow[\s\S]*buildWhoAssignment[\s\S]*facilityOverviewIsMeaningfulStream/, "By stream should reuse the On shift assignment and stream classifier");
assert.match(appSource, /function facilityOverviewPreferredStreamKey[\s\S]*buildWhoAssignments[\s\S]*active[\s\S]*next/, "By stream should prefer the viewer's active or next stream before a catalogue fallback");
assert.match(appSource, /function facilityOverviewPreferredFacilityFromEvents[\s\S]*sole-current-week-facility[\s\S]*active-shift[\s\S]*today-next-shift[\s\S]*next-shift/, "At a glance should choose the preferred ED from the already-loaded calendar events");
assert.match(appSource, /function openFacilityOverview[\s\S]*refreshFacilityOverviewPreferredFacility[\s\S]*facilityOverviewSection\?\.classList\.remove\("hidden"\)[\s\S]*loadFacilityOverviewOnShift/, "At a glance should render before loading its On shift request");
assert.match(appSource, /function toggleFacilityOverview[\s\S]*facilityOverviewNavigationLocked[\s\S]*openFacilityOverview/, "At a glance navigation should ignore duplicate opening clicks");
assert.match(appSource, /function openFacilityOverviewByStream[\s\S]*Loading available streams[\s\S]*loadFacilityOverviewMetadata[\s\S]*loadFacilityOverviewByStream/, "By stream should fetch its catalogue only when that tab is opened");
assert.match(stateSource, /action === "queryFacilityOverviewMetadata"[\s\S]*const catalog = await queryFacilityOverviewRange[\s\S]*action === "queryFacilityOverviewByStream"/, "The metadata API should fetch only the lazy By stream catalogue");
assert.match(findmyshiftCheckSource, /const fileName = `Dandenong-FindMyShift-\$\{IMPORT_FORMAT\}[\s\S]*currentFormatRun\?\.status === "success"/, "FindMyShift should only call an import unchanged after this parser-format file has succeeded");
assert.match(automationIngestSource, /findSuccessfulRosterSyncByHash\([^\n]*file\.name\)[\s\S]*findQueuedRosterSyncByHash\([^\n]*file\.name\)/, "automation ingestion should include the retained filename when deduplicating an identical workbook");
assert.match(automationDerivedSource, /\["parser-rule", "creator-reprocess"\]\.includes\(run\.triggerType\)[\s\S]*activeFileId: preserveActiveFile \? existing\.activeFileId : run\.fileId/, "historical and creator-requested reparses should not replace an automated source's active-file pointer");
assert.match(findmyshiftCheckSource, /requestBody[\s\S]*force[\s\S]*queueCurrentFindmyshiftReprocess[\s\S]*status: "reprocess-queued"/, "a creator refresh should reprocess the retained FindMyShift file when its provider version is unchanged");
assert.match(stateSource, /action === "refreshAutomatedRosterSource"[\s\S]*source\.provider === "findmyshift"[\s\S]*force: true[\s\S]*queueAutomatedSourceReprocess/, "the auto-sync refresh action should check FindMyShift remotely and reprocess retained push-only sources");
const facilityOverviewEventHelpers = appSource.match(/function eventRosterDateKey[\s\S]*?(?=\nfunction filterWhenInsightEvents)/)?.[0] || "";
const facilityOverviewDateHelpers = appSource.match(/function parseDateOnly[\s\S]*?(?=\nfunction formatLongDate)/)?.[0] || "";
const facilityOverviewPreferredHelper = appSource.match(/function facilityOverviewMelbourneClock[\s\S]*?(?=\nfunction refreshFacilityOverviewPreferredFacility)/)?.[0] || "";
assert.ok(facilityOverviewEventHelpers && facilityOverviewDateHelpers && facilityOverviewPreferredHelper, "At a glance preferred-ED helper dependencies should be available for behavioural tests");
const resolvePreferredFacility = new Function(`${facilityOverviewEventHelpers}\n${facilityOverviewDateHelpers}\n${facilityOverviewPreferredHelper}\nreturn facilityOverviewPreferredFacilityFromEvents;`)();
const facilityOverviewScrollLatchHelpers = appSource.match(/function facilityOverviewScrollerHasOverflow[\s\S]*?(?=\nfunction setFacilityOverviewCompactMode)/)?.[0] || "";
assert.ok(facilityOverviewScrollLatchHelpers, "At a glance scroll-latch helpers should be available for behavioural tests");
const { facilityOverviewShouldCompact, facilityOverviewShouldReleaseCompact } = new Function(`const FACILITY_OVERVIEW_COMPACT_SCROLL_THRESHOLD = 28; const FACILITY_OVERVIEW_SCROLL_TOLERANCE = 0;\n${facilityOverviewScrollLatchHelpers}\nreturn { facilityOverviewShouldCompact, facilityOverviewShouldReleaseCompact };`)();
assert.equal(
  facilityOverviewShouldCompact({ scrollHeight: 100, clientHeight: 100, scrollTop: 40 }, { userDirection: 1 }),
  false,
  "results that fit in the expanded viewport should not compact the header",
);
assert.equal(
  facilityOverviewShouldCompact({ scrollHeight: 101, clientHeight: 100, scrollTop: 40 }, { userDirection: 1 }),
  true,
  "any genuine expanded-viewport overflow should allow downward scrolling to compact the header",
);
assert.equal(
  facilityOverviewShouldCompact({ scrollHeight: 200, clientHeight: 100, scrollTop: 20 }, { userDirection: 1 }),
  false,
  "the compact header should retain its scroll activation threshold",
);
assert.equal(
  facilityOverviewShouldReleaseCompact({ scrollTop: 0 }, 1),
  false,
  "a layout-driven reset to the top after a downward scroll should remain latched compact",
);
assert.equal(
  facilityOverviewShouldReleaseCompact({ scrollTop: 0 }, -1),
  true,
  "an explicit upward action at the top should release the compact header",
);
const preferredShift = (source, start, end, title = `${source}: Shift`) => ({ source, title, start, end });
const preferredNow = new Date("2026-08-10T00:00:00Z"); // Monday 10:00 in Melbourne.
assert.deepEqual(
  resolvePreferredFacility([
    preferredShift("ddh", "2026-08-11T07:30:00", "2026-08-11T15:30:00"),
    preferredShift("ddh", "2026-08-13T07:30:00", "2026-08-13T15:30:00"),
  ], { today: "2026-08-10", now: preferredNow }),
  { facilityKey: "DDH", reason: "sole-current-week-facility", evidenceDate: "2026-08-11" },
  "one current-week facility should be selected without querying metadata",
);
assert.deepEqual(
  resolvePreferredFacility([
    preferredShift("mmc", "2026-08-11T07:30:00", "2026-08-11T15:30:00"),
    preferredShift("ddh", "2026-08-12T07:30:00", "2026-08-12T15:30:00"),
  ], { today: "2026-08-11", now: preferredNow }),
  { facilityKey: "MMC", reason: "active-shift", evidenceDate: "2026-08-11" },
  "an active shift should win when more than one ED is rostered this week",
);
assert.deepEqual(
  resolvePreferredFacility([
    preferredShift("mmc", "2026-08-10T15:00:00", "2026-08-10T23:00:00"),
    preferredShift("ddh", "2026-08-11T07:30:00", "2026-08-11T15:30:00"),
  ], { today: "2026-08-10", now: preferredNow }),
  { facilityKey: "MMC", reason: "today-next-shift", evidenceDate: "2026-08-10" },
  "a later shift today should win over a later shift this week",
);
assert.deepEqual(
  resolvePreferredFacility([
    preferredShift("ddh", "2026-08-12T07:30:00", "2026-08-12T15:30:00"),
    preferredShift("mmc", "2026-08-13T07:30:00", "2026-08-13T15:30:00"),
  ], { today: "2026-08-10", now: preferredNow }),
  { facilityKey: "DDH", reason: "next-shift", evidenceDate: "2026-08-12" },
  "the next shift should be selected when there is no rostered shift today",
);
assert.deepEqual(
  resolvePreferredFacility([
    preferredShift("ddh", "2026-08-17T07:30:00", "2026-08-17T15:30:00"),
  ], { today: "2026-08-10", now: preferredNow }),
  { facilityKey: "DDH", reason: "next-shift", evidenceDate: "2026-08-17" },
  "the next shift should be used when no shift occurs in the current week",
);
assert.deepEqual(
  resolvePreferredFacility([
    preferredShift("ddh", "2026-08-11T07:30:00", "2026-08-11T15:30:00", "DDH: Annual leave"),
  ], { today: "2026-08-10", now: preferredNow, linkedSourceTypes: ["mch"] }),
  { facilityKey: "MCH", reason: "sole-or-first-linked-facility", evidenceDate: "2026-08-10" },
  "leave should not determine the preferred ED",
);
assert.match(d1CalendarSource, /export async function queryFacilityOverviewRange[\s\S]*roster_events\.source_type IN[\s\S]*roster_events\.start_date >= \?/, "By stream should query the requested EDs and date range in one database operation");
assert.match(styleSource, /\.facility-overview-by-stream \{[\s\S]*grid-template-columns:[\s\S]*\.facility-overview-by-stream-selectors \{[\s\S]*position: sticky[\s\S]*@media \(max-width: 900px\)[\s\S]*\.facility-overview-by-stream \{[\s\S]*grid-template-columns: 1fr/, "By stream should use a desktop selector rail and stack it on narrow screens");
assert.match(appSource, /Each row is one result lane[\s\S]*selected\.length > 1[\s\S]*facility-overview-by-stream-comparison-day-grid[\s\S]*facility-overview-by-stream-comparison-head-grid/, "By stream should render aligned comparison lanes for multiple selected streams");
assert.match(appSource, /const previous = facilityOverviewState\.byStreamRows\.at\(-1\);[\s\S]*newFacilityOverviewByStreamRow\(previous\)/, "adding a By stream selection should copy the most recent selection");
assert.match(appSource, /function facilityOverviewByStreamDistinctRows[\s\S]*facilityOverviewByStreamContentFromData[\s\S]*const selected = facilityOverviewByStreamDistinctRows\(\)/, "duplicate By stream selections should remain editable but share one result lane");
assert.match(appSource, /const rows = facilityOverviewByStreamDistinctRows\(\)\.filter[\s\S]*selections: rows/, "By stream coverage queries should send only distinct selections");
assert.match(appSource, /function facilityOverviewByStreamDuplicateRows[\s\S]*Choose each ED, stream, and seniority combination only once[\s\S]*Duplicate selections are shown once/, "duplicate By stream selections should warn without removing unique result lanes");
assert.match(appSource, /facilityOverviewState\.byStreamRequestId \+= 1;[\s\S]*facilityOverviewByStreamContentFromData/, "switching back to a duplicate selection should invalidate stale stream requests and retain valid results");
assert.match(stateSource, /const uniqueSelections = \[\];[\s\S]*selectionKeys\.has\(key\)[\s\S]*selections: uniqueSelections/, "the By stream API should deduplicate repeated combinations defensively");
assert.match(styleSource, /grid-template-areas: "results selectors"[\s\S]*facility-overview-by-stream-selectors \{[\s\S]*grid-area: selectors[\s\S]*facility-overview-by-stream-results \{[\s\S]*grid-area: results/, "By stream results should appear left of the selector rail on desktop");
assert.match(styleSource, /facility-overview-by-stream-comparison-head-grid,[\s\S]*grid-template-columns: repeat\(auto-fit, minmax\(200px, 1fr\)\)[\s\S]*@media \(max-width: 900px\)[\s\S]*minmax\(150px, 1fr\)/, "By stream should fit several desktop comparison lanes and two readable mobile lanes when space permits");
assert.match(
  appSource.match(/function renderFacilityOverviewTogetherResults[\s\S]*?function facilityOverviewFormatOverlap/)?.[0] || "",
  /eventSourceCode[\s\S]*nextStart >= nextEnd[\s\S]*facilityOverviewSubtractIntervals/,
  "Working together should match true same-hospital shift intervals and remove all-staff time from pair-only matches",
);
assert.match(
  appSource.match(/function renderFacilityOverviewTogetherResults[\s\S]*?function facilityOverviewFormatOverlap/)?.[0] || "",
  /selectedDoctors\.length === 2[\s\S]*showGroups[\s\S]*All selected staff[\s\S]*Two-person overlaps/,
  "Working together should only add all-staff and pair headings when they are relevant to a selection of three or more",
);
assert.match(
  appSource.match(/function renderFacilityOverviewTogetherResults[\s\S]*?function facilityOverviewWorkingIntervals/)?.[0] || "",
  /selectedDoctors\.length === 1[\s\S]*No rostered shifts found[\s\S]*singlePerson: true/,
  "Working together should show a selected person's own shifts when only one staff member is chosen",
);
assert.match(appSource, /data-facility-overview-together-remove="\$\{index\}"[\s\S]*🗑/, "Every Working together staff row should have a compact remove control");
assert.match(appSource, /selectedCount >= 1[\s\S]*loadFacilityOverviewTogether/, "Choosing one Working together staff member should load their shifts");
assert.match(stateSource, /doctorKeys\.length < 1[\s\S]*Choose at least one staff member/, "The Working together API should allow a single staff member");
assert.doesNotMatch(appSource, /data-facility-overview-together-edit-staff/, "Working together should not render a redundant Edit staff control");
assert.match(styleSource, /#facilityOverviewSection\.is-compact[\s\S]*\.facility-overview-tabs/, "At a glance tabs should compact after scrolling");
assert.match(styleSource, /#facilityOverviewSection \{[\s\S]*?grid-template-rows: auto auto auto minmax\(0, 1fr\);[\s\S]*?overflow: hidden;/, "At a glance should keep its header stack outside the scroll container");
assert.match(styleSource, /#facilityOverviewBody \{[\s\S]*?overflow-y: auto;/, "At a glance results should be the sole desktop scroll container");
assert.match(appSource, /FACILITY_OVERVIEW_COMPACT_SCROLL_THRESHOLD = 28[\s\S]*?function facilityOverviewScrollerHasOverflow[\s\S]*?scrollHeight > scroller\.clientHeight[\s\S]*?function facilityOverviewShouldReleaseCompact[\s\S]*?userDirection < 0/, "At a glance should latch header compaction only for overflowing results and release it on explicit upward input");
assert.doesNotMatch(appSource, /facilityOverviewCompactReleaseTimer/, "At a glance should not use a delayed scroll-position release that can rebound after compaction");
assert.doesNotMatch(appSource, /pointerScroller|setFacilityOverviewScrollDirection\(scroller, movement\)/, "layout-driven scroll movement should never be mistaken for explicit upward input");
assert.match(appSource, /function facilityOverviewLabel\(\)[\s\S]*?\? "Director overview"[\s\S]*?: "At a glance"/, "Director accounts should call the shared overview Director overview while other accounts retain At a glance");
assert.match(appSource, /facilityOverviewButton\.textContent = open \? "My calendar" : label/, "The sidebar overview control should become My calendar while the overview is open");
assert.match(appSource, /addEventListener\("input"[\s\S]*?refreshFacilityOverviewStaffContent\(\);[\s\S]*?renderFacilityOverviewStaffBody\(\);/, "ED staff search should refresh results without replacing its focused input");
assert.match(
  appSource.match(/function renderFacilityOverviewOnShiftResults[\s\S]*?async function loadFacilityOverviewStaff/)?.[0] || "",
  /buildWhoAssignment[\s\S]*renderFacilityOverviewOnShiftPeriod[\s\S]*facilityOverviewIsMeaningfulStream[\s\S]*renderFacilityOverviewStreamCard[\s\S]*renderFacilityOverviewSeniorityLink[\s\S]*renderFacilityOverviewStaffName/,
  "On shift should group recognised streams with seniority subgroups and fall back to seniority for unstreamed staff",
);
const onShiftResultsSource = appSource.match(/function renderFacilityOverviewOnShiftResults[\s\S]*?function renderFacilityOverviewOnShiftPeriod/)?.[0] || "";
assert.match(onShiftResultsSource, /const doctorKey = String\(row\.doctorKey \|\| ""\)\.trim\(\);/, "On shift should retain each person's raw roster key for staff actions");
assert.match(onShiftResultsSource, /const groupKey = `\$\{String\(row\.sourceType/, "On shift may group matching names by source without changing the roster key");
assert.match(onShiftResultsSource, /people\.get\(groupKey\)[\s\S]*doctorKey,[\s\S]*people\.set\(groupKey, entry\);/, "On shift source grouping must not be passed into Working together as a doctor key");
assert.match(
  appSource.match(/function renderFacilityOverviewTogetherMatchCards[\s\S]*?function facilityOverviewFormatOverlap/)?.[0] || "",
  /canUseCreatorDoctorSwitcher\(\)[\s\S]*directCalendar: canOpenStaffCalendars[\s\S]*renderFacilityOverviewSeniorityLink\(person\.event\?\.seniority/,
  "Working together results should retain direct Creator calendar links and seniority drill-throughs",
);
assert.match(
  appSource.match(/function renderFacilityOverviewStaffName[\s\S]*?function renderFacilityOverviewSeniorityLink/)?.[0] || "",
  /data-facility-overview-open-working-together[\s\S]*data-facility-overview-staff-menu[\s\S]*renderFacilityOverviewStaffActionMenu[\s\S]*Person's calendar[\s\S]*When working together/,
  "At a glance staff names should left-click into Working together and retain a Creator-only context menu",
);
assert.match(
  appSource.match(/addEventListener\("contextmenu"[\s\S]*?addEventListener\("change"/)?.[0] || "",
  /data-facility-overview-staff-designation-menu[\s\S]*isViewingCreatorAccount\(\)[\s\S]*data-facility-overview-staff-menu[\s\S]*preventDefault\(\)[\s\S]*refreshFacilityOverviewStaffActionContent\(\)/,
  "Only the active Creator profile should receive staff and no-shift designation context menus",
);
assert.match(
  appSource.match(/document\.addEventListener\("pointerdown"[\s\S]*?facilityOverviewSection\?\.addEventListener\("scroll"/)?.[0] || "",
  /closeFacilityOverviewStaffActionMenu\(\)[\s\S]*event\.key !== "Escape"[\s\S]*closeFacilityOverviewStaffActionMenu\(\)/,
  "Outside clicks and Escape should dismiss a staff context menu",
);
assert.match(
  appSource.match(/function facilityOverviewTogetherStaffOptions[\s\S]*?function initializeFacilityOverviewTogetherState/)?.[0] || "",
  /availableRosterDoctors[\s\S]*doctorPickerOptions\(\)[\s\S]*activeViewer[\s\S]*togetherPinnedDoctors/,
  "Working together options should include roster staff, the active viewer, and pinned roster results",
);
assert.match(
  appSource.match(/function facilityOverviewTogetherFallbackOption[\s\S]*?function closeFacilityOverviewStaffActionMenu/)?.[0] || "",
  /facilityOverviewTogetherFallbackOption\(target\)[\s\S]*if \(!selectedPerson\) return;[\s\S]*togetherPinnedDoctors = viewer \? \[selectedPerson, viewer\] : \[selectedPerson\];[\s\S]*void loadFacilityOverviewTogether\(\)/,
  "Working together should always open for an authorised staff selection, prefill the clinical viewer when available, and load the relevant shifts",
);
assert.match(
  appSource.match(/function openFacilityOverviewWorkingTogether[\s\S]*?function closeFacilityOverviewStaffActionMenu/)?.[0] || "",
  /const activeDoctor = currentNonClinical \? null : selectedDoctor\(\);[\s\S]*const viewer = activeDoctor &&[\s\S]*togetherPinnedDoctors = viewer \? \[selectedPerson, viewer\] : \[selectedPerson\];/,
  "Non-clinical Directors should open Working together without requiring a personal calendar profile",
);
assert.match(appSource, /currentNonClinical && currentDirectorViewEnabled && canUseFacilityOverview\(\) && !isFacilityOverviewOpen\(\)/, "Only non-clinical Directors should open directly into Director overview; clinical accounts should retain their calendar");
assert.match(
  appSource.match(/function facilityOverviewAccountKey[\s\S]*?async function openFacilityOverview/)?.[0] || "",
  /FACILITY_OVERVIEW_TAB_PREFERENCES_KEY[\s\S]*savedFacilityOverviewTabForCurrentAccount[\s\S]*rememberFacilityOverviewTabForCurrentAccount[\s\S]*resetFacilityOverviewSessionState/,
  "Director overview should remember only the active tab separately for each account",
);
assert.match(
  appSource.match(/function resetFacilityOverviewSessionState[\s\S]*?async function openFacilityOverview/)?.[0] || "",
  /date = today[\s\S]*staffExpanded = new Set\(\)[\s\S]*byStreamFrom = today[\s\S]*byStreamTo = today[\s\S]*togetherStaffKeys = \[""\]/,
  "A new account session should reset On shift, ED staff, Working together, and By stream state while retaining the saved tab",
);
assert.match(
  appSource.match(/async function openFacilityOverviewStaffSection[\s\S]*?function focusFacilityOverviewStaffSection/)?.[0] || "",
  /facilityOverviewState\.tab = "staff"[\s\S]*facilityOverviewState\.facilityKey = source[\s\S]*staffExpanded = new Set\(\[sectionKey\]\)[\s\S]*openFacilityOverview\(\{ preserveStaffTerm: true \}\)/,
  "A seniority link should preserve its ED and term while opening the requested ED staff accordion",
);
assert.match(
  appSource.match(/function facilityOverviewDoctorOptionFor[\s\S]*?async function openFacilityOverviewStaffSection/)?.[0] || "",
  /doctor\?\.aliases[\s\S]*normalizedDoctorSourceTypes[\s\S]*switchDoctorSelection\(doctor\.key/,
  "At a glance calendar links should resolve roster aliases through the Creator switcher",
);
assert.match(styleSource, /\.facility-overview-seniority-link[\s\S]*text-decoration: underline;/, "Clickable seniority labels should have a visible link treatment");
assert.match(styleSource, /\.facility-overview-staff-section \{[\s\S]*?overflow: visible;/, "ED staff accordions should not clip open staff action menus");
assert.match(appSource, /facility-overview-staff-term-control[\s\S]*data-facility-overview-staff-term/, "ED staff should expose a synchronized desktop term selector beside Find staff");
assert.match(appSource, /previous_staff[\s\S]*Previous staff[\s\S]*Restore to current staff/, "SMS previous staff should remain visible and be restorable");
assert.match(appSource, /long_service_leave[\s\S]*sabbatical_leave[\s\S]*sick_leave[\s\S]*personal_leave/, "No-shift staff menu should offer the approved leave designations");
assert.match(styleSource, /facility-overview-seniority-label-width[\s\S]*facility-overview-staff-section-count/, "ED staff counts should use a shared invisible alignment column");
assert.match(stateSource, /action === "setFacilityStaffDesignation"[\s\S]*Creator access on the Creator profile is required[\s\S]*action === "clearFacilityStaffDesignation"/, "Staff designation changes must be Creator-only server actions");
assert.match(d1CalendarSource, /facility_staff_designations[\s\S]*reconcileFacilityStaffDesignationsForRosterFile[\s\S]*queryFacilityDesignationLeaveEvents/, "Designations should persist, surface as calendar leave, and reconcile on newer rosters");
assert.match(appSource, /renderFacilityOverviewOnShiftNames[\s\S]*renderFacilityOverviewOnShiftSeniority[\s\S]*facilityOverviewCompactSeniorityLabel/, "On shift cards should show a compact seniority beside each staff name");
assert.match(appSource, /function facilityOverviewEffectiveTeam\(assignment\) \{\s*return String\(assignment\?\.team \|\| ""\)\.trim\(\);\s*\}/, "On shift stream placement should come from the roster, not the contact list");
assert.doesNotMatch(appSource, /Roster only/, "On shift must not move rostered clinicians into a contact-list-only section");
assert.match(appSource, /facility-overview-on-shift-identity[\s\S]*renderFacilityOverviewOnShiftSeniority[\s\S]*facility-overview-on-shift-details[\s\S]*renderFacilityOverviewContactAllocation/, "On shift rows should keep the grade with the clinician name and phone details separate");
assert.match(styleSource, /\.facility-overview-on-shift-person \{[\s\S]*grid-template-columns: minmax\(0, 1fr\) auto;[\s\S]*\.facility-overview-on-shift-identity,[\s\S]*\.facility-overview-on-shift-details[\s\S]*\.facility-overview-on-shift-details \{[\s\S]*justify-content: flex-end;/, "On shift cards should keep grades adjacent to names and phone details right-aligned");
assert.match(styleSource, /\.facility-overview-on-shift-names \{[\s\S]*gap: 8px;[\s\S]*\.facility-overview-on-shift-person \{[\s\S]*align-items: start;[\s\S]*padding-block: 2px;[\s\S]*\.facility-overview-on-shift-identity \{[\s\S]*row-gap: 4px;/, "On shift cards should reserve vertical space for wrapped clinician names");
assert.match(styleSource, /\.facility-overview-staff-card \{[\s\S]*align-content: start;/, "On shift cards should place names directly beneath their heading even when neighbouring cards are taller");
assert.match(appSource, /renderFacilityOverviewContactListStatus[\s\S]*contact\.shift[\s\S]*contact\.reviewReason/, "Creator review should identify the period and reason for any unresolved contact allocation");
assert.match(d1CalendarSource, /export async function queryFacilityOverviewOnShift[\s\S]*\.filter\(\(row\) => row\.doctorKey && row\.displayName && row\.event\);[\s\S]*return events;/, "On shift should return its compact daily roster without a term-wide grade-resolution scan");
assert.match(d1CalendarSource, /applyFacilityStaffSeniorityOverridesToCoworkerEvents[\s\S]*FROM roster_file_doctors[\s\S]*membershipGradesByPerson/, "Effective-grade resolution should use a known active membership grade when a shift is Unknown");
assert.match(
  d1CalendarSource,
  /function sanitizeFileDoctors[\s\S]*providerStaffId:[\s\S]*function bulkInsertFileDoctorStatements[\s\S]*provider_staff_id/,
  "chunked roster saves must preserve FindMyShift staff IDs on DDH membership rows",
);
assert.match(
  d1CalendarSource,
  /function collectDerivedEventAndIssueRows[\s\S]*event\.providerStaffId \|\| doctor\.providerStaffId[\s\S]*function bulkInsertEventStatements[\s\S]*provider_staff_id/,
  "chunked roster saves must preserve FindMyShift staff IDs on DDH event rows",
);
assert.match(d1CalendarSource, /const overrideTerms = new Map\(\)[\s\S]*Promise\.all\(\[\.\.\.overrideTerms\.entries\(\)\]/, "Effective-grade resolution should load overrides once per facility and term, rather than once per staff row");
assert.match(d1CalendarSource, /export async function queryFacilityOverviewStaff[\s\S]*roster_events\.start_date[\s\S]*event: \{ start: String\(row\.start_date/, "ED staff should return lightweight dated grade records rather than full calendar-event JSON");
assert.match(appSource, /physiotherapist[\s\S]*nurse practitioner[\s\S]*Fast Track[\s\S]*facilityOverviewDetectedSeniority/, "Physios and nurse practitioners should be identified and placed in Fast Track");
assert.match(appSource, /Senior Registrar"\) return "SR"[\s\S]*Transitional\/Intermediate Registrar"\) return "TR"[\s\S]*Junior Registrar"\) return "JR"/, "On shift should abbreviate registrar seniorities");
assert.match(appSource, /Edit designation[\s\S]*data-facility-overview-set-staff-seniority[\s\S]*Use roster designation/, "Creator staff menus should provide effective seniority editing and a roster reset");
assert.match(stateSource, /action === "setFacilityStaffSeniorityOverride"[\s\S]*Creator access on the Creator profile is required/, "Seniority corrections must be Creator-only server actions");
assert.match(d1CalendarSource, /facility_staff_seniority_overrides[\s\S]*term_start <= \?[\s\S]*use_roster_seniority/, "Seniority overrides should be effective-dated and support returning to roster data");
assert.match(appSource, /sourceCode === "MMC"[\s\S]*"am shift"[\s\S]*"pm shift"[\s\S]*return false/, "MMC generic AM and PM shift labels should fall back to seniority grouping");
assert.match(appSource, /queueFacilityOverviewMenuPositioning[\s\S]*maxHeight[\s\S]*window\.innerHeight - rect\.height - margin/, "At a glance menus should clamp themselves inside the visible viewport");
assert.match(appSource, /applyFacilityOverviewStaffSeniorityOverrides[\s\S]*staffData\.seniorityOverrides[\s\S]*renderFacilityOverviewStaffBodyPreservingViewport/, "Staff designation changes should update the cached ED staff data in place");
assert.doesNotMatch(appSource, /async function refreshFacilityOverviewAfterStaffChange/, "Staff designation changes should not reload the complete ED staff list");
assert.match(appSource, /renderFacilityOverviewStaffBodyPreservingViewport[\s\S]*scrollTop[\s\S]*requestAnimationFrame/, "In-place ED staff updates should preserve the current viewport");
assert.match(appSource, /grade === "Unknown"[\s\S]*data-facility-overview-staff-multi-select[\s\S]*Multi-select/, "Expanded Unknown staff sections should offer multi-select mode");
assert.match(appSource, /data-facility-overview-staff-multi-select-name[\s\S]*toggleFacilityOverviewStaffMultiSelectMember/, "Multi-select mode should toggle names instead of opening the normal staff action");
assert.match(appSource, /data-facility-overview-staff-designation-menu], \[data-facility-overview-staff-seniority-menu][\s\S]*facilityOverviewStaffMultiSelectSectionForControl[\s\S]*openFacilityOverviewStaffBulkSeniorityMenu/, "Selected staff designation controls should open the bulk context menu");
assert.match(appSource, /mobileDesignationTrigger[\s\S]*isMobileLayout\(\)[\s\S]*data-facility-overview-staff-seniority-menu[\s\S]*renderFacilityOverviewStaffBodyPreservingViewport/, "Mobile taps on staff designation controls should open their edit menu");
assert.match(appSource, /data-facility-overview-staff-bulk-seniority-action-menu[\s\S]*data-facility-overview-set-bulk-staff-seniority[\s\S]*Use roster designation/, "Selected Unknown staff should have a bulk designation context menu");
assert.match(stateSource, /action === "setFacilityStaffSeniorityOverrides"[\s\S]*staff\.length > 100[\s\S]*setFacilityStaffSeniorityOverrides/, "Bulk staff designation changes should be Creator-only and bounded");
assert.match(d1CalendarSource, /export async function setFacilityStaffSeniorityOverrides[\s\S]*db\.batch[\s\S]*loadFacilityStaffSeniorityOverride/, "Bulk designation changes should be persisted together and return their overrides");
assert.match(d1CalendarSource, /facility_sms_memberships[\s\S]*recordFacilitySmsMembershipsForRosterFile[\s\S]*first_seen_date <= \?/, "SMS membership should persist independently of roster-file supersession and carry into later terms");
assert.match(styleSource, /#facilityOverviewBody \{[\s\S]*?height: 100%;[\s\S]*?max-height: 100%;[\s\S]*?overflow-y: auto;/, "Working together content should scroll within the bounded overview body");
assert.match(styleSource, /#facilityOverviewBody\.is-working-together > \.facility-overview-together \{[\s\S]*?height: 100%;[\s\S]*?overflow-y: auto;/, "Working together should use an explicit full-height tab scroller");
assert.match(stateSource, /action === "downloadFindmyshiftExceptions"[\s\S]*findmyshiftDandenongAssignmentExceptions[\s\S]*findmyshiftExceptionCsv/, "FindMyShift exception downloads must be creator-only server-side report reads");
assert.match(findmyshiftCheckSource, /isTransientFindmyshiftRateLimitError[\s\S]*current\?\.lastSuccessAt[\s\S]*returned HTTP 429/, "a transient FindMyShift rate limit should neither mark a successful source failed nor cause it to be downloaded again");
assert.match(findmyshiftModuleSource, /NEXT_TERM_LOOKAHEAD_DAYS = 28[\s\S]*findmyshiftPublicationWindow/, "FindMyShift should use a four-week early-publication window for the next term");
assert.match(findmyshiftCheckSource, /IMPORT_FORMAT = "stream-paired-v7"[\s\S]*term-window change deliberately[\s\S]*rangeState\.requested[\s\S]*importFormat: IMPORT_FORMAT/, "a new FindMyShift parser revision or term window should bypass an unchanged provider version and persist its requested range");
assert.match(findmyshiftCheckSource, /findmyshift-no-shifts[\s\S]*waiting-for-publication/, "an unpublished upcoming FindMyShift term should wait for a provider update instead of surfacing as an import failure");
assert.match(
  findmyshiftCheckSource,
  /current\.lastSuccessAt[\s\S]*reconcileCurrentFindmyshiftRoster[\s\S]*reconcileRosterFileSupersession/,
  "an unchanged FindMyShift check should still reconcile previously imported duplicate roster rows",
);
assert.doesNotMatch(
  stateSource.match(/if \(action === "testFindmyshiftConnection"\)[\s\S]*?if \(action === "adminCreateUser"\)/)?.[0] || "",
  /Promise\.all/,
  "FindMyShift diagnostics must not make concurrent API requests",
);
assert.doesNotMatch(
  findmyshiftModuleSource.match(/export async function findmyshiftRosterWorkbook[\s\S]*?export async function findmyshiftShiftReport/)?.[0] || "",
  /Promise\.all/,
  "FindMyShift import preparation must not make concurrent API requests",
);
assert.doesNotMatch(
  findmyshiftModuleSource.match(/export async function findmyshiftShiftReport[\s\S]*?export async function findmyshiftStaffList/)?.[0] || "",
  /publishedShifts|timesheetData|daysToInclude|groupingInterval|groupByStaff/,
  "FindMyShift report import should use the verified minimal Developer API request",
);
assert.ok(
  automationIngestSource.indexOf("findRosterSyncByProviderVersion") < automationIngestSource.indexOf("file.arrayBuffer()"),
  "automation ingress should reject an unchanged provider version before hashing or storing file bytes",
);
assert.match(
  automationWorkflowSource,
  /workflow_dispatch:[\s\S]*dispatch_id[\s\S]*Record processor start[\s\S]*Record processor completion/,
  "the processor workflow should receive a dispatch id and report lifecycle state",
);
assert.doesNotMatch(automationWorkflowSource, /schedule:/, "GitHub cron must not be the roster processor trigger");
assert.match(automationIngestSource, /requestQueuedRosterProcessing/, "a newly retained roster should request the processor immediately");
assert.match(
  findmyshiftCheckSource,
  /IMPORT_FORMAT = "stream-paired-v7"[\s\S]*Dandenong-FindMyShift-\$\{IMPORT_FORMAT\}[\s\S]*saved\?\.importFormat[\s\S]*importFormat: IMPORT_FORMAT/,
  "a corrected FindMyShift parser should retain a fresh generated source and bypass an older parser revision",
);
assert.match(automationDispatchSource, /GITHUB_ACTIONS_TOKEN[\s\S]*actions\/workflows[\s\S]*\/dispatches/, "dispatches should use a server-side GitHub Actions token");
assert.match(automationDispatchEndpointSource, /ROSTER_WATCHDOG_TOKEN/, "the watchdog should use a dedicated credential");
assert.match(
  appSource.match(/function renderSystemAdminCard[\s\S]*?function renderAdminConsoleMarkup/)?.[0] || "",
  /<details class="advanced-roster-recovery">[\s\S]*Rebuild all retained rosters/,
  "full roster rebuild should be kept behind the advanced recovery disclosure",
);
assert.match(
  appSource.match(/async function replaceActiveRostersWithCurrentUploads[\s\S]*?function retainedRosterEntriesFromStatus/)?.[0] || "",
  /hasPendingRosterAutomation\(\)[\s\S]*window\.confirm[\s\S]*window\.prompt[\s\S]*confirmation/,
  "advanced recovery should be blocked by pending automation and require explicit confirmation",
);
assert.match(
  appSource.match(/function mergeRosterFileEntries[\s\S]*?function mergeSelectedFilesWithRosterStoreStatus/)?.[0] || "",
  /startsWith\("automation:"\)[\s\S]*!storeIds\.has\(entry\.id\)/,
  "obsolete queued or failed automation references should be removed from the creator's browser file list",
);
assert.match(
  appSource.match(/function openLoginModal[\s\S]*?function closeLoginModal/)?.[0] || "",
  /STAY_LOGGED_IN_PREFERENCE_KEY[\s\S]*stayLoggedInPreference === null \? true/,
  "the login screen should default to staying logged in while preserving an explicit opt-out",
);
assert.ok(
  indexSource.indexOf('id="status"') > indexSource.indexOf('id="previewSection"'),
  "live status messages should render below the calendar so mobile updates cannot displace it",
);
assert.doesNotMatch(appSource, /import[^;]+from ["']xlsx["']/, "calendar startup must not eagerly load the spreadsheet engine");
assert.doesNotMatch(rosterSource, /import[^;(]+from ["'](?:xlsx|fflate)["']/, "roster startup must not eagerly load parsing engines");
assert.match(rosterSource, /spreadsheetDependencyPromise = import\("xlsx"\)/, "spreadsheet parsing should lazy-load its engine");
assert.match(rosterSource, /pdfDependencyPromise = import\("fflate"\)/, "PDF parsing should lazy-load its decompression engine");
assert.match(
  appSource,
  /function pasteCopiedEvent[\s\S]*openCustomEventModal\(previewEventToCustomEvent\(shifted\), targetDate, \{ draft: true \}\);/,
  "pasting should open a custom-event draft instead of persisting immediately",
);
assert.doesNotMatch(
  appSource.match(/function openCustomEventModal[\s\S]*?function closeCustomEventModal/)?.[0] || "",
  /renderInlineWhoInsight/,
  "custom-event modal should not request roster coworker insights",
);
assert.match(
  appSource.match(/function openReviewModal[\s\S]*?function closeReviewModal/)?.[0] || "",
  /canUseRosterInsights\(\) && !isLeaveEvent\(event\)[\s\S]*renderInlineWhoInsight/,
  "leave review modals should not request roster coworker insights",
);
assert.match(appSource, /let cloudStateSaveQueue = Promise\.resolve\(\);/, "cloud saves should be serialized");
assert.match(
  appSource.match(/async function enterUserAccount[\s\S]*?async function enterDoctorProfileView/)?.[0] || "",
  /cancelScheduledCloudStateSave\(\)[\s\S]*outgoingSnapshotSavePayload\(previousState\)[\s\S]*queueBackgroundCloudStateSave\(outgoingSave, \{ delayMs: 1500 \}\)/,
  "switching from the creator account should queue the outgoing profile save without blocking entry",
);
assert.match(
  appSource.match(/async function enterUserAccount[\s\S]*?async function enterDoctorProfileView/)?.[0] || "",
  /accountSwitchStartedAt[\s\S]*renderCachedCalendarSnapshotForContextAsync\(targetContext[\s\S]*validateClaimedAccountCalendarInBackground\(targetContext/,
  "switched-account entry should render a cached target snapshot before background validation",
);
assert.match(appSource, /function markAccountSwitchPhase/, "account switching should expose debug timings separately from login timings");
assert.match(appSource, /function renderCachedCalendarSnapshotForContext/, "calendar switching should have an explicit-context snapshot renderer");
assert.match(
  appSource.match(/function renderCachedCalendarSnapshotForContext[\s\S]*?function saveWorkspaceSnapshotForEmail/)?.[0] || "",
  /expectedRevision[\s\S]*cached\.calendarRevision[\s\S]*return false/,
  "browser snapshot rendering should reject stale local cache when the server supplied a current revision",
);
assert.match(appSource, /function validateDoctorProfileCalendarInBackground/, "doctor-profile switching should validate cached snapshots in the background");
assert.match(appSource, /function queueCreatorSwitchTargetPrefetch\(\)/, "creator login should prefetch switch targets into browser snapshot cache");
assert.match(
  appSource.match(/function hasFileDrag[\s\S]*?function abortRosterFileDrag/)?.[0] || "",
  /public\.file-url[\s\S]*application\/x-moz-file[\s\S]*const active = hasFileDrag\(dataTransfer\)/,
  "roster drag overlay should recognise native file drags without requiring file metadata before drop",
);
assert.match(
  appSource.match(/window\.addEventListener\(\"dragleave\"[\s\S]*?document\.addEventListener\(\"keydown\"/)?.[0] || "",
  /is-roster-dragging[\s\S]*relatedTarget[\s\S]*document\.documentElement\.contains\(related\)[\s\S]*clearRosterDragState\(\)[\s\S]*window\.addEventListener\(\"blur\", clearRosterDragState\)/,
  "roster drag overlay should clear promptly when a native drag exits or is cancelled, even if drag metadata disappears",
);
assert.match(
  styleSource,
  /body\.is-roster-dragging #addRosterFilesButton/,
  "roster drag overlay should spotlight the add roster files button",
);
assert.match(
  appSource.match(/function clearRosterDragState[\s\S]*?async function validateFreshRosterUploads/)?.[0] || "",
  /rosterDragDepth = 0;[\s\S]*clearRosterDragVisualState\(\)/,
  "roster drag overlay should clear only through an explicit drag-end, leave, drop, or cancellation path",
);
assert.doesNotMatch(
  appSource,
  /rosterDragStaleTimer|touchRosterDragActivity/,
  "roster drag overlay should not expire while a file remains held over the page",
);
assert.match(
  appSource.match(/function validateIncomingFiles[\s\S]*?async function analyzeFiles/)?.[0] || "",
  /showRosterImportError[\s\S]*Please drop an Excel or PDF roster file[\s\S]*window\.setTimeout[\s\S]*5000/,
  "invalid dropped files should show a temporary dismissible valid-roster-file modal instead of a console error",
);
assert.match(
  appSource.match(/document\.addEventListener\(\"keydown\"[\s\S]*?window\.addEventListener\(\"drop\"/)?.[0] || "",
  /Escape[\s\S]*abortRosterFileDrag/,
  "escape should abort an active roster file drag",
);
assert.match(
  appSource.match(/function handleRosterDragOver[\s\S]*?function syncRosterDragState/)?.[0] || "",
  /rosterDragAborted[\s\S]*dropEffect = \"none\"/,
  "aborted roster drags should reject drops until the drag session ends",
);
assert.match(
  d1CalendarSource.match(/export async function queryCalendarRevision[\s\S]*?export async function upsertAccountMirror/)?.[0] || "",
  /FROM parser_rules[\s\S]*local_parser_extensions_json/,
  "calendar revisions should include parser-rule state so resolved warnings invalidate stale snapshots",
);
assert.match(
  stateSource,
  /filterCachedSnapshotForReturn[\s\S]*filterSnapshotPreviewIssuesForOwner[\s\S]*filterStoredRosterIssuesForPreview\(issues, ruleSets/,
  "cached server snapshots should be filtered through current parser rules before return",
);
assert.match(
  appSource.match(/async function hydrateAuthenticatedWorkspace[\s\S]*?function markLoginPhase/)?.[0] || "",
  /currentUserEmail === OWNER_EMAIL[\s\S]*forceCreatorDoctorSession\(\)[\s\S]*loadCloudCalendarEvents/,
  "creator hydration should normalize the creator doctor before calendar events load",
);
assert.match(
  stateSource.match(/if \(action === "adminLoadUser"\)[\s\S]*?if \(action === "claimRosterName"\)/)?.[0] || "",
  /prepareAccountResponse[\s\S]*loadAccountSnapshotPayload[\s\S]*snapshotStatus[\s\S]*snapshotSource[\s\S]*snapshotRevision[\s\S]*viewedAccountPayload/,
  "admin account switching should hydrate the target account and return an inline snapshot payload",
);
assert.match(
  appSource.match(/async function returnToCreatorAccount[\s\S]*?async function clearLocalWorkspace/)?.[0] || "",
  /forceCreatorDoctorSession\(\);[\s\S]*renderCachedCalendarSnapshotForContextAsync\(targetContext[\s\S]*validateClaimedAccountCalendarInBackground/,
  "returning to the creator should render the cached creator calendar before calendar validation",
);
assert.match(
  appSource.match(/function currentAccount[\s\S]*?function canUseRosterInsights/)?.[0] || "",
  /activeCalendarMode\(\) === "doctor-profile" && activeDoctorProfile[\s\S]*function viewedAccountEmail[\s\S]*activeCalendarMode\(\) === "doctor-profile"[\s\S]*function isOwnerAccount\(\) \{\s*return isViewingCreatorAccount\(\);/,
  "switched-user account surfaces should be driven by viewed identity rather than creator authentication",
);
assert.match(
  appSource.match(/function canUseCreatorDoctorSwitcher[\s\S]*?function canReturnToCreator/)?.[0] || "",
  /canUseDoctorPicker\(\)[\s\S]*activeCalendarMode\(\) === "doctor-profile"[\s\S]*activeCalendarMode\(\) === "claimed-account" && isImpersonating/,
  "the global doctor switcher should remain available while the authenticated Creator views claimed or unclaimed calendars",
);
assert.doesNotMatch(
  appSource.match(/function canUseCreatorDoctorSwitcher[\s\S]*?function canReturnToCreator/)?.[0] || "",
  /adminViewingEmail/,
  "Creator switcher access should use explicit impersonation state rather than the viewed email alone",
);
assert.doesNotMatch(
  appSource,
  /currentSnapshot\.session/,
  "asynchronous account transitions should never dereference a snapshot that another transition may have cleared",
);
assert.match(
  appSource.match(/function calendarFilesForActiveView[\s\S]*?function rosterDisplayFiles/)?.[0] || "",
  /activeCalendarMode\(\) === "doctor-profile"[\s\S]*activeDoctorProfile\.sourceTypes/,
  "doctor-profile account files should come from the viewed calendar snapshot rather than creator roster status",
);
assert.match(
  stateSource.match(/async function repositoryImportRefsForDoctorProfile[\s\S]*?async function loadDoctorProfileSnapshotInfo/)?.[0] || "",
  /doctorDiagnostics[\s\S]*doctorKeysForOption[\s\S]*queryRosterFileRefsForDoctors/,
  "doctor profile file refs should resolve from every matched roster alias key",
);
assert.match(
  appSource.match(/async function deleteAccount[\s\S]*?function deleteLocalAccountData/)?.[0] || "",
  /creatorDeletingSwitchedUser[\s\S]*if \(creatorDeletingSwitchedUser\)[\s\S]*returnToCreatorCalendar[\s\S]*if \(deletingViewedAccount\)[\s\S]*setActiveCalendarContext\(\"claimed-account\", \{ email: \"\" \}\)/,
  "deleting a switched user should return to creator, while user self-delete should clear into logged-out context",
);
assert.doesNotMatch(
  appSource.match(/function deleteLocalAccountData[\s\S]*?function clearDeletedAccountClaims/)?.[0] || "",
  /accountState\.currentEmail = OWNER_EMAIL/,
  "deleting the current local user should not silently select the creator account",
);
assert.match(
  appSource.match(/function snapshotCloudSavePayload[\s\S]*?function forceCreatorDoctorSession/)?.[0] || "",
  /accountEmail: viewedAccountEmail\(\)[\s\S]*requestEmail: adminViewingEmail \? authenticatedAccountEmail\(\) : viewedAccountEmail\(\)[\s\S]*targetEmail: adminViewingEmail \? viewedAccountEmail\(\) : \"\"/,
  "switched-user settings should save to the viewed account using creator authentication only for the request",
);
assert.match(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/if \(action === "login"\)[\s\S]*?const account = await verifyD1Account/)?.[0] || "",
  /responseMode === "fast"[\s\S]*prepareFastLoginEnvelope\(loginRecord,[\s\S]*loadFastAccountSnapshotPayload[\s\S]*snapshotStatus[\s\S]*snapshotSource[\s\S]*snapshotRevision[\s\S]*viewedAccountPayload/,
  "fast login should build a minimal account payload and use the lightweight snapshot path",
);
assert.match(
  stateSource.match(/if \(action === "login"\)[\s\S]*?const account = await verifyD1Account/)?.[0] || "",
  /loadFastAccountSnapshotPayload\(context, \{[\s\S]*cachedRevision: body\?\.cachedRevision/,
  "fast login should pass the browser revision through to avoid returning an unchanged snapshot",
);
const fastSnapshotSource = stateSource.match(/async function loadFastAccountSnapshotPayload[\s\S]*?function scheduleFastAccountSnapshotValidation/)?.[0] || "";
assert.doesNotMatch(fastSnapshotSource, /queryCalendarRevision/, "fast login should not block on live calendar revision validation");
assert.match(
  fastSnapshotSource,
  /loadSnapshotRegistryEntry[\s\S]*cachedRevision === registryRevision[\s\S]*loadCachedSnapshot[\s\S]*revisionSkipped: true/,
  "fast login should return the ready browser or R2 snapshot using its built revision",
);
assert.doesNotMatch(
  stateSource.match(/function scheduleFastAccountSnapshotValidation[\s\S]*?function scheduleAccountSnapshotRebuild/)?.[0] || "",
  /context\.waitUntil|queryCalendarRevision|buildAndStoreAccountSnapshot/,
  "fast login should not attach an expensive snapshot rebuild to authentication",
);
assert.match(
  stateSource.match(/async function loadSnapshotPayloadFromRegistry[\s\S]*?async function loadFastAccountSnapshotPayload/)?.[0] || "",
  /allowInlineBuild === false \|\| buildInProgress[\s\S]*scheduleRebuild[\s\S]*snapshot: null[\s\S]*snapshotSource: buildInProgress \? "server-cache-building" : cacheBucket\?\.get \? "server-cache-miss" : "d1-inline-disabled"/,
  "non-inline snapshot requests should schedule rebuilds instead of building on cache misses",
);
assert.match(
  stateSource.match(/async function loadSnapshotPayloadFromRegistry[\s\S]*?async function loadFastAccountSnapshotPayload/)?.[0] || "",
  /snapshotRegistryBuildInProgress[\s\S]*server-cache-building/,
  "recent building snapshot rows should suppress duplicate inline rebuilds",
);
assert.match(
  stateSource.match(/async function loadSnapshotPayloadFromRegistry[\s\S]*?async function filterCachedSnapshotForReturn/)?.[0] || "",
  /registry\?\.status === "ready"[\s\S]*registry\?\.builtRevision === calendarRevision[\s\S]*cachedRevision === calendarRevision && registryCurrent/,
  "browser revisions should only be accepted after the matching server snapshot is ready",
);
assert.doesNotMatch(
  stateSource.match(/async function prepareFastLoginEnvelope[\s\S]*?async function applySqlHospitalLocationSettings/)?.[0] || "",
  /repositoryDoctorCandidates/,
  "fast login should not build the expensive creator doctor list on the authentication request",
);
assert.match(
  stateSource.match(/function snapshotWarmupSourceTypeSet[\s\S]*?function scheduleSnapshotWarmupForAllAccounts/)?.[0] || "",
  /accountWarmupAffectedBySourceTypes[\s\S]*doctorProfileWarmupAffectedBySourceTypes[\s\S]*SNAPSHOT_GLOBAL_WARMUP_LIMIT[\s\S]*listSnapshotRegistryWarmupCandidates[\s\S]*statuses: \["ready"\]/,
  "scoped snapshot warmup should only rebuild creator, affected claimed accounts, and matching doctor profiles",
);
assert.match(
  appSource.match(/function doctorProfileSourceTypes[\s\S]*?function doctorOptionsForCurrentAccount/)?.[0] || "",
  /normalizedDoctorSourceTypes\(repositoryDoctor\)/,
  "doctor profile source types should come from doctor metadata rather than defaulting to every hospital",
);
assert.doesNotMatch(
  appSource.match(/function doctorProfileSourceTypes[\s\S]*?function doctorOptionsForCurrentAccount/)?.[0] || "",
  /\["casey", "ddh", "mch", "mmc"\]/,
  "doctor profile source types must not fall back to all hospitals",
);
assert.match(
  appSource.match(/function calendarSnapshotCacheAffectedBySourceTypes[\s\S]*?function invalidateCalendarSnapshotCachesForSourceTypes/)?.[0] || "",
  /profileSources\.some\(\(sourceType\) => changed\.has\(sourceType\)\)/,
  "browser snapshot cache invalidation should be scoped to overlapping hospital source types",
);
assert.match(
  appSource.match(/async function validateDoctorProfileCalendarInBackground[\s\S]*?async function enterUserAccount/)?.[0] || "",
  /browser profile cache can be complete enough to render immediately[\s\S]*cachedRevision: ""[\s\S]*allowInlineBuild: true[\s\S]*waitForDoctorProfileCalendarBuild/,
  "explicit Creator profile switching should obtain the profile's authoritative server snapshot without login-path polling",
);
assert.match(
  appSource.match(/async function waitForDoctorProfileCalendarBuild[\s\S]*?async function enterUserAccount/)?.[0] || "",
  /while \(calendarTransitionStillCurrent\(options\.transition\)\)[\s\S]*retryDelays\[Math\.min\(attempt, retryDelays\.length - 1\)\]/,
  "doctor profile switching should keep polling a scheduled snapshot until the transition changes",
);
assert.doesNotMatch(
  appSource.match(/async function loadUnclaimedDoctorCalendar[\s\S]*?function hasDoctorProfileImportCandidates/)?.[0] || "",
  /calendar is not ready yet\. Try again in a moment/,
  "a cache miss while opening a doctor profile should remain a loading state rather than rejecting the switch",
);
assert.match(
  await readFile(new URL("../public/index.html", import.meta.url), "utf8"),
  /id="switchOverlayCancelButton"[\s\S]*Cancel and return to Creator/,
  "calendar switch overlay should let the Creator cancel a slow target load",
);
assert.match(
  appSource.match(/function showSwitchOverlay[\s\S]*?function showRosterImportOverlay/)?.[0] || "",
  /activeSwitchOverlayCancel[\s\S]*switchOverlayRunId/,
  "switch overlay ownership should stop stale switch completions from hiding a newer overlay",
);
assert.match(
  appSource.match(/async function cancelCreatorCalendarSwitch[\s\S]*?function showRosterImportOverlay/)?.[0] || "",
  /returnToCreatorCalendar\(\{ skipOutgoingSave: true, restoreOnFailure: false \}\)/,
  "cancelling a target switch should return without saving the incomplete target workspace",
);
assert.match(
  appSource.match(/function groupWhoAssignments[\s\S]*?function groupWhoTeams/)?.[0] || "",
  /coalesceWhoAssignments[\s\S]*normalizeWhoRole\(existing\.role\)[\s\S]*normalizeWhoRole\(assignment\.role\)/,
  "who-is-working groups should merge duplicate same-person shifts and retain the row with seniority",
);
assert.match(
  appSource.match(/function loginSnapshotReadyForRender[\s\S]*?function visibleSnapshotIsCurrent/)?.[0] || "",
  /calendarSnapshotMatchesActiveContext/,
  "visible snapshot checks should require the rendered calendar to match the active profile",
);
assert.match(
  appSource.match(/async function validateClaimedAccountCalendarInBackground[\s\S]*?async function validateDoctorProfileCalendarInBackground/)?.[0] || "",
  /preserveRenderedSnapshot && visibleSnapshotIsCurrent\(\{ requireNotStale: true \}\)/,
  "claimed account background validation should skip network work when the browser cache is current",
);
assert.match(
  appSource.match(/async function prefetchCreatorSwitchTarget[\s\S]*?function queueCreatorSwitchTargetPrefetch/)?.[0] || "",
  /allowInlineBuild: false[\s\S]*skipRebuild: true/,
  "creator switch-target prefetch should not build missing snapshots in the background",
);
assert.match(
  stateSource.match(/async function loadFastAccountSnapshotPayload[\s\S]*?function scheduleAccountSnapshotRebuild/)?.[0] || "",
  /scheduleFastAccountSnapshotValidation[\s\S]*revisionSkipped: true[\s\S]*function scheduleFastAccountSnapshotValidation[\s\S]*return false/,
  "fast login snapshots should defer revision validation without attaching snapshot rebuilding to authentication",
);
assert.match(
  stateSource.match(/if \(action === "adminLoadUser"\)[\s\S]*?if \(action === "claimRosterName"\)/)?.[0] || "",
  /if \(action === "loadAccountContext"\)[\s\S]*prepareAccountResponse[\s\S]*issueConfig/,
  "deferred login bootstrap should have a dedicated authenticated account-context route",
);
assert.doesNotMatch(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/async function prepareLightweightAccountResponse[\s\S]*?export async function prepareAccountResponse/)?.[0] || "",
  /queryRosterFiles\(options\.db\)/,
  "lightweight login responses should not load full roster-file metadata",
);
assert.match(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/async function prepareLightweightAccountResponse[\s\S]*?export async function prepareAccountResponse/)?.[0] || "",
  /includeImportRefs === false[\s\S]*imports: \[\]/,
  "lightweight login responses should be able to skip import refs entirely",
);
assert.match(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/if \(action === "save"\)[\s\S]*?if \(action === "loadDoctorProfile"\)/)?.[0] || "",
  /removedImportIds\.length[\s\S]*!repositoryAlreadySynced[\s\S]*syncRosterRepositoryToKeepFileIds/,
  "ordinary saves must only delete roster database rows when explicit removed import ids are supplied and repository is not already synced",
);
assert.match(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8")),
  /deleteRetainedRosterSource\(context\.env\.ROSTER_DB, context\.env\.ROSTER_FILES/,
  "roster deletion paths must remove retained R2 source bytes alongside D1 metadata",
);
assert.doesNotMatch(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/if \(action === "save"\)[\s\S]*?if \(action === "loadDoctorProfile"\)/)?.[0] || "",
  /staleFileIds|canReconcileToFullSet|queryRosterFileRanges/,
  "ordinary saves must not infer roster deletions from partial account import snapshots",
);
assert.doesNotMatch(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/if \(action === "adminCreateUser"\)[\s\S]*?if \(action === "resolveAccountClaims"\)/)?.[0] || "",
  /autoClaimMatchedRosterNames|prepareAccountResponse|loadRepositoryIndex|buildIssueConfig/,
  "admin user creation should stay lightweight and avoid broad account hydration",
);
assert.doesNotMatch(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/if \(action === "loadCalendarEvents"\)[\s\S]*?if \(action === "loadInsightImports"\)/)?.[0] || "",
  /loadRepositoryIndex/,
  "calendar event loads should be D1-first and avoid repository-index hydration",
);
assert.match(
  stateSource.match(/if \(action === "loadDoctorProfile"\)[\s\S]*?if \(action === "saveDoctorProfile"\)/)?.[0] || "",
  /loadDoctorProfileSnapshotPayload[\s\S]*cachedRevision[\s\S]*snapshotStatus[\s\S]*snapshotSource[\s\S]*calendarRevision/,
  "doctor profile loads should delegate cached-revision validation to the profile snapshot registry",
);
assert.match(
  stateSource,
  /async function queryDoctorProfileCalendarRevision[\s\S]*queryCalendarRevision/,
  "doctor profile cache validation should use a lightweight calendar revision",
);
assert.doesNotMatch(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/if \(action === "save"\)[\s\S]*?if \(action === "loadDoctorProfile"\)/)?.[0] || "",
  /loadRepositoryIndex/,
  "ordinary cloud saves should not load the repository index",
);
assert.match(
  stateSource.match(/if \(action === "claimRosterName"\)[\s\S]*?if \(action === "listUsers"\)/)?.[0] || "",
  /loadSqlDoctorCandidates[\s\S]*findDoctorClaimCandidate/,
  "manual roster claims should resolve from SQL doctor candidates",
);
assert.match(
  stateSource.match(/if \(action === "save"\)[\s\S]*?if \(action === "loadDoctorProfile"\)/)?.[0] || "",
  /replaceAccountCustomEvents[\s\S]*stripRelationalCustomEventsFromSession/,
  "account saves should keep custom-event truth in D1 rows rather than session JSON",
);
assert.match(
  stateSource.match(/if \(action === "save"\)[\s\S]*?if \(action === "loadDoctorProfile"\)/)?.[0] || "",
  /repositoryAlreadySynced[\s\S]*syncRosterRepositoryToKeepFileIds/,
  "roster deletion saves should purge derived rows and R2 sources before updating account state unless repositorySynced is set",
);
assert.match(
  stateSource.match(/async function syncRosterRepositoryToKeepFileIds[\s\S]*?async function purgeRosterImports/)?.[0] || "",
  /verifyRosterFilesPurged[\s\S]*repositoryDoctorCandidates[\s\S]*availableDoctors/,
  "repository sync should return live availableDoctors after verified purge",
);
assert.match(
  stateSource.match(/async function syncRosterRepositoryToKeepFileIds[\s\S]*?async function purgeRosterImports/)?.[0] || "",
  /querySourceTypesForFileIds[\s\S]*deleteDerivedRosterFile[\s\S]*deleteRetainedRosterSource[\s\S]*refreshCanonicalDoctors[\s\S]*verifyRosterFilesPurged/,
  "repository sync should purge orphans, refresh canonical doctors, and verify removal",
);
assert.match(
  d1CalendarSource,
  /export async function verifyRosterFilesPurged[\s\S]*purged: !d1File && !r2Raw && eventCount <= 0/,
  "roster purge verification should confirm D1 and R2 rows are absent",
);
assert.match(
  stateSource.match(/if \(action === "syncRosterRepository"\)[\s\S]*?if \(action === "removeRosterImports"\)/)?.[0] || "",
  /syncRosterRepositoryToKeepFileIds[\s\S]*sanitizeRepositoryFileIds\(body\?\.keepFileIds\)/,
  "syncRosterRepository should accept keepFileIds including empty arrays",
);
assert.match(
  stateSource.match(/if \(action === "removeRosterImports"\)[\s\S]*?if \(action === "appendConsoleMessage"\)/)?.[0] || "",
  /syncRosterRepositoryToKeepFileIds/,
  "dedicated roster removal should sync repository to remaining files without a full account save",
);
assert.match(
  stateSource.match(/async function repositoryDoctorCandidates[\s\S]*?async function refreshCanonicalDoctors/)?.[0] || "",
  /queryRosterFileDoctors\(db\)[\s\S]*buildCanonicalDoctorOptionsFromRows[\s\S]*queryCanonicalDoctors/,
  "doctor directory responses should prefer live roster file doctors over stale canonical cache rows",
);
assert.match(
  stateSource.match(/if \(action === "listUsers"\)[\s\S]*?if \(action === "calendarStoreStatus"\)/)?.[0] || "",
  /repositoryDoctorCandidates[\s\S]*preferCanonical: true/,
  "routine creator doctor-directory loading should use the refreshed canonical cache instead of rebuilding every identity",
);
assert.match(
  stateSource.match(/if \(action === "loadAccountContext"\)[\s\S]*?if \(action === "claimRosterName"\)/)?.[0] || "",
  /includeAvailableDoctors:[\s\S]*=== "creator"[\s\S]*=== "owner"/,
  "deferred creator account context should hydrate the complete doctor switcher",
);
assert.match(
  stateSource.match(/if \(action === "resolveAccountClaims"\)[\s\S]*?if \(action === "adminLoadUser"\)/)?.[0] || "",
  /resolvedClaims[\s\S]*includeAvailableDoctors:[\s\S]*&& !resolvedClaims\.length/,
  "claim resolution should only load full doctor candidates while an account still has no claims",
);
assert.doesNotMatch(
  stateSource.match(/if \(action === "listUsers"\)[\s\S]*?if \(action === "calendarStoreStatus"\)/)?.[0] || "",
  /cleanupResolvedAdminIssues/,
  "creator user-list loading should not synchronously clean every account",
);
assert.match(
  stateSource.match(/if \(action === "listUsers"\)[\s\S]*?if \(action === "calendarStoreStatus"\)/)?.[0] || "",
  /globalParserExtensions[\s\S]*repairAccountClaimsIfNeeded[\s\S]*userSummaryFromRecord\(record\.email, record, \{ db: context\.env\.ROSTER_DB, globalParserExtensions \}\)/,
  "creator user-list loading should reuse one parser-rule load for stale issue filtering",
);
assert.match(
  stateSource.match(/async function userSummaryFromRecord[\s\S]*?function insightsEnabledForRecord/)?.[0] || "",
  /filterResolvedAdminIssuesForSummary/,
  "user summaries should hide resolved stale parser issues without foreground cleanup",
);
assert.match(
  appSource.match(/async function hydrateAuthenticatedWorkspace[\s\S]*?function markLoginPhase/)?.[0] || "",
  /adminTargetEmail && adminTargetEmail !== OWNER_EMAIL && !currentRosterClaims\.length[\s\S]*resolveCurrentAccountClaims\(adminTargetEmail\)[\s\S]*!adminTargetEmail && currentUserEmail !== OWNER_EMAIL && !currentRosterClaims\.length[\s\S]*resolveCurrentAccountClaims\(\)[\s\S]*loadCloudCalendarEvents[\s\S]*queueDeferredBootstrapImports[\s\S]*void loadServerUsers\(\)/,
  "claim resolution should be skipped for already-claimed account entry while still resolving unclaimed users before calendar load",
);
assert.match(
  appSource.match(/async function validateClaimedAccountCalendarInBackground[\s\S]*?async function validateDoctorProfileCalendarInBackground/)?.[0] || "",
  /restoreCloudState[\s\S]*!cachedSnapshot\?\.preview && currentSnapshot\?\.preview[\s\S]*renderWorkspaceFromSnapshot\(currentSnapshot[\s\S]*hydrateAuthenticatedWorkspace/,
  "claimed-account switching should render an inline server snapshot before background hydration finishes",
);
assert.match(
  stateSource.match(/async function buildDerivedAccountSnapshot[\s\S]*?function rawRosterObjectKey/)?.[0] || "",
  /snapshotFileRefs[\s\S]*d1RepositoryImportRefsForClaims\(db, claims\)[\s\S]*fileRefs: snapshotFileRefs/,
  "claimed-account snapshots should derive source file refs from D1 claims when lightweight account state omits imports",
);
assert.match(
  appSource.match(/async function claimSelectedRosterName[\s\S]*?async function updatePreview/)?.[0] || "",
  /applyCloudStateData\(data\)[\s\S]*loadCloudCalendarEvents\(\{ adminTargetEmail: adminViewingEmail \? viewedAccountEmail\(\) : "" \}\)[\s\S]*if \(loadedCalendar\)[\s\S]*bootstrapImports\(\)/,
  "manual roster claims should load a D1 calendar snapshot before bootstrapping repository refs",
);
assert.match(
  appSource.match(/async function bootstrapImports[\s\S]*?function snapshotHasUnresolvablePreviewEvents/)?.[0] || "",
  /cloudAvailable && selectedFiles\.length && selectedFiles\.some\(\(entry\) => !entry\.file\)[\s\S]*loadCloudCalendarEvents\(\{[\s\S]*adminTargetEmail: adminViewingEmail \? viewedAccountEmail\(\) : ""[\s\S]*transition: options\.transition[\s\S]*No D1 calendar events were found/,
  "cloud repository refs should use D1 calendar loading instead of falling through to local browser parsing",
);
assert.doesNotMatch(
  appSource.match(/async function applyCloudStateData[\s\S]*?async function loadCloudCalendarEvents/)?.[0] || "",
  /shouldRebuildAccountSnapshot\(currentSnapshot\)/,
  "server-delivered account snapshots should not be discarded before first render",
);
assert.match(
  appSource.match(/async function bootstrapImports[\s\S]*?function snapshotHasUnresolvablePreviewEvents/)?.[0] || "",
  /currentSnapshot\?\.preview[\s\S]*renderWorkspaceFromSnapshot\(currentSnapshot[\s\S]*loadCloudCalendarEvents\(\{[\s\S]*adminTargetEmail: adminViewingEmail \? viewedAccountEmail\(\) : ""[\s\S]*preserveExistingSnapshot: true/,
  "bootstrap should keep rendering server snapshots while cloud refresh runs in the background",
);
assert.match(
  appSource.match(/function renderFilesMarkup[\s\S]*?function renderFilesList/)?.[0] || "",
  /rosterDisplayFiles\(hasUsableStatus, statusOnlyEntries\)[\s\S]*pendingRemovedImportIds/,
  "creator roster lists should merge D1 status files and hide files pending removal",
);
assert.match(
  appSource.match(/async function removeStoredImport[\s\S]*?function loadConflictSelections/)?.[0] || "",
  /syncRosterRepositoryToSelection[\s\S]*completeRosterRemovalAfterSync[\s\S]*scheduleRosterRemovalRetry/,
  "roster deletion should sync repository to remaining selection and finish only after verified purge",
);
assert.match(
  appSource.match(/function accountImportsSavePayload[\s\S]*?function restoreRemovedImportAfterFailedRemoval/)?.[0] || "",
  /removedImportIds: \[\][\s\S]*repositorySynced: true/,
  "post-delete account saves should skip re-purge when repository is already synced",
);
assert.match(
  appSource.match(/async function completeRosterRemovalAfterSync[\s\S]*?function scheduleRosterRemovalRetry/)?.[0] || "",
  /waitForCreatorSwitcherRemovalSettled[\s\S]*tryAnnounceCreatorSwitcherRosterUpdate[\s\S]*queueAccountImportsSave/,
  "roster removal completion should wait for switcher settlement before announcing and account sync",
);
assert.match(
  appSource.match(/function syncRosterRepositoryToSelection[\s\S]*?function restoreRemovedImportAfterFailedRemoval/)?.[0] || "",
  /applyAuthoritativeAvailableDoctors\(syncResult\.availableDoctors\)/,
  "repository sync should replace available doctors from the server response",
);
assert.match(
  appSource.match(/function pickerHasRemovedSourceDoctors[\s\S]*?function restoreRemovedImportAfterFailedRemoval/)?.[0] || "",
  /isCreatorSwitcherRemovalStable[\s\S]*waitForCreatorSwitcherRemovalSettled/,
  "delete switcher settlement should wait until removed source doctors disappear from the picker",
);
assert.match(
  appSource.match(/function queueAccountImportsSave[\s\S]*?function keepFileIdsAfterRemoval/)?.[0] || "",
  /reportErrors: false/,
  "background account sync after roster removal should not surface cloud save failures as delete errors",
);
assert.match(
  appSource.match(/function scheduleRosterRemovalRetry[\s\S]*?async function removeStoredImport/)?.[0] || "",
  /Retrying removal of \$\{removedName\}/,
  "background roster removal retries should report progress in the admin console",
);
assert.match(
  appSource.match(/function rosterDisplayFiles[\s\S]*?function selectedFilesNeedD1CalendarReload/)?.[0] || "",
  /pendingRemovedImportIds\.has\(entry\.id\)/,
  "creator roster lists should hide files pending background removal even when D1 status is stale",
);
assert.match(
  appSource.match(/async function mergeFiles[\s\S]*?function validateIncomingFiles/)?.[0] || "",
  /needsD1Resync: true/,
  "explicit roster uploads should mark files for a forced D1 resync",
);
assert.match(
  appSource.match(/async function refreshCreatorCalendarAfterFileChange[\s\S]*?async function refreshAvailableDoctorsAfterRosterChange/)?.[0] || "",
  /needsD1Resync === true[\s\S]*saveSelectedRosterFilesToD1[\s\S]*force:/,
  "creator roster changes should force D1 resync for freshly uploaded local files",
);
assert.match(
  appSource.match(/function mergeRosterFileEntries[\s\S]*?function mergeSelectedFilesWithRosterStoreStatus/)?.[0] || "",
  /existing\.sourceType === "pending"[\s\S]*storeEntry\.sourceType \|\| existing\.sourceType/,
  "merged roster file entries should prefer D1 source types over pending placeholders",
);
assert.match(
  appSource.match(/async function importRosterFiles[\s\S]*?async function switchDoctorSelection/)?.[0] || "",
  /showRosterImportOverlay[\s\S]*await refreshCreatorCalendarAfterFileChange/,
  "roster imports should keep progress visible until the initial calendar refresh completes",
);
assert.match(
  appSource.match(/async function finishRosterImportOverlay[\s\S]*?async function importRosterFiles/)?.[0] || "",
  /ROSTER_IMPORT_OVERLAY_MAX_MS/,
  "roster import overlay should hide after the configured maximum duration",
);
assert.match(
  appSource.match(/async function refreshCalendarStoreStatus[\s\S]*?async function toggleAdminConsole/)?.[0] || "",
  /includeAvailableDoctors === true && Array\.isArray\(data\.availableDoctors\)/,
  "roster status refreshes should apply doctor-directory updates even when the list shrinks",
);
assert.match(
  appSource.match(/async function refreshAvailableDoctorsAfterRosterChange[\s\S]*?function normalizeSavedExportRange/)?.[0] || "",
  /mergeAvailableDoctors = options\.mergeAvailableDoctors === true[\s\S]*10000, 20000[\s\S]*mergeAvailableDoctors[\s\S]*syncCreatorDoctorPickerWithRemainingRosters\(\{ localOnly \}\)[\s\S]*renderDoctorState/,
  "roster changes should poll repository doctors and optionally replace the local picker list on delete",
);
assert.match(
  appSource.match(/async function completeRosterRemovalAfterSync[\s\S]*?function scheduleRosterRemovalRetry/)?.[0] || "",
  /refreshAvailableDoctorsAfterRosterChange\(\{ localOnly: true, mergeAvailableDoctors: false \}\)/,
  "roster deletion completion should poll server doctors without merging stale removed-source names back in",
);
assert.doesNotMatch(
  appSource.match(/async function removeStoredImport[\s\S]*?setStatus\(`Removing \$\{removedName\}\.\.\.`\)/)?.[0] || "",
  /syncCreatorDoctorPickerWithRemainingRosters\(\{ localOnly: true \}\)/,
  "roster deletion should not optimistically mutate the switcher before server purge completes",
);
assert.match(
  appSource.match(/function mergeAvailableRosterDoctors[\s\S]*?async function syncCreatorDoctorPickerWithRemainingRosters/)?.[0] || "",
  /options\.localOnly === true[\s\S]*if \(localOnly\) return sanitizeAvailableRosterDoctors\(merged\)/,
  "local-only doctor merges should not re-add repository doctors removed with a deleted roster",
);
assert.match(
  appSource.match(/async function refreshCreatorCalendarAfterFileChange[\s\S]*?async function rosterDoctorsFromSelectedFiles/)?.[0] || "",
  /afterRosterRemoval[\s\S]*preserveExistingSnapshot: !afterRosterRemoval[\s\S]*localOnly: afterRosterRemoval/,
  "creator roster deletion should reload calendar state and prune switcher doctors",
);
assert.match(
  appSource.match(/function tryAnnounceCreatorSwitcherRosterUpdate[\s\S]*?function renderDoctorState/)?.[0] || "",
  /isCreatorSwitcherRepositorySettled[\s\S]*isCreatorSwitcherVisibleAndAligned[\s\S]*Switcher menu updated/,
  "creator switcher announcements should require repository settlement and visible DOM alignment",
);
assert.match(
  appSource.match(/function visibleCreatorSwitcherSignature[\s\S]*?function captureCreatorSwitcherVisibleBaseline/)?.[0] || "",
  /mobileDoctorSelect[\s\S]*return null/,
  "creator switcher visibility checks should use DOM-only signatures without in-memory fallback",
);
assert.match(
  appSource.match(/async function refreshCreatorCalendarAfterFileChange[\s\S]*?async function rosterDoctorsFromSelectedFiles/)?.[0] || "",
  /captureCreatorSwitcherVisibleBaseline[\s\S]*tryAnnounceCreatorSwitcherRosterUpdate/,
  "roster file changes should capture a visible baseline first and announce only after the refresh completes",
);
assert.doesNotMatch(
  appSource.match(/function renderDoctorState[\s\S]*?function renderClaimSection/)?.[0] || "",
  /announceCreatorSwitcherUpdateIfChanged|tryAnnounceCreatorSwitcherRosterUpdate/,
  "renderDoctorState should rebuild the switcher without emitting roster console messages directly",
);
assert.match(
  appSource.match(/async function removeStoredImport[\s\S]*?cancelScheduledCloudStateSave/)?.[0] || "",
  /creatorSwitcherAnnouncementBaseline = null[\s\S]*captureCreatorSwitcherVisibleBaseline/,
  "roster deletion should reset and capture the visible switcher baseline before optimistic file-list refresh",
);
assert.match(
  appSource.match(/function applyLoadedCalendarFileRefs[\s\S]*?function rosterStoreFileToClientEntry/)?.[0] || "",
  /mergeRosterFileEntries\(fromSnapshot, calendarStoreStatus\)/,
  "creator calendar loads should keep roster files that exist in the store but not yet in the snapshot",
);
assert.match(
  appSource.match(/async function syncCreatorDoctorPickerWithRemainingRosters[\s\S]*?async function pollCalendarAfterRosterChange/)?.[0] || "",
  /ensureSelectedFilesLoaded\(\)/,
  "creator doctor picker sync should hydrate retained roster files before parsing",
);
assert.match(
  appSource.match(/async function refreshCalendarStoreStatus[\s\S]*?async function toggleAdminConsole/)?.[0] || "",
  /mergeAvailableDoctors === true[\s\S]*mergeAvailableRosterDoctors\(availableRosterDoctors, incomingDoctors\)/,
  "roster status refreshes should merge repository doctors instead of replacing the local picker list",
);
assert.match(
  appSource.match(/accountsBody\.addEventListener\("click"[\s\S]*?if \(adminTab\)[\s\S]*?return;/)?.[0] || "",
  /nextTab === "system" \|\| currentAdminTab === "system"[\s\S]*adminConsoleOpen = false/,
  "opening or leaving the system admin tab should collapse the console by default",
);
assert.match(
  appSource.match(/function setStatus[\s\S]*?function removeSupersededStatusMessages/)?.[0] || "",
  /adminConsoleOpen && currentAdminTab === "system" && !adminConsoleLoading[\s\S]*appendLiveAdminConsoleMessage[\s\S]*refreshAdminConsoleMarkupIfVisible/,
  "status messages should append live to the admin console while it is open on the system tab",
);
assert.match(
  appSource.match(/async function openAccountsSurface[\s\S]*?async function refreshDoctorProfileFileRefs/)?.[0] || "",
  /adminConsoleOpen = false/,
  "opening the admin modal should collapse the console by default",
);
assert.match(
  appSource.match(/function closeAccountsModal[\s\S]*?function loadAccountState/)?.[0] || "",
  /adminConsoleOpen = false/,
  "closing the admin modal should collapse the console",
);
assert.match(
  appSource.match(/async function refreshCalendarStoreStatus[\s\S]*?async function toggleAdminConsole/)?.[0] || "",
  /renderFileSurfaces\(\);/,
  "roster status refreshes should re-render file surfaces for the active admin tab",
);
assert.doesNotMatch(
  appSource.match(/function setRosterSyncState[\s\S]*?function finishRosterSync/)?.[0] || "",
  /currentAdminTab === "system"/,
  "roster sync state updates should not be limited to the system admin tab",
);
assert.match(
  appSource.match(/async function refreshCreatorCalendarAfterFileChange[\s\S]*?async function rosterDoctorsFromSelectedFiles/)?.[0] || "",
  /syncCreatorDoctorPickerWithRemainingRosters[\s\S]*loadCloudCalendarEvents[\s\S]*pollCalendarAfterRosterChange/,
  "creator roster imports should refresh the doctor picker before requesting a calendar snapshot reload",
);
assert.match(
  appSource.match(/async function bootstrapImports[\s\S]*?function snapshotHasUnresolvablePreviewEvents/)?.[0] || "",
  /isViewingCreatorAccount\(\) && cloudAvailable[\s\S]*refreshCreatorSnapshotInBackground/,
  "stale creator bootstrap should refresh from cloud snapshots instead of browser file re-parse",
);
assert.match(
  appSource.match(/async function refreshSnapshotInBackground[\s\S]*?async function bootstrapApp/)?.[0] || "",
  /isViewingCreatorAccount\(\) && cloudAvailable[\s\S]*refreshCreatorSnapshotInBackground/,
  "background snapshot refresh should delegate creator calendars to the cloud refresh helper",
);
assert.match(
  appSource.match(/function beginCalendarTransition[\s\S]*?function calendarTransitionStillCurrent/)?.[0] || "",
  /cancelCalendarImportPoll\(\)/,
  "calendar transitions should cancel in-flight roster import polling",
);
assert.match(
  appSource.match(/async function updatePreview[\s\S]*?async function buildBrowserPreviewData/)?.[0] || "",
  /preserveVisiblePreview = Boolean\(latestPreview && isViewingCreatorAccount\(\) && cloudAvailable\)/,
  "creator cloud preview failures should keep the last visible calendar on screen",
);
assert.match(
  appSource.match(/async function buildBrowserPreviewData[\s\S]*?function rebuildClientPreview/)?.[0] || "",
  /OWNER_DOCTOR_KEY && canUseCreatorDoctorSwitcher\(\)/,
  "creator owner route should not require the owner name to appear in parsed roster files",
);
assert.match(
  appSource.match(/async function analyzeFiles[\s\S]*?async function parseCurrentRosterForm/)?.[0] || "",
  /isViewingCreatorAccount\(\) && cloudAvailable[\s\S]*updatePreview/,
  "creator cloud workspaces should skip browser preview rebuild after file analysis",
);
assert.match(
  appSource.match(/async function enterDoctorProfileView[\s\S]*?async function exitDoctorProfileView/)?.[0] || "",
  /rememberCreatorCalendarSourceRefs\(\)/,
  "doctor-profile entry should preserve creator import refs before switching context",
);
assert.doesNotMatch(
  appSource.match(/async function enterDoctorProfileView[\s\S]*?async function exitDoctorProfileView/)?.[0] || "",
  /if \(!renderedCachedSnapshot\) \{[\s\S]*selectedFiles = \[\]/,
  "doctor-profile entry should not clear creator import files when cache is unavailable",
);
assert.match(
  appSource.match(/function restoreCreatorImportFilesIfNeeded[\s\S]*?function applyLoadedCalendarFileRefs/)?.[0] || "",
  /pendingRemovedImportIds[\s\S]*creatorCalendarSourceFileRefs/,
  "creator import restore should skip files pending removal",
);
assert.match(
  appSource.match(/function applyLoadedCalendarFileRefs[\s\S]*?function rosterStoreFileToClientEntry/)?.[0] || "",
  /pendingRemovedImportIds[\s\S]*!selectedFiles\.length[\s\S]*restoreCreatorImportFilesIfNeeded/,
  "loaded calendar snapshots should preserve existing creator files when snapshot refs are empty",
);
assert.match(
  appSource.match(/function renderCachedCalendarSnapshotForContext[\s\S]*?function saveWorkspaceSnapshotForEmail/)?.[0] || "",
  /applyLoadedCalendarFileRefs\(cached\)/,
  "cached calendar rendering should preserve creator import files when cache refs are empty",
);
assert.match(
  appSource.match(/async function returnToCreatorAccount[\s\S]*?async function clearLocalWorkspace/)?.[0] || "",
  /restoreCreatorImportFilesIfNeeded\(\)[\s\S]*renderFileSurfaces\(\)/,
  "returning to creator should restore and render import files before background validation",
);
assert.match(
  appSource.match(/async function syncCreatorFileListFromStore[\s\S]*?function applyLoadedCalendarFileRefs/)?.[0] || "",
  /calendarStoreStatus\?\.files[\s\S]*mergeSelectedFilesWithRosterStoreStatus[\s\S]*removeMissingFromStore/,
  "creator file sync should reconcile against D1 without restoring stale refs after a successful delete",
);
assert.match(
  appSource.match(/async function refreshCalendarStoreStatus[\s\S]*?async function toggleAdminConsole/)?.[0] || "",
  /if \(!options\.silent\)[\s\S]*calendarStoreStatusError/,
  "silent roster status refreshes should keep the last known file list when the server is overloaded",
);
assert.match(
  appSource.match(/async function removeStoredImport[\s\S]*?function loadConflictSelections/)?.[0] || "",
  /creatorCalendarSourceFileRefs = creatorCalendarSourceFileRefs\.filter/,
  "roster deletion should drop removed files from remembered creator import refs",
);
assert.match(
  appSource,
  /function visibleSnapshotIsCurrent\(/,
  "account switching should expose a shared visible-snapshot revision helper",
);
assert.match(
  appSource.match(/async function returnToCreatorAccount[\s\S]*?async function clearLocalWorkspace/)?.[0] || "",
  /visibleSnapshotIsCurrent\(\{ requireNotStale: true \}\)[\s\S]*cachedRevision:[\s\S]*allowInlineBuild: false[\s\S]*preserveExistingSnapshot: true/,
  "creator return validation should skip or lightweight-check calendar loads when the visible snapshot is current",
);
assert.doesNotMatch(
  appSource.match(/async function returnToCreatorAccount[\s\S]*?async function clearLocalWorkspace/)?.[0] || "",
  /await syncCreatorFileListFromStore\(\)/,
  "creator return validation should not block account switching on D1 file-status refresh",
);
assert.doesNotMatch(
  appSource.match(/async function bootstrapImports[\s\S]*?function snapshotHasUnresolvablePreviewEvents/)?.[0] || "",
  /await syncCreatorFileListFromStore\(\)/,
  "bootstrap imports should defer D1 file-status refresh instead of blocking workspace render",
);
assert.match(
  appSource.match(/async function enterUserAccount[\s\S]*?async function enterDoctorProfileView/)?.[0] || "",
  /reportBackgroundValidationError\(error, \{[\s\S]*preserveRenderedSnapshot: true/,
  "claimed-account background validation should suppress overload errors when cached calendar remains valid",
);
assert.match(
  appSource.match(/async function enterDoctorProfileView[\s\S]*?async function exitDoctorProfileView/)?.[0] || "",
  /reportBackgroundValidationError\(error, \{[\s\S]*preserveRenderedSnapshot: true/,
  "doctor-profile background validation should suppress overload errors when cached calendar remains valid",
);
assert.match(
  appSource.match(/async function refreshCreatorSnapshotInBackground[\s\S]*?async function refreshAvailableDoctorsAfterRosterChange/)?.[0] || "",
  /preserveExistingSnapshot: true[\s\S]*allowInlineBuild: false[\s\S]*cachedRevision:/,
  "creator background snapshot refresh should pass cachedRevision for lightweight server checks",
);
assert.match(
  await readFile(new URL("../public/index.html", import.meta.url), "utf8"),
  /id="rosterImportOverlay"/,
  "index should include the roster import overlay shell",
);
assert.match(
  appSource.match(/function isRosterFileStatusHealthy[\s\S]*?function reconcileRosterSyncStates/)?.[0] || "",
  /retainedSourceOnly === true[\s\S]*function isLocalRosterFileSyncedToD1/,
  "retained R2-only roster files should not count as synced derived D1 files",
);
assert.match(
  appSource.match(/async function refreshCreatorCalendarAfterFileChange[\s\S]*?async function refreshAvailableDoctorsAfterRosterChange/)?.[0] || "",
  /unsyncedLocalEntries[\s\S]*saveSelectedRosterFilesToD1[\s\S]*loadCloudCalendarEvents/,
  "creator roster changes should save unsynced local files then reload from D1",
);
assert.match(
  appSource.match(/async function refreshCreatorCalendarAfterFileChange[\s\S]*?async function rosterDoctorsFromSelectedFiles/)?.[0] || "",
  /cachedRevision: ""[\s\S]*loaded && currentSnapshot && !currentSnapshotStale[\s\S]*pollCalendarAfterRosterChange/,
  "roster imports should keep polling until a genuinely current snapshot is returned",
);
assert.match(
  stateSource.match(/async function runCoreDerivedRosterSave[\s\S]*?async function runDeferredDerivedRosterSave/)?.[0] || "",
  /deferCanonicalDoctorRefresh[\s\S]*scheduleSnapshotWarmupForSourceTypes/,
  "post-save snapshot warmup should wait for canonical doctor refresh",
);
assert.match(
  appSource.match(/async function loadCloudCalendarEvents[\s\S]*?function cloudCalendarEventRange/)?.[0] || "",
  /data\.snapshotStale !== true[\s\S]*if \(!currentSnapshotStale\)[\s\S]*saveCalendarSnapshotCacheForContext/,
  "stale server snapshots should not be relabelled or persisted as current browser snapshots",
);
assert.match(
  appSource.match(/async function saveCloudStateNow[\s\S]*?async function replaceActiveRostersWithCurrentUploads/)?.[0] || "",
  /isDeleteSave[\s\S]*snapshotPayload = isDeleteSave \? null/,
  "roster deletion saves should skip uploading a full workspace snapshot",
);
assert.match(
  appSource.match(/async function saveSelectedRosterFilesToD1[\s\S]*?function mergeLightweightRosterStatus/)?.[0] || "",
  /failStep = "parse"[\s\S]*buildDerivedCalendarFilePayload[\s\S]*retainRosterSource/,
  "roster save should parse before attempting source retention",
);
assert.match(
  appSource.match(/function slimDerivedCalendarRequest[\s\S]*?async function saveDerivedCalendarFilePayload/)?.[0] || "",
  /eventsByDoctor: _eventsByDoctor[\s\S]*issuesByDoctor: _issuesByDoctor/,
  "derived calendar save requests should strip full event payloads from non-chunk phases",
);
assert.match(
  appSource.match(/async function saveCloudStateNow[\s\S]*?async function replaceActiveRostersWithCurrentUploads/)?.[0] || "",
  /removedImportIds[\s\S]*buildCloudStateWithoutRosterSync/,
  "roster deletion saves should skip redundant D1 roster sync",
);
assert.match(
  appSource.match(/async function replaceActiveRostersWithCurrentUploads[\s\S]*?async function reparseRosterFile/)?.[0] || "",
  /retainedRosterEntriesFromStatus[\s\S]*Rebuild requires at least one retained roster file/,
  "rebuild-all should use retained roster sources and refuse an empty retained roster set",
);
assert.match(
  appSource.match(/async function reparseRosterFile[\s\S]*?async function refreshCalendarStoreStatus/)?.[0] || "",
  /calendarStoreStatus\?\.files[\s\S]*ensureRosterEntrySource\(entry\)/,
  "per-file refresh should be able to hydrate status-only files from retained R2 source",
);
assert.match(
  d1CalendarSource.match(/async function runTransactionalBatch[\s\S]*?function chunkRowsForBindLimit/)?.[0] || "",
  /runStatementBatches\(db, statements, D1_MAX_BATCH_STATEMENTS\)/,
  "D1 roster imports should sub-batch large db.batch calls",
);
assert.match(
  d1CalendarSource.match(/function dailyPresenceInsertStatements[\s\S]*?async function rebuildDailyPresenceForFile/)?.[0] || "",
  /chunkRowsForBindLimit\(rows, 5/,
  "daily presence inserts should stay under D1 bind-parameter limits",
);
assert.match(
  stateSource.match(/if \(action === "saveDerivedCalendarFile"\)[\s\S]*?if \(action === "uploadRawRosterFile"\)/)?.[0] || "",
  /await runCoreDerivedRosterSave\(context, saveJob\)/,
  "derived roster saves should write core D1 data synchronously before responding",
);
assert.match(
  stateSource.match(/async function runCoreDerivedRosterSave[\s\S]*?async function runDeferredDerivedRosterSave/)?.[0] || "",
  /context\.waitUntil\(postSave\(\)/,
  "derived roster saves should defer only post-save indexing to waitUntil",
);
assert.match(
  stateSource.match(/if \(action === "calendarStoreStatus"\)[\s\S]*?if \(action === "appendConsoleMessage"\)/)?.[0] || "",
  /includeAvailableDoctors === true[\s\S]*repositoryDoctorCandidates/,
  "calendar status should optionally return available doctors for switcher refresh",
);
assert.match(
  appSource.match(/function isRosterFileStatusHealthy[\s\S]*?function reconcileRosterSyncStates/)?.[0] || "",
  /function isRosterFileStatusHealthy/,
  "roster sync reconciliation should detect healthy D1 file status",
);
assert.match(
  appSource.match(/function rosterSyncLabel[\s\S]*?function rosterSyncSummary/)?.[0] || "",
  /state\.status === "failed"[\s\S]*isRosterFileStatusHealthy\(statusFile\)/,
  "roster sync labels should clear stale failed states when D1 status is healthy",
);
assert.match(
  appSource.match(/async function refreshCreatorCalendarAfterFileChange[\s\S]*?function normalizeSavedExportRange/)?.[0] || "",
  /refreshAvailableDoctorsAfterRosterChange\(\{/,
  "creator roster changes should refresh switcher doctors after add or delete",
);
assert.match(
  d1CalendarSource.match(/function bulkInsertEventStatements[\s\S]*?function bulkInsertIssueStatements/)?.[0] || "",
  /chunkRowsForBindLimit\(rows, 17, D1_MAX_BIND_PARAMS\)/,
  "D1 roster event inserts should batch multiple rows per statement",
);
assert.match(
  d1CalendarSource.match(/function bulkInsertIssueStatements[\s\S]*?async function runTransactionalBatch/)?.[0] || "",
  /chunkRowsForBindLimit\(rows, 14, D1_MAX_BIND_PARAMS\)/,
  "D1 roster issue inserts should batch multiple rows per statement",
);
assert.match(
  d1CalendarSource.match(/export async function startDerivedRosterFileSave[\s\S]*?export async function appendDerivedRosterFileEvents/)?.[0] || "",
  /startDerivedRosterFileSave/,
  "large roster saves should support phased D1 writes",
);
assert.match(
  d1CalendarSource.match(/export async function compareDerivedRosterFiles[\s\S]*?export async function promoteVerifiedStagedRosterFile/)?.[0] || "",
  /doctor_key[\s\S]*?sameRosterOccurrence[\s\S]*?approvedOmissionCount/,
  "reparse comparison must distinguish source occurrences from corrected parser event ids",
);
assert.match(
  d1CalendarSource.match(/export async function promoteVerifiedStagedRosterFile[\s\S]*?export async function deleteDerivedRosterFile/)?.[0] || "",
  /comparison\.removedCount > 0[\s\S]*?reason: "unreviewed-removals"/,
  "a retained-file reparse must refuse activation when it would remove events",
);
assert.match(
  automationDerivedSource,
  /active: false,[\s\S]*?staged: true,[\s\S]*?replacesFileId:/,
  "automation writes must remain staged until their finish phase",
);
assert.match(
  stateSource.match(/async function queueActiveParserRuleReparse[\s\S]*?async function queueAutomatedSourceReprocess/)?.[0] || "",
  /stagingFileId[\s\S]*?fileId: stagingFileId,[\s\S]*?sourceFileId: file\.id/,
  "parser-rule reparses must retain an immutable source file and use a staging destination",
);
assert.match(
  safeStagedActivationMigrationSource,
  /parser_version = 'legacy-unverified'/,
  "pre-shared-parser roster rows must be marked legacy rather than current",
);
assert.match(
  appSource.match(/async function saveDerivedCalendarFilePayload[\s\S]*?async function saveSelectedRosterFilesToD1/)?.[0] || "",
  /phase: "start"[\s\S]*phase: "events"[\s\S]*phase: "finish"/,
  "large roster uploads should save to D1 in client-driven chunks",
);
assert.match(
  d1CalendarSource.match(/INSERT INTO account_claims[\s\S]*?bind\(\.\.\.chunk\.flat\(\)\)\.run\(\)/)?.[0] || "",
  /ON CONFLICT\(email, source_type, doctor_key\) DO UPDATE/,
  "account claim writes should be idempotent for duplicate or concurrent claim repairs",
);
assert.match(
  stateSource.match(/async function calendarStoreStatus[\s\S]*?function summarizeExpectedRosterFiles/)?.[0] || "",
  /queryRawRosterFiles[\s\S]*retainedOnlyFiles[\s\S]*retainedSourceOnly/,
  "calendar status should include retained R2 source pointers without derived rows",
);
assert.match(
  appSource.match(/async function loginWithEmail[\s\S]*?async function restoreCloudState/)?.[0] || "",
  /restoreCloudState\(\{[\s\S]*deferContext: true[\s\S]*deferSnapshotPersistence: true[\s\S]*responseMode: "fast"/,
  "login should request the fast phased cloud-state response",
);
assert.match(
  appSource.match(/async function bootstrapApp[\s\S]*?function setStatus/)?.[0] || "",
  /renderCachedCalendarSnapshotForContextAsync[\s\S]*hideLoadingScreen\(\)[\s\S]*restoreCloudState\(\{[\s\S]*responseMode: "fast"[\s\S]*queuePostLoginHydration/,
  "remembered sessions should paint the browser calendar before authentication refresh and workspace hydration",
);
assert.match(
  appSource.match(/function compactCalendarSnapshotCacheStore[\s\S]*?function calendarSnapshotContext/)?.[0] || "",
  /MAX_HOT_SNAPSHOT_CACHE_ENTRIES[\s\S]*localStorage\.setItem\(CALENDAR_SNAPSHOT_CACHE_KEY, JSON\.stringify\(compactStore\)\)/,
  "the synchronous calendar hot cache should be proactively bounded",
);
assert.match(
  appSource.match(/async function pruneStoredCalendarSnapshots[\s\S]*?function queueStoredCalendarSnapshotPersist/)?.[0] || "",
  /MAX_STORED_SNAPSHOT_CACHE_AGE_MS[\s\S]*MAX_STORED_SNAPSHOT_CACHE_ENTRIES[\s\S]*requestIdleCallback/,
  "the larger switcher cache should be pruned only through idle maintenance",
);
assert.match(
  appSource.match(/async function loginWithEmail[\s\S]*?async function restoreCloudState/)?.[0] || "",
  /renderLoginState\(\);\s*closeLoginModal\(\);[\s\S]*queueDeferredAccountContextLoad[\s\S]*queuePostLoginHydration\(\{[\s\S]*includeBootstrap: true[\s\S]*allowInlineBuild: !renderedCachedSnapshot/,
  "successful login should reveal the shell before background workspace hydration completes",
);
assert.match(
  appSource.match(/async function hydrateAuthenticatedWorkspace[\s\S]*?function markLoginPhase/)?.[0] || "",
  /forceCalendarRefresh[\s\S]*allowInlineBuild: options\.allowInlineBuild !== false[\s\S]*preserveExistingSnapshot: true/,
  "post-login hydration should refresh calendar data in the background without blanking the visible snapshot",
);
assert.match(
  appSource.match(/async function loadCloudCalendarEvents[\s\S]*?function cloudCalendarEventRange/)?.[0] || "",
  /allowInlineBuild: options\.allowInlineBuild !== false[\s\S]*!data\.snapshot && options\.preserveExistingSnapshot === true/,
  "non-inline calendar refreshes should preserve existing cached snapshots on server cache misses",
);
assert.match(
  appSource.match(/async function warmInsightData[\s\S]*?function syncActionState/)?.[0] || "",
  /fetchRosterOverlapDoctors[\s\S]*allowFallback: true/,
  "post-render insight warmup should only prefetch the When doctor list after first paint",
);
assert.doesNotMatch(
  appSource.match(/async function warmInsightData[\s\S]*?function syncActionState/)?.[0] || "",
  /fetchRosterInsightRows/,
  "post-render insight warmup must not fan out broad coworker SQL requests",
);
assert.match(
  appSource.match(/async function fetchRosterOverlapDoctors[\s\S]*?function renderRosterInsightUnavailable/)?.[0] || "",
  /readPersistentRosterOverlapDoctors\(cacheKey\)[\s\S]*writePersistentRosterOverlapDoctors\(cacheKey, data\.doctors\)/,
  "When doctor-list queries should use persistent revision-keyed browser cache",
);
assert.match(
  appSource.match(/function rosterOverlapDoctorCacheKey[\s\S]*?function resetVisibleInsightWarmCache/)?.[0] || "",
  /insightWarmBaseKey\(\)[\s\S]*ROSTER_OVERLAP_DOCTOR_CACHE_KEY[\s\S]*currentCalendarRevision/,
  "When doctor-list cache keys should include the active calendar revision",
);
assert.match(
  appSource.match(/async function renderInsightsModal[\s\S]*?async function ensureInsightRosterAnalysis/)?.[0] || "",
  /renderWhenInsightLoading\(\);\s*showInsightsModal\(\);\s*await renderWhenInsight/,
  "When insight modal should open immediately before SQL-backed rows finish loading",
);
assert.match(
  appSource.match(/async function renderWhenInsight[\s\S]*?function renderWhenInsightResult/)?.[0] || "",
  /isCurrentInsightRender\(renderRunId, "when"\)[\s\S]*isCurrentInsightRender\(renderRunId, "when"\)/,
  "When insight async renders should ignore stale results after close or selection changes",
);
assert.match(
  appSource.match(/async function fetchRosterInsightRows[\s\S]*?async function fetchRosterOverlapDoctors/)?.[0] || "",
  /allowFallback = true[\s\S]*allowFallback/,
  "explicit roster insight actions should be able to request the correctness fallback",
);
assert.doesNotMatch(
  stateSource.match(/if \(action === "queryRosterInsights"\)[\s\S]*?if \(action === "queryRosterOverlapDoctors"\)/)?.[0] || "",
  /waitUntil|scheduleDailyPresenceRepair|rebuildDailyPresenceForActiveFiles/,
  "coworker insight queries must not schedule full daily-presence repairs",
);
assert.doesNotMatch(
  stateSource.match(/if \(action === "queryRosterOverlapDoctors"\)[\s\S]*?if \(action === "loadCalendarEvents"\)/)?.[0] || "",
  /waitUntil|scheduleDailyPresenceRepair|rebuildDailyPresenceForActiveFiles/,
  "overlap doctor queries must not schedule full daily-presence repairs",
);
assert.match(
  appSource.match(/async function restoreCloudState[\s\S]*?async function hydrateAuthenticatedWorkspace/)?.[0] || "",
  /applyCloudStateData\(data, \{[\s\S]*deferContext[\s\S]*deferSnapshotPersistence[\s\S]*skipSnapshotCacheWriteIfCurrent[\s\S]*\}\)[\s\S]*recordLoginServerTimings/,
  "restoreCloudState should support the fast-login phased apply path and record server timings for every account",
);
assert.match(
  stateSource.match(/if \(action === "login"\)[\s\S]*?const account = await verifyD1Account/)?.[0] || "",
  /registryLookupMs[\s\S]*r2ReadMs[\s\S]*serverTotalMs[\s\S]*validationDeferred[\s\S]*Server-Timing/,
  "fast login responses should expose server and R2 timing instrumentation",
);
assert.match(
  appSource.match(/function markLoginPhase[\s\S]*?function markAccountSwitchPhase/)?.[0] || "",
  /recordLoginServerTimings[\s\S]*firstCalendarPaintCommitted/,
  "client timings should include server phases and a browser-committed calendar paint",
);
assert.match(
  appSource.match(/function renderSystemAdminCard[\s\S]*?function renderAccountHospitalLocationsCard/)?.[0] || "",
  /renderLoginPerformanceCard[\s\S]*firstCalendarPaintCommitted[\s\S]*serverTotalMs[\s\S]*registryLookupMs[\s\S]*r2ReadMs/,
  "Admin System should present the latest client, server, registry, and R2 login timings",
);
assert.match(
  appSource.match(/function queuePostLoginSnapshotRefresh[\s\S]*?function markLoginPhase/)?.[0] || "",
  /allowInlineBuild: false[\s\S]*currentSnapshotStale[\s\S]*preserveScroll: true[\s\S]*backgroundCalendarUpdated/,
  "stale-while-revalidate should replace a rebuilt calendar after paint without moving the visible calendar",
);
assert.match(
  appSource.match(/function renderWorkspaceFromSnapshot[\s\S]*?async function ensureSelectedFilesLoaded/)?.[0] || "",
  /preserveScroll[\s\S]*pageY[\s\S]*previewY[\s\S]*pendingPreviewSnapToToday = options\.preserveScroll !== true/,
  "background calendar replacement should preserve mobile and desktop scroll positions",
);
assert.match(
  appSource.match(/function saveCalendarSnapshotCacheForContext[\s\S]*?function invalidateCalendarSnapshotCache/)?.[0] || "",
  /skipIfRevisionMatches[\s\S]*requestIdleCallback[\s\S]*requestAnimationFrame/,
  "snapshot cache persistence should be deferrable and skip redundant warm-login rewrites",
);
assert.doesNotMatch(
  stateSource,
  /ADMIN_ISSUE_DISMISS_PREFIX|ADMIN_ISSUE_IGNORE_PREFIX|PARSER_EXTENSION_RULES_KEY|PARSER_RULE_SUGGESTIONS_KEY|loadParserExtensionRules\(|saveParserExtensionRules\(/,
  "D1-only state routes should not retain dead KV-era helper scaffolding",
);
assert.match(
  appSource.match(/function buildResolvedPreviewEvents[\s\S]*?function latestPreviewEventsByIdentity/)?.[0] || "",
  /activeCustomEventIds[\s\S]*!activeCustomEventIds\.has\(event\.id\)[\s\S]*previewCustomEventIds[\s\S]*customEventsMaterialized === true && previewCustomEventIds\.has\(event\.id\)/,
  "D1-loaded previews should merge newly added local custom events and drop materialized custom events removed from active state",
);
assert.match(
  appSource.match(/function openReviewModal[\s\S]*?function closeReviewModal/)?.[0] || "",
  /isCustomPreviewEvent\(previewEvent\)[\s\S]*openCustomEventModal/,
  "materialized custom events should open in the custom-event editor before generic review handling",
);
assert.match(
  appSource.match(/function deletePreviewEvent[\s\S]*?function resetImportedEvent/)?.[0] || "",
  /isCustomPreviewEvent\(event\)[\s\S]*ensureEditableCustomEvent\(event\)[\s\S]*removeCustomEventForActiveCalendar/,
  "materialized custom events should be rehydrated before deletion",
);
assert.match(
  appSource.match(/function renderWorkspaceFromSnapshot[\s\S]*?async function ensureSelectedFilesLoaded/)?.[0] || "",
  /applySessionState[\s\S]*reconcileMaterializedPreviewCustomEvents\(\)/,
  "rendering a D1 snapshot should reconcile materialized custom events into editable state",
);
assert.match(
  appSource.match(/function renderCalendarStoreCard[\s\S]*?function renderAdminConsoleMarkup/)?.[0] || "",
  /serverStatusComplete[\s\S]*serverSyncedCount[\s\S]*roster file\$\{serverSyncedCount === 1 \? "" : "s"\} synced/,
  "healthy server roster status should drive the System-card synced count",
);
assert.match(
  appSource.match(/function renderCalendarStoreCard[\s\S]*?function renderAdminConsoleMarkup/)?.[0] || "",
  /Sync issue detected: \$\{serverSyncedCount\}\/\$\{serverExpectedCount\}/,
  "System-card sync issues should compare server counts rather than empty local file handles",
);
assert.match(appSource, /data-admin-user-seniority-filter/, "admin users should expose a seniority filter");
assert.match(
  appSource.match(/function renderAdminFilesMarkup[\s\S]*?function adminRosterTerms/)?.[0] || "",
  /Auto-sync[\s\S]*Manual imports[\s\S]*Previously imported files/,
  "admin files should group automated and manual roster files into the new sections",
);
assert.match(
  appSource.match(/function renderAdminAutoSyncRow[\s\S]*?function adminLatestTermFile/)?.[0] || "",
  /data-refresh-automated-source[\s\S]*Current term[\s\S]*Next term/,
  "automated roster rows should report current/next term files and expose a refresh control",
);
assert.match(
  stateSource.match(/async function calendarStoreStatus[\s\S]*?function summarizeExpectedRosterFiles/)?.[0] || "",
  /sourceId: String\(file\.sourceId[\s\S]*startDate: String\(file\.startDate/,
  "calendar store status should expose roster source and coverage metadata for Admin grouping",
);
assert.doesNotMatch(
  appSource.match(/async function deleteAccount[\s\S]*?function deleteLocalAccountData/)?.[0] || "",
  /if \(creatorCanDelete\) await loadServerUsers\(\);\s*closeAccountsModal\(\);/,
  "deleting another account should not close the Admin modal",
);
assert.match(
  appSource.match(/function renderCalendarStoreCard[\s\S]*?function renderAdminConsoleMarkup/)?.[0] || "",
  /hasUsableStatus[\s\S]*missingSelectedFiles = hasUsableStatus/,
  "roster database status card should only list missing files when status is usable",
);
assert.match(
  appSource.match(/function renderCalendarStoreCard[\s\S]*?function renderAdminConsoleMarkup/)?.[0] || "",
  /Status check pending — roster files may already be synced/,
  "unchecked roster database status should not claim files are unsynced",
);
assert.match(
  appSource.match(/async function refreshCalendarStoreStatus[\s\S]*?async function toggleAdminConsole/)?.[0] || "",
  /calendarStoreRequestWithRetry\("calendarStoreStatus"/,
  "roster status checks should retry transient SQL store failures",
);
assert.match(
  appSource.match(/data-refresh-calendar-store[\s\S]*?data-replace-active-rosters/)?.[0] || "",
  /refreshCalendarStoreStatus\(\{ silent: false, syncSwitcher: true \}\)/,
  "Check status should refresh the creator switcher after a successful status check",
);
assert.match(
  appSource.match(/function applyAvailableRosterDoctorsFromData[\s\S]*?async function applyCloudStateSnapshot/)?.[0] || "",
  /!incomingAvailableDoctors\.length && isCreatorAuthenticated\(\)[\s\S]*availableRosterDoctors = incomingAvailableDoctors/,
  "creator hydration should preserve locally known doctors when a lightweight response omits the list",
);
assert.match(
  appSource.match(/async function loadServerUsers[\s\S]*?async function refreshCalendarStoreStatus/)?.[0] || "",
  /applyAuthoritativeAvailableDoctors\(data\.availableDoctors\)/,
  "loading admin users should apply the authoritative repository doctor list",
);
assert.match(
  appSource.match(/async function bootstrapImports[\s\S]*?function snapshotHasUnresolvablePreviewEvents/)?.[0] || "",
  /syncCreatorDoctorPickerWithRemainingRosters\(\{ snapshotDoctors: currentSnapshot\.doctorOptions/,
  "creator bootstrap should sync the switcher from roster files and snapshot doctors before rendering",
);
assert.match(
  appSource.match(/async function syncCreatorDoctorPickerWithRemainingRosters[\s\S]*?async function pollCalendarAfterRosterChange/)?.[0] || "",
  /snapshotDoctors[\s\S]*availableDoctorsFromRosterDoctorOptions\(snapshotDoctors\)/,
  "creator doctor picker sync should fall back to snapshot doctors when local parse is unavailable",
);
assert.match(
  await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"),
  /lightweight: body\?\.lightweight === true/,
  "calendar status API should accept a lightweight status request",
);
assert.match(
  await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"),
  /async function syncRosterRepositoryToKeepFileIds[\s\S]*await refreshCanonicalDoctors\(db\)/,
  "repository sync should refresh canonical doctors before warming replacement snapshots",
);
assert.match(
  await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"),
  /const lightweight = options\.lightweight === true[\s\S]*if \(!lightweight && selectedDoctorKey\)/,
  "lightweight calendar status should skip selected-doctor event counts",
);
assert.match(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/async function canonicalDoctorOptionForProfile[\s\S]*?async function doctorProfileImportRefs/)?.[0] || "",
  /queryCanonicalDoctors[\s\S]*queryRosterFileDoctorsForKeys\(db, requestedKeys\)/,
  "doctor profile load should expand the requested canonical identity and query only its roster keys",
);
assert.match(
  (await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"))
    .match(/async function doctorProfileDiagnostics[\s\S]*?async function doctorProfileImportRefs/)?.[0] || "",
  /if \(!diagnostics\.length\)[\s\S]*queryRosterFileDoctors\(db\)[\s\S]*resolveCanonicalDoctorOptionForKey/,
  "doctor profile load should reserve full canonical rebuilds for the rare targeted-lookup miss path",
);
assert.match(appSource, /data-replace-active-rosters/, "creator UI should expose a roster recovery action");
assert.match(appSource, /<strong>Roster database<\/strong>/, "system card should use plain roster-database language");
assert.match(appSource, /source file\$\{retainedSourceTotal === 1 \? \"\" : \"s\"\} retained/, "system card should report retained raw source coverage");
assert.match(appSource, />Check status<\/button>/, "system card should expose a non-mutating status check");
assert.match(appSource, />Rebuild all retained rosters<\/button>/, "advanced recovery should expose an explicit full rebuild action");
assert.doesNotMatch(
  await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8"),
  /loadRepositoryIndex|loadSnapshotRecord|persistSnapshotRecord|storeSnapshotForAccount|storeSnapshotForDoctorProfile/,
  "D1 serving code should not retain repository-index or snapshot serving helpers",
);
assert.doesNotMatch(
  appSource.match(/async function refreshCalendarStoreStatus[\s\S]*?async function toggleAdminConsole/)?.[0] || "",
  /calendarStoreStatus = \{ unavailable: true \}/,
  "failed status checks should not erase the last valid D1 status",
);
assert.match(
  appSource.match(/async function saveSelectedRosterFilesToD1[\s\S]*?function emptyRosterPersistenceSummary/)?.[0] || "",
  /entriesToSave = options\.force === true[\s\S]*entries\.filter\(\(entry\) => !isLocalRosterFileSyncedToD1\(entry\) \|\| failedIds\.has\(entry\.id\)\)/,
  "ordinary roster sync should process only missing or failed files",
);
assert.match(
  appSource.match(/async function saveSelectedRosterFilesToD1[\s\S]*?function emptyRosterPersistenceSummary/)?.[0] || "",
  /if \(!summary\.complete\)[\s\S]*invalidateCalendarSnapshotCachesForChangedRosterFiles\(entriesToSave\)/,
  "roster sync should invalidate snapshot cache only after a successful save",
);
assert.match(appSource, /function rosterSyncLabel/, "file cards should expose live roster sync labels");
assert.match(appSource, /Missing \/ unresolved shift codes/, "system card should expose unresolved shift-code review");
assert.match(appSource, /data-add-manual-shift-code/, "missing shift-code heading should expose a manual Add action");
assert.match(appSource, /parserRuleSeniorityOption/, "shift-code editor should expose multi-seniority selection");
assert.match(appSource, /function selectedParserRuleSeniorities/, "shift-code editor should save selected seniorities as a batch");
assert.match(appSource, /function openManualParserRuleModal/, "shift-code editor should support manual rule creation without an existing issue");
assert.match(appSource, /parserRuleExistsForIssue/, "system shift-code review should hide only issues resolved by active rules");
assert.match(appSource, /matchingParserRuleGroup/, "saved shift-code rules should reopen with equivalent seniorities selected");
assert.match(appSource, /<span class="who-team-role">/, "Who role labels should render separately from doctor buttons");
assert.doesNotMatch(appSource, /data-who-role-shift-code|handleWhoRoleRuleClick/, "Who role labels should not be clickable shift-code controls");
assert.match(appSource, /isWhoNightIcShift[\s\S]*nightIcRank/, "Night IC shifts should rank above other night staff in Who lists");
assert.match(appSource, /function whoDisplayTeamLabel[\s\S]*Night main team[\s\S]*Night Hub[\s\S]*Night SSU/, "Night teams should use explicit main, hub, and SSU labels");
assert.match(appSource, /night main team", "night hub", "night ssu"/, "Night Hub should sort between Night main team and Night SSU");
assert.match(appSource, /Night Hub/, "Hub night shift-code rules should preview as Night Hub");
assert.match(stateSource, /facilityKey,[\s\S]*isFacilityOverviewWorkingEvent[\s\S]*facilityKey[\s\S]*DDH[\s\S]*hith\|vhh/, "DDH HITH and VHH roster notes should be excluded from the On shift response only");
assert.match(appSource, /function renderFacilityOverviewDdhNightPeriod[\s\S]*Night SR[\s\S]*Main team[\s\S]*SSU team/, "DDH nights should render senior registrar, main-team, and SSU blocks in order");
assert.match(appSource, /function renderFacilityOverviewMmcNightPeriod[\s\S]*showSpecialTimes: false[\s\S]*Night SR[\s\S]*Hub[\s\S]*SSU[\s\S]*Main team/, "MMC nights should render ordered SR, Hub, SSU, and main-team blocks without times");
assert.match(appSource, /refreshActiveWhoInsightSurfaces/, "saving shift-code rules should refresh active Who insight panels");
assert.match(appSource, /function synthesizeIncompleteShiftCodeIssues/, "derived code-only shift titles should synthesize unresolved shift-code issues");
assert.match(appSource, /parserRuleIgnore/, "shift-code editor should expose persistent ignore mode");
assert.match(appSource, /data-ignore-shift-code/, "missing shift-code queue should open the shared ignore rule flow");
assert.match(appSource, /Ignored shift codes/, "ignored shift codes should remain editable in hospital rule sections");
assert.match(appSource, /parserRuleSeniorityAll/, "shift-code seniority picker should expose an All option");
assert.match(appSource, /function normalizeParserRuleSenioritySelection/, "shift-code seniority picker should keep All and Unknown selections consistent");
assert.match(appSource, /const key = `\$\{item\.source\}\|\$\{sanitizeRuleSeniority\(item\.seniority\)\}\|\$\{item\.code\}`/, "unresolved shift-code grouping should be by hospital, seniority, and code");
assert.match(appSource, /renderUnknownShiftCodeHierarchy/, "unresolved shift-code rows should render as a collapsible hospital and seniority hierarchy");
assert.match(appSource, /data-go-to-unresolved-event/, "unresolved shift-code rows should offer a direct calendar jump for their sampled roster event");
assert.match(appSource, /function openUnresolvedShiftIssueEvent[\s\S]*focusPreviewIssueDate/, "unresolved shift-code jumps should switch doctor and focus the relevant calendar date");
assert.match(appSource, /doctorKey: item\.doctorKey[\s\S]*sampleDate: item\.sampleDate/, "grouped unresolved shift-code rows should retain their sampled doctor and date for the calendar jump");
assert.match(appSource, /data-preview-back-to-shift-codes[\s\S]*function returnToShiftCodeReview/, "a calendar opened from shift-code review should return to the selected code row");
assert.match(appSource, /function closeShiftCodeReviewModal[\s\S]*shiftCodeReviewFilter = \{ query: "", source: "all" \}/, "closing a focused shift-code review should restore the full unfiltered list");
assert.match(appSource, /pendingUnresolvedIssueFocusDate[\s\S]*if \(!pendingUnresolvedIssueFocusDate\)[\s\S]*snapPreviewToCurrentMonth/, "a shift-code event jump should suppress current-month snapping while ordinary calendar entry retains it");
assert.match(styleSource, /is-unresolved-issue-focus[\s\S]*border: 3px solid #c83232/, "the focused unresolved event date should receive a red ring");
assert.match(styleSource, /#parserRuleForm[\s\S]*overflow-y: auto/, "shift-code editor form should scroll vertically when it exceeds available height");
assert.match(appSource, /normalizeDdhParserRuleCodeText/, "DDH shift-code issues should use parser-equivalent label codes");
assert.match(appSource, /seniority !== "Unknown"[\s\S]*some\(\(rule\) => rule\.code === code\)/, "Unknown-seniority shift-code issues should resolve by source/code");
assert.match(appSource, /isKnownResolvedShiftCodeValue/, "derived warnings should not be synthesized for built-in recognised shift labels");
assert.match(appSource, /function shouldShowPreviewIssue/, "Warnings panel should centralize resolved shift-code issue filtering");
assert.match(
  appSource.match(/function shouldShowPreviewIssue[\s\S]*?function pruneResolvedLatestPreviewIssues/)?.[0] || "",
  /isKnownResolvedShiftCodeValue[\s\S]*isShiftCodeResolvedByActiveRules/,
  "Warnings panel should hide review-derived shift-code warnings resolved by built-ins or active rules",
);
assert.match(
  appSource.match(/function buildClientPreviewData[\s\S]*?function synthesizeIncompleteShiftCodeIssues/)?.[0] || "",
  /issues\.filter\(shouldShowPreviewIssue\)/,
  "active preview data should use the shared Warnings issue filter",
);
assert.match(
  appSource.match(/function openParserRuleModalFromPreviewIssue[\s\S]*?async function saveParserRuleFromModal/)?.[0] || "",
  /shouldShowPreviewIssue[\s\S]*That parser warning has already been resolved/,
  "shift-code warning editor should refresh resolved stale cards instead of opening missing issues",
);
assert.match(
  appSource.match(/function rebuildClientPreview[\s\S]*?function buildClientPreviewData/)?.[0] || "",
  /pruneResolvedLatestPreviewIssues\(\)/,
  "preview rendering should prune resolved warnings once review context is current",
);
assert.match(
  calendarMigrationSource,
  /CREATE TABLE IF NOT EXISTS roster_issues/,
  "calendar migration should persist parser diagnostics beside indexed roster events",
);
assert.match(
  d1CalendarSource,
  /INSERT INTO roster_issues/,
  "derived roster saves should persist import-time parser diagnostics",
);
assert.match(
  d1CalendarSource,
  /export async function queryDoctorIssues/,
  "calendar loads should read stored parser diagnostics from D1",
);
assert.match(
  appSource.match(/function synthesizeIncompleteShiftCodeIssues[\s\S]*?function incompleteShiftCodeIssueForReviewItem/)?.[0] || "",
  /if \(baseData\?\.derivedFromD1\) return \[\];/,
  "D1 calendar loads should not synthesize parser warnings from already-indexed events",
);
assert.match(
  appSource.match(/async function reportPreviewIssues[\s\S]*?async function reportAccountError/)?.[0] || "",
  /if \(latestPreview\?\.derivedFromD1\) return;/,
  "D1 calendar loads should not report parser warnings during normal rendering",
);
assert.doesNotMatch(
  appSource.match(/function applyIssueConfig[\s\S]*?function sanitizeParserExtensions/)?.[0] || "",
  /pruneResolvedLatestPreviewIssues\(\)/,
  "parser config application should stay lightweight during login and account switching",
);
assert.match(appSource, /if \(parsedRosterSources\)[\s\S]*await updatePreview\(\)[\s\S]*else if \(latestPreview\)/, "saving parser rules should refresh the visible preview before trying to reparse cloud file refs");
assert.match(
  appSource.match(/function renderParserRulesCard[\s\S]*?function collectUnknownShiftIssues/)?.[0] || "",
  /parserRuleSuggestions\.length \? `[\s\S]*<strong>User suggestions<\/strong>/,
  "empty user-suggestion sections should be omitted",
);
assert.match(appSource, /function exportHospitalOptions/, "one-off exports should expose hospital options");
assert.match(appSource, /Recognised hospitals &amp; default locations/, "account modal should expose recognised hospital locations");
assert.match(appSource, /data-account-location-key/, "account modal locations should bind to shared settings keys");
assert.match(appSource, /data-admin-user-real-name/, "creator Current users editing should expose the account holder's name");
assert.match(appSource, /async function saveAdminUserName[\s\S]*action: "updateAccount"[\s\S]*targetEmail/, "creator name edits should persist to the selected account");
assert.match(appSource, /data-facility-overview-back-to-creator/, "the non-clinical dashboard header should expose Back to creator during impersonation");
assert.match(appSource, /ACCOUNT_HOSPITAL_LOCATION_ORDER = \["mmc", "ddh", "mch", "casey"\]/, "account modal should keep hospital locations in the expected vertical order");
assert.match(d1CalendarSource, /CREATE TABLE IF NOT EXISTS account_hospital_locations/, "D1 should store account hospital locations relationally");
assert.match(d1CalendarSource, /function applyAccountHospitalLocations/, "SQL-first roster reads should apply account hospital defaults");
assert.match(
  appSource.match(/function buildLocationOptionMarkup[\s\S]*?function detectLocationPreset/)?.[0] || "",
  /locationOptionSourceTypes\(source\)/,
  "location options should consider the roster event source, not only detected imports",
);
assert.match(
  appSource.match(/async function buildDerivedCalendarFilePayload[\s\S]*?function assertDerivedCalendarFilePayload/)?.[0] || "",
  /includeLocations: true/,
  "D1 roster materialisation should retain canonical onsite locations for rebuilds",
);
assert.match(appSource, /skipStatus: true/, "rebuild saves should avoid full aggregate status checks after every file");
assert.match(appSource, /mergeLightweightRosterStatus/, "rebuild saves should merge lightweight per-file status");
assert.match(appSource, /function matchesExportHospitals/, "one-off exports should support hospital filtering");
assert.match(appSource, /function canCopySubscriptionUrl/, "subscription URL availability should be separate from one-off exports");
assert.match(
  appSource.match(/async function handleExportAction[\s\S]*?function downloadIcs/)?.[0] || "",
  /await navigator\.clipboard\.writeText\(url\)[\s\S]*saveCloudState\(snapshot\)\.catch/,
  "subscription URLs should be copied before async feed persistence runs",
);
assert.match(
  appSource.match(/async function handleExportAction[\s\S]*?function downloadIcs/)?.[0] || "",
  /subscriptionUrl\("webcal", exportConfig\.mode === "range" \? "range" : "full"\)/,
  "Apple Calendar action should open the subscription URL via webcal",
);
assert.match(
  appSource.match(/function renderExportModal[\s\S]*?async function handleExportAction/)?.[0] || "",
  /data-export-action="apple">Open in Apple Calendar/,
  "Apple Calendar one-off import should not be disabled by subscription availability",
);
assert.doesNotMatch(
  appSource.match(/async function reportAccountError[\s\S]*?async function updateAccountDetails/)?.[0] || "",
  /currentUserRole === "creator" && !adminViewingEmail/,
  "creator-owned unresolved shift codes should be reportable into the admin issue flow",
);
assert.doesNotMatch(
  appSource.match(/async function buildDerivedCalendarFilePayload[\s\S]*?function assertDerivedCalendarFilePayload/)?.[0] || "",
  /rawFile:/,
  "derived-file saves should not carry retained raw file bytes",
);
assert.match(
  d1CalendarSource,
  /queryRosterFileDoctorsForKeys/,
  "D1 helpers should expose a narrow selected-doctor file lookup",
);
assert.match(appSource, /function rosterSyncSummary/, "system card should expose aggregate roster sync progress");
assert.doesNotMatch(appSource, /function scheduleFailedRosterRetry/, "failed roster syncs should not schedule automatic retry storms");
assert.match(appSource, /data-reparse-import/, "file cards should expose a visible reparse action");
assert.match(appSource, /activeManualReparseIds[\s\S]*is-processing[\s\S]*activeAutomatedSourceRefreshIds[\s\S]*waitForAutomatedRosterSourceRefresh/, "manual and auto-sync reparse controls should retain an in-progress state until their work completes");
assert.match(styleSource, /\.file-reparse\.is-processing[\s\S]*roster-refresh-spin[\s\S]*prefers-reduced-motion/, "an in-progress reparse control should rotate while respecting reduced-motion preferences");
assert.match(appSource, /Reparse produced 0 events/, "zero-event reparses should remain visibly failed");
assert.match(
  appSource.match(/function formatTimestamp\([\s\S]*?function formatIssueHeading/)?.[0] || "",
  /timeZone: "Australia\/Melbourne"[\s\S]*hour: "numeric"[\s\S]*hour12: true[\s\S]*formatToParts/,
  "timestamps should use Australian 12-hour formatting without a leading zero",
);
assert.match(
  appSource.match(/accountsBody\.addEventListener\(\"click\"[\s\S]*?\n\}\);/)?.[0] || "",
  /data-reparse-import[\s\S]*reparseRosterFile[\s\S]*data-refresh-automated-source[\s\S]*refreshAutomatedRosterSource/,
  "admin file controls should invoke their respective manual and auto-sync reparse paths",
);
assert.doesNotMatch(
  appSource.match(/async function renderWhoInsight[\s\S]*?async function renderWhenInsight/)?.[0] || "",
  /ensureInsightRosterAnalysis/,
  "who insights should not fall back to reparsing roster files",
);
assert.doesNotMatch(
  appSource.match(/async function renderInlineWhoInsight[\s\S]*?function renderInlineWhoGroups/)?.[0] || "",
  /ensureInsightRosterAnalysis/,
  "inline who insights should not fall back to reparsing roster files",
);
assert.doesNotMatch(
  appSource.match(/async function renderWhenInsight[\s\S]*?function renderWhenInsightResult/)?.[0] || "",
  /ensureInsightRosterAnalysis/,
  "when insights should not fall back to reparsing roster files",
);
assert.doesNotMatch(
  appSource.match(/function isRosterShiftEvent[\s\S]*?function chooseNextOverlapDate/)?.[0] || "",
  /clinical support|\\bcso\?\\b/i,
  "clinical support shifts should count as rostered work for insight lookups",
);
assert.doesNotMatch(
  appSource.match(/function renderWhenInsightResult[\s\S]*?function comparisonDoctorOptions/)?.[0] || "",
  /doctor\.key === selectedKey/,
  "when insight rendering should not reference an out-of-scope selectedKey",
);
assert.match(
  appSource.match(/async function renderWhenInsight[\s\S]*?function renderWhenInsightResult/)?.[0] || "",
  /fetchRosterOverlapDoctors[\s\S]*const options = prioritizeDoctorOptions[\s\S]*doctorKeys: \[selectedKey\][\s\S]*renderWhenInsightResult\(\{ options, selectedComparison/,
  "general when insights should keep all compact overlap doctors in the dropdown while loading one selected doctor's events",
);
assert.match(
  appSource,
  /function isClinicalSupportEvent\(event\)[\s\S]*function filterWhenInsightEvents\(events, includeCs = false\)/,
  "when insights should detect clinical support separately from roster shift filtering",
);
assert.match(
  appSource.match(/async function openWhenInsight[\s\S]*?async function openWhenInsightForDoctor/)?.[0] || "",
  /includeCs: false/,
  "when insights modal should default Include CS to off",
);
assert.match(
  appSource.match(/function renderWhenInsightResult[\s\S]*?function comparisonDoctorOptions/)?.[0] || "",
  /data-insights-when-include-cs/,
  "when insights modal should render an Include CS toggle",
);
assert.match(
  appSource.match(/async function renderWhenInsight[\s\S]*?function renderWhenInsightResult/)?.[0] || "",
  /filterWhenInsightEvents/,
  "when insights should filter clinical support overlaps via filterWhenInsightEvents",
);
assert.match(
  appSource.match(/async function renderInlineWhenInsight[\s\S]*?function renderInlineWhenInsightResult/)?.[0] || "",
  /filterWhenInsightEvents\([^,]+, false\)/,
  "inline when insights should exclude clinical support overlaps by default",
);
assert.match(
  calendarMigrationSource,
  /CREATE INDEX IF NOT EXISTS idx_roster_events_source_range ON roster_events \(source_type, start_date, end_date\);/,
  "calendar migration should include the source/date range index used by insight lookups",
);
assert.match(
  insightIndexMigrationSource,
  /CREATE INDEX IF NOT EXISTS idx_roster_events_source_range ON roster_events \(source_type, start_date, end_date\);/,
  "existing databases should receive the insight query index through a forward migration",
);
assert.match(
  d1CalendarSource,
  /FROM roster_daily_presence[\s\S]*INNER JOIN roster_events AS ev ON ev\.id = p\.event_id/,
  "coworker lookup should read from daily presence materialization and join roster events for details",
);
const caseyFormData = new FormData();
caseyFormData.append("rosterFiles", new File([caseyBytes], "Casey_Term_2_2026_DRAFT.xlsm", { type: "application/vnd.ms-excel.sheet.macroEnabled.12" }));
const parsedCaseyUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: caseyFormData }));
assert.equal(parsedCaseyUpload.sources.casey.length, 1);
const mchFormData = new FormData();
mchFormData.append("rosterFiles", new File([mchBytes], "Paeds_Term_2_2026.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
const parsedMchUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: mchFormData }));
assert.equal(parsedMchUpload.sources.mch.length, 1);

const mmcUpload = new FormData();
mmcUpload.append("rosterFiles", workbookFile(mmcWorkbook, "AdultTerm1.2026.xlsx"));
const parsedMmcUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: mmcUpload }));
assert.equal(parsedMmcUpload.sources.mmc.length, 1);

const shiftedMmcHeaderWorkbook = XLSX.utils.book_new();
const shiftedMmcHeaderRows = Array.from({ length: 8 }, () => []);
shiftedMmcHeaderRows[2] = ["Role", "Pager No (uploaded)", "", "Cost Centre", "Name (Not Used)", "Emp No", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
shiftedMmcHeaderRows[3] = ["Date", "", "", "", "", "", new Date("2026-08-10T00:00:00Z"), new Date("2026-08-11T00:00:00Z"), new Date("2026-08-12T00:00:00Z"), new Date("2026-08-13T00:00:00Z"), new Date("2026-08-14T00:00:00Z"), new Date("2026-08-15T00:00:00Z"), new Date("2026-08-16T00:00:00Z")];
shiftedMmcHeaderRows[5] = ["", "", "", "SMS", "SMS"];
shiftedMmcHeaderRows[6] = ["", "", "", "SHIFTED SMS", "Shifted COLUMNS", "", "0800-1730 CS", "1430-0000 PGC", "0800-1730 D C", "0800-1730 CSM", "0800-1730 CS OS", "0800-1730 CS Exam"];
XLSX.utils.book_append_sheet(shiftedMmcHeaderWorkbook, XLSX.utils.aoa_to_sheet([[]]), "Whole thing");
XLSX.utils.book_append_sheet(shiftedMmcHeaderWorkbook, XLSX.utils.aoa_to_sheet(shiftedMmcHeaderRows, { cellDates: true }), "Week 2");
const shiftedMmcUpload = new FormData();
shiftedMmcUpload.append("rosterFiles", workbookFile(shiftedMmcHeaderWorkbook, "AdultTerm3.2026.xlsx"));
const parsedShiftedMmcUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: shiftedMmcUpload }));
const shiftedMmcDoctor = doctorOptions(parsedShiftedMmcUpload.sources.mmc, []).find((doctor) => doctor.displayName === "Shifted COLUMNS");
assert.ok(shiftedMmcDoctor, "MMC doctor extraction should follow a shifted Name header");
const shiftedMmcEvents = buildRosterView(parsedShiftedMmcUpload.sources.mmc, [], shiftedMmcDoctor.key).events;
assert.deepEqual(
  shiftedMmcEvents.map((event) => event.start.slice(0, 10)),
  ["2026-08-10", "2026-08-11", "2026-08-13", "2026-08-14", "2026-08-15"],
  "MMC parsing should follow shifted weekday headers, recognise confirmed CS variants, and hide a Dandenong allocation annotation",
);
assert.ok(shiftedMmcEvents.every((event) => event.seniority === "SMS"), "MMC parsing should follow the shifted Cost Centre seniority marker");
assert.ok(shiftedMmcEvents.some((event) => event.title === "MMC: CSM" && event.rawValue === "0800-1730 CSM"), "CSM should be a recognised MMC shift");
assert.ok(shiftedMmcEvents.some((event) => event.title === "MMC: CS OS" && event.rawValue === "0800-1730 CS OS"), "CS OS should be a recognised MMC shift");
assert.ok(shiftedMmcEvents.some((event) => event.title === "MMC: CS Exam" && event.rawValue === "0800-1730 CS Exam" && event.allDay), "CS Exam should be a recognised all-day MMC event");

const mmcInternWorkbook = XLSX.utils.book_new();
const mmcInternRows = Array.from({ length: 10 }, () => []);
mmcInternRows[2] = ["Role", "Pager No", "", "Cost Centre", "Name (Not Used)", "Emp No", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
mmcInternRows[3] = ["Date", "", "", "", "", "", new Date("2026-08-10T00:00:00Z"), new Date("2026-08-11T00:00:00Z"), new Date("2026-08-12T00:00:00Z"), new Date("2026-08-13T00:00:00Z"), new Date("2026-08-14T00:00:00Z"), new Date("2026-08-15T00:00:00Z"), new Date("2026-08-16T00:00:00Z")];
mmcInternRows[5] = ["", "", "", "INTERN"];
mmcInternRows[6] = ["", "", "", "", "Aaron Mahoney", "703292", "0900-1300", "1230-1800 SWA", "1430-0000 SWA", "1430-0000", "1430-0000", "", ""];
mmcInternRows[7] = ["", "", "", "LOCUM"];
mmcInternRows[8] = ["", "", "", "", "Must Not Parse", "", "0800-1730", "", "", "", "", "", ""];
XLSX.utils.book_append_sheet(mmcInternWorkbook, XLSX.utils.aoa_to_sheet([[]]), "Whole thing");
XLSX.utils.book_append_sheet(mmcInternWorkbook, XLSX.utils.aoa_to_sheet(mmcInternRows, { cellDates: true }), "Week 2");
const mmcInternUpload = new FormData();
mmcInternUpload.append("rosterFiles", workbookFile(mmcInternWorkbook, "AdultTerm3.2026.xlsx"));
const parsedMmcInternUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: mmcInternUpload }));
const aaronIntern = doctorOptions(parsedMmcInternUpload.sources.mmc, []).find((doctor) => doctor.key === "AARON MAHONEY");
assert.ok(aaronIntern, "MMC parsing should retain clinicians below the INTERN section heading");
assert.equal(doctorOptions(parsedMmcInternUpload.sources.mmc, []).some((doctor) => doctor.key === "MUST NOT PARSE"), false, "MMC parsing should still stop at the LOCUM section");
const aaronInternEvents = buildRosterView(parsedMmcInternUpload.sources.mmc, [], aaronIntern.key).events;
assert.equal(aaronInternEvents.length, 5, "MMC Intern rows should produce every rostered shift");
assert.ok(aaronInternEvents.every((event) => event.seniority === "Intern"), "MMC Intern rows should retain Intern seniority");
assert.ok(aaronInternEvents.some((event) => event.title === "MMC: Swing AM" && event.start.includes("12:30:00")), "MMC Intern swing allocations should use the normal MMC shift rules");

await assertRejectsMixedTermUpload(
  "MMC date typo should identify the worksheet and cell",
  withWorkbookDate(mmcWorkbook, "Week 1", "H4", new Date("2025-02-17T00:00:00")),
  "AdultTerm1.2026 typo.xlsx",
  ["AdultTerm1.2026 typo.xlsx has dates from multiple terms", "Week 1 cell H4", "2025-02-17", "Term 1 2025"],
);

await assertRejectsMixedTermUpload(
  "Casey date typo should identify the worksheet and cell",
  withWorkbookCell(caseyWorkbook, "May 4", "B2", { t: "s", v: "01-Feb", w: "01-Feb" }),
  "Casey_Term_2_2026 typo.xlsm",
  ["Casey_Term_2_2026 typo.xlsm has dates from multiple terms", "May 4 cell B2", "2026-02-01", "Term 4 2025"],
);

await assertRejectsMixedTermUpload(
  "MCH date typo should identify the worksheet and cell",
  withWorkbookDate(mchWorkbook, "Week 1", "F19", new Date("2025-05-04T00:00:00")),
  "Paeds_Term_2_2026 typo.xlsx",
  ["Paeds_Term_2_2026 typo.xlsx has dates from multiple terms", "Week 1 cell F19", "2025-05-04", "Term 1 2025"],
);

const ddhSheetName = ddhWorkbook.SheetNames[0];
await assertRejectsMixedTermUpload(
  "DDH date typo should identify the worksheet and cell",
  withWorkbookCell(ddhWorkbook, ddhSheetName, "B1", { t: "s", v: "Mon. Feb. 3, 2025", w: "Mon. Feb. 3, 2025" }),
  "Dandenong typo.xlsx",
  ["Dandenong typo.xlsx has dates from multiple terms", `${ddhSheetName} cell B1`, "2025-02-03", "Term 1 2025"],
);

const mixedUpload = new FormData();
mixedUpload.append("rosterFiles", workbookFile(mmcWorkbook, "AdultTerm1.2026.xlsx"));
mixedUpload.append("rosterFiles", workbookFile(withWorkbookDate(mmcWorkbook, "Week 1", "H4", new Date("2025-02-17T00:00:00")), "OnlyThisFileIsBad.xlsx"));
await assert.rejects(
  () => parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: mixedUpload })),
  (error) => error.message.includes("OnlyThisFileIsBad.xlsx has dates from multiple terms") && !error.message.includes("AdultTerm1.2026.xlsx has dates from multiple terms"),
);

const doctors = doctorOptions(mmcWorkbook, ddhWorkbook, caseyWorkbook);
const defaultRules = parserRuleDefaults();
const mmcRules = defaultRules.mmc || [];
const mchRules = defaultRules.mch || [];
const ddhRules = defaultRules.ddh || [];
const hasMmcRule = (seniority, code) => mmcRules.some((rule) => rule.seniority === seniority && rule.code === code);
const hasMchRule = (seniority, code, base = "") => mchRules.some((rule) => rule.seniority === seniority && rule.code === code && (!base || rule.base === base));
assert.ok(hasMmcRule("SMS", "AGC"));
assert.ok(hasMmcRule("CMO", "AGC"));
assert.ok(hasMmcRule("SMS", "CS"));
assert.ok(hasMmcRule("CMO", "CSO"));
assert.ok(hasMmcRule("Senior Registrar", "AGC"), "Senior Registrars can be rostered acting-up consultant allocations");
assert.equal(hasMmcRule("HMO", "AGC"), false);
assert.equal(hasMmcRule("Senior Registrar", "CS"), false, "acting-up consultant rules should not broaden generic Clinical Support allocations");
assert.equal(hasMmcRule("HMO", "CSO"), false);
assert.equal(hasMmcRule("SMS", "ACR"), false);
assert.equal(hasMmcRule("SMS", "ARR"), false);
assert.equal(hasMmcRule("SMS", "ASSR"), false);
assert.ok(hasMmcRule("Senior Registrar", "SWA"));
assert.ok(hasMmcRule("Transitional/Intermediate Registrar", "SWP"));
assert.ok(hasMmcRule("Junior Registrar", "AHJ"));
assert.ok(hasMmcRule("HMO", "PHJ"));
assert.ok(hasMmcRule("Intern", "NSSJ"));
for (const seniority of ["Senior Registrar", "Transitional/Intermediate Registrar", "Junior Registrar", "HMO", "Intern"]) {
  assert.ok(hasMmcRule(seniority, "ASSJ"), `default MMC rules should include ASSJ for ${seniority}`);
  assert.ok(hasMmcRule(seniority, "PSSJ"), `default MMC rules should include PSSJ for ${seniority}`);
}
assert.ok(hasMchRule("SMS", "CS", "CS"));
assert.ok(hasMchRule("CMO", "OCS", "CS Office"));
assert.ok(hasMchRule("HMO", "PHNW", "PHNW"));
assert.ok(ddhRules.some((rule) => rule.code === "ROVER AM" && rule.base === "Rover" && rule.period === "AM"));
assert.ok(ddhRules.some((rule) => rule.code === "ROVER PM" && rule.base === "Rover" && rule.period === "PM"));
const nssjRule = mmcRules.find((rule) => rule.seniority === "HMO" && rule.code === "NSSJ");
assert.equal(nssjRule.startTime, "23:00");

const unmappedTimedMmcWorkbook = XLSX.utils.book_new();
const unmappedTimedMmcSheet = XLSX.utils.aoa_to_sheet([
  [],
  [],
  [],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "SENIOR REG"],
  ["", "", "", "Patrick TAN", "", "0800-1730 ASSJ"],
]);
for (let index = 0; index < 7; index += 1) {
  unmappedTimedMmcSheet[XLSX.utils.encode_cell({ r: 3, c: 5 + index })] = { t: "d", v: new Date(`2026-05-${String(4 + index).padStart(2, "0")}T00:00:00`) };
}
XLSX.utils.book_append_sheet(unmappedTimedMmcWorkbook, XLSX.utils.aoa_to_sheet([[]]), "Whole thing");
XLSX.utils.book_append_sheet(unmappedTimedMmcWorkbook, unmappedTimedMmcSheet, "Week 1");
const unmappedTimedMmcView = buildRosterView([{ id: "unmapped-assj", workbook: unmappedTimedMmcWorkbook, file: { name: "AdultTerm.xlsx", size: 1, lastModified: 1 } }], [], "PATRICK TAN");
assert.ok(
  unmappedTimedMmcView.events.some((event) => event.rawValue === "0800-1730 ASSJ" && event.title === "MMC: SSU AM" && event.start.includes("08:00:00") && event.end.includes("17:30:00")),
  "default ASSJ rules should resolve Senior Registrar SSU AM while preserving explicit roster time",
);
assert.equal(unmappedTimedMmcView.issues.some((issue) => issue.rawValue === "0800-1730 ASSJ"), false, "default ASSJ rules should not enter the unresolved shift-code workflow");
assert.equal(nssjRule.endTime, "08:30");
for (const sourceRules of [defaultRules.ddh, defaultRules.casey]) {
  assert.equal(sourceRules.some((rule) => rule.seniority === "Senior Registrar" && rule.base === "CS"), false);
  assert.equal(sourceRules.some((rule) => rule.seniority === "HMO" && rule.base === "CS"), false);
  assert.ok(sourceRules.some((rule) => rule.seniority === "SMS" && rule.base === "CS"));
  assert.ok(sourceRules.some((rule) => rule.seniority === "CMO" && rule.base === "CS"));
}
assert.ok(doctors.length > 100);
const richard = doctors.find((doctor) => doctor.displayName === "Richard HAYDON");
assert.ok(richard);
assert.deepEqual(richard.sourceTypes, ["mmc", "ddh"]);
assert.ok(doctors.find((doctor) => doctor.displayName === "Brianna Dawn MURPHY"));
assert.ok(doctors.find((doctor) => doctor.displayName === "Patrick TAN"));
assert.equal(doctors.find((doctor) => doctor.displayName === "Aarushi Pathania"), undefined);
assert.equal(doctors.find((doctor) => doctor.displayName === "HMO MUST BE"), undefined);

const ananthMmc = doctors.find((doctor) => doctor.displayName === "Ananth SUNDARALINGAM");
assert.ok(ananthMmc);
const ananthMmcSabbatical = buildRosterView(mmcWorkbook, [], ananthMmc.key).events.find((event) => event.rawValue === "Sabbatical leave");
assert.ok(ananthMmcSabbatical);
assert.equal(ananthMmcSabbatical.title, "Sabbatical");
assert.equal(ananthMmcSabbatical.end, "2026-02-16", "MMC weekly sabbatical markers should cover the full roster week");
const harmeenKaur = doctors.find((doctor) => doctor.displayName === "Harmeen KAUR");
assert.ok(buildRosterView(mmcWorkbook, [], harmeenKaur.key).events.some((event) => event.rawValue === "SL PM" && event.title === "Sabbatical"), "SL without a slash must be sabbatical leave");
const scottJosey = doctors.find((doctor) => doctor.displayName === "Scott JOSEY");
const scottLeaveEvents = buildRosterView(mmcWorkbook, [], scottJosey.key).events.filter((event) => event.title === "Annual Leave");
assert.ok(scottLeaveEvents.some((event) => event.rawValue.includes("AL 9.5hrs")));
assert.ok(scottLeaveEvents.some((event) => event.rawValue.includes("Annual leave 19hrs")));
assert.ok(
  buildRosterView(mmcWorkbook, [], scottJosey.key).events.some((event) => event.rawValue === "1430-0000 PCC" && event.start.startsWith("2026-04-17T14:30:00")),
  "a weekly leave marker must not suppress Scott's Friday PCC shift",
);
const ericaChan = doctors.find((doctor) => doctor.displayName === "Erica CHAN");
assert.ok(buildRosterView(mmcWorkbook, [], ericaChan.key).events.some((event) => event.rawValue === "CME leave - 38hrs" && event.title === "Conference Leave"));

const michaelMerged = doctorOptions(mmcWorkbook, [], [], mchWorkbook).filter((doctor) => doctor.displayName.toUpperCase().includes("MICHAEL COMAN"));
assert.equal(michaelMerged.length, 1);
assert.deepEqual(michaelMerged[0].sourceTypes, ["mmc", "mch"]);
assert.ok(michaelMerged[0].aliases.some((alias) => alias.sourceType === "mmc" && alias.key === "MICHAEL COMAN"));
assert.ok(michaelMerged[0].aliases.some((alias) => alias.sourceType === "mch" && alias.key === "DR MICHAEL COMAN"));
const michaelMergedView = buildRosterView(mmcWorkbook, [], michaelMerged[0].key, undefined, {}, {}, michaelMerged[0].aliases, [], mchWorkbook);
assert.ok(michaelMergedView.events.some((event) => event.source === "MMC"));
assert.ok(michaelMergedView.events.some((event) => event.source === "MCH"));
assert.ok(michaelMergedView.events.some((event) => event.title === "Conference Leave"));

const markDouglas = doctors.find((doctor) => doctor.displayName === "Mark DOUGLAS");
assert.ok(markDouglas);
const markView = buildRosterView(mmcWorkbook, [], markDouglas.key);
assert.ok(markView.events.some((event) => event.title === "MMC: AM shift"));
assert.ok(markView.events.some((event) => event.title === "MMC: PM shift"));

const deslinAraullo = doctors.find((doctor) => doctor.displayName === "Deslin ARAULLO");
assert.ok(deslinAraullo);
const deslinView = buildRosterView(mmcWorkbook, [], deslinAraullo.key);
assert.ok(deslinView.events.some((event) => event.title === "MMC: Hub PM"));
assert.ok(deslinView.events.some((event) => event.title === "MMC: Swing AM"));

const caseyDoctors = doctorOptions([], [], caseyWorkbook);
assert.ok(caseyDoctors.length > 130);
const defaultParserRules = parserRuleDefaults();
setParserExtensions({
  ...defaultParserRules,
  casey: [
    ...defaultParserRules.casey.filter((rule) => rule.code !== "AM MIC"),
    {
      source: "Casey",
      seniority: "SMS",
      code: "AM MIC",
      kind: "shift",
      base: "MIC",
      period: "AM",
      suffix: "",
      allDay: true,
      startTime: "",
      endTime: "",
      location: "",
      includeAsShift: true,
    },
  ],
});
const aliAsadpourCasey = caseyDoctors.find((doctor) => doctor.displayName === "Ali ASADPOUR");
assert.ok(aliAsadpourCasey, "fixture should include Ali ASADPOUR");
const aliAsadpourCaseyView = buildRosterView([], [], aliAsadpourCasey.key, undefined, {}, {}, [], caseyWorkbook);
assert.ok(
  aliAsadpourCaseyView.events.some((event) => event.source === "Casey" && event.title === "Casey: MIC AM" && event.allDay === true),
  "all-day Casey parser rules should render MIC shifts without crashing import",
);
setParserExtensions(defaultParserRules);
assert.ok(caseyDoctors.find((doctor) => doctor.displayName === "Andrew DYALL"));
assert.ok(caseyDoctors.find((doctor) => doctor.displayName === "Dennis CHUNG"));
assert.ok(caseyDoctors.find((doctor) => doctor.displayName === "Rizwana SADAF"));
assert.ok(caseyDoctors.find((doctor) => doctor.displayName === "Victor Ki Chung LI"));
assert.equal(caseyDoctors.find((doctor) => doctor.displayName === "Rostered staff"), undefined);

const patrickTan = doctorOptions(mmcWorkbook, [], caseyWorkbook).find((doctor) => doctor.displayName === "Patrick TAN");
assert.ok(patrickTan);
assert.deepEqual(patrickTan.sourceTypes, ["mmc", "casey"]);
const patrickTanView = buildRosterView(mmcWorkbook, [], patrickTan.key, undefined, {}, {}, [], caseyWorkbook);
const patrickCaseyEvents = patrickTanView.events.filter((event) => event.source === "Casey");
assert.ok(patrickCaseyEvents.length > 40);
assert.equal(patrickCaseyEvents.some((event) => event.start.startsWith("2025")), false);
assert.ok(patrickCaseyEvents.some((event) => event.title === "Casey: MIC PM"));
assert.ok(patrickCaseyEvents.some((event) => event.title === "Casey: Orientation" && event.rawValue === "Orient 0800-1730" && event.start.includes("08:00:00") && event.end.includes("17:30:00")));
const patrickMergedLeave = patrickCaseyEvents.find((event) => event.title === "Annual Leave" && event.start === "2026-07-27");
assert.ok(patrickMergedLeave);
assert.equal(patrickMergedLeave.end, "2026-08-03");
assert.equal(patrickMergedLeave.rawValue, "Annual Leave");

const suzanFoxCasey = caseyDoctors.find((doctor) => doctor.displayName === "Suzan FOX");
assert.ok(suzanFoxCasey);
const suzanFoxCaseyView = buildRosterView([], [], suzanFoxCasey.key, undefined, {}, {}, suzanFoxCasey.aliases, caseyWorkbook);
const suzanMergedLeave = suzanFoxCaseyView.events.find((event) => event.title === "Annual Leave" && event.start === "2026-07-27");
assert.ok(suzanMergedLeave);
assert.equal(suzanMergedLeave.end, "2026-08-03");
assert.equal(suzanMergedLeave.rawValue, "Annual Leave");

const derivedLeavePreview = buildPreviewFromDerivedEvents([
  {
    id: "leave-a",
    title: "MMC: Annual Leave",
    rawValue: "Annual Leave",
    source: "MMC",
    sources: ["MMC"],
    start: "2026-07-27",
    end: "2026-08-03",
    allDay: true,
  },
  {
    id: "leave-b",
    title: "Casey: Conference Leave",
    rawValue: "Conference Leave",
    source: "Casey",
    sources: ["Casey"],
    start: "2026-07-27",
    end: "2026-08-03",
    allDay: true,
  },
  {
    id: "shift-a",
    title: "MMC: AM",
    rawValue: "AM",
    source: "MMC",
    start: "2026-07-28T08:00:00",
    end: "2026-07-28T17:00:00",
    allDay: false,
  },
]);
const derivedLeaveEvents = derivedLeavePreview.events.filter((event) => /leave/i.test(event.title));
assert.equal(derivedLeaveEvents.length, 1, "overlapping leave from multiple sources should render once");
assert.deepEqual(derivedLeaveEvents[0].sources, ["MMC", "Casey"]);
assert.equal(derivedLeavePreview.events.some((event) => event.id === "shift-a"), true, "non-leave shifts must remain visible");
const adjacentMixedLeavePreview = buildPreviewFromDerivedEvents([
  { id: "conference-week", title: "Conference Leave", rawValue: "Conference Leave", source: "MMC", start: "2026-02-02", end: "2026-02-09", allDay: true },
  { id: "annual-week", title: "Annual Leave", rawValue: "Annual Leave", source: "MMC", start: "2026-02-09", end: "2026-02-16", allDay: true },
]);
assert.deepEqual(
  adjacentMixedLeavePreview.events.map((event) => [event.title, event.start, event.end]),
  [["Conference Leave", "2026-02-02", "2026-02-09"], ["Annual Leave", "2026-02-09", "2026-02-16"]],
  "adjacent different leave types should not be merged",
);

const andrewDyallCasey = caseyDoctors.find((doctor) => doctor.displayName === "Andrew DYALL");
const andrewCaseyView = buildRosterView([], [], andrewDyallCasey.key, undefined, {}, {}, [], caseyWorkbook);
assert.ok(andrewCaseyView.events.some((event) => event.title === "Casey: TL AM"));
assert.ok(andrewCaseyView.events.some((event) => event.title === "Casey: UFD PM"));
assert.ok(andrewCaseyView.events.some((event) => event.title === "Casey: MIC AM"));
assert.equal(andrewCaseyView.events.some((event) => /PAEDS/i.test(event.title)), false, "Paediatrics references outside the MCH roster are not local Casey shifts");
assert.ok(andrewCaseyView.events.some((event) => event.title === "Casey: CS" && event.start.includes("08:00:00") && event.end.includes("17:30:00")));

const bashirCasey = caseyDoctors.find((doctor) => doctor.displayName === "Bashir GONDAL");
const bashirCaseyView = buildRosterView([], [], bashirCasey.key, undefined, {}, {}, [], caseyWorkbook);
assert.ok(bashirCaseyView.events.some((event) => event.title === "Casey: SSU AM"));

const dennisCasey = caseyDoctors.find((doctor) => doctor.displayName === "Dennis CHUNG");
const dennisCaseyView = buildRosterView([], [], dennisCasey.key, undefined, {}, {}, [], caseyWorkbook);
assert.ok(dennisCaseyView.events.some((event) => event.title === "Casey: Night shift" && event.start.includes("23:00:00") && event.end.startsWith("2026-05-06")));
assert.equal(dennisCaseyView.events.filter((event) => event.title === "Casey: Night shift" && event.start.startsWith("2026-05-05") && event.end.startsWith("2026-05-06")).length, 1);

const jasonAwCasey = caseyDoctors.find((doctor) => doctor.displayName === "Jason AW");
const jasonAwCaseyView = buildRosterView([], [], jasonAwCasey.key, undefined, {}, {}, [], caseyWorkbook);
assert.ok(jasonAwCaseyView.events.some((event) => event.title === "Annual Leave"));

const mustafaCasey = caseyDoctors.find((doctor) => doctor.displayName === "Mustafa Al ASAAD");
const mustafaCaseyView = buildRosterView([], [], mustafaCasey.key, undefined, {}, {}, [], caseyWorkbook);
assert.ok(mustafaCaseyView.events.some((event) => event.title === "Conference Leave"));

const mchDoctors = doctorOptions([], [], [], mchWorkbook);
assert.ok(mchDoctors.length >= 60);
assert.ok(mchDoctors.find((doctor) => doctor.displayName === "Adam WEST"));
assert.ok(mchDoctors.find((doctor) => doctor.displayName === "Mark LIM"));
assert.ok(mchDoctors.find((doctor) => doctor.displayName === "Firas HAMDAN"));
assert.ok(mchDoctors.find((doctor) => doctor.displayName === "Peter AHN"));
assert.equal(mchDoctors.find((doctor) => doctor.displayName === "ONCALL 0000-0800"), undefined);
assert.equal(mchDoctors.find((doctor) => doctor.displayName === "requested off"), undefined);

const adamWestMch = mchDoctors.find((doctor) => doctor.displayName === "Adam WEST");
const adamWestMchView = buildRosterView([], [], adamWestMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(adamWestMchView.events.some((event) => event.title === "MCH: CS" && event.rawValue === "0800-1730 CS"));
assert.ok(adamWestMchView.events.some((event) => event.title === "MCH: PM shift" && event.rawValue === "1430-0000" && event.end.startsWith("2026-05-09")));

const bobSeithMch = mchDoctors.find((doctor) => doctor.displayName === "Bob SEITH");
const bobSeithMchView = buildRosterView([], [], bobSeithMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(bobSeithMchView.events.some((event) => event.title === "MCH: DEMT" && event.rawValue === "0800-1730 DEMT"));
assert.ok(bobSeithMchView.events.some((event) => event.title === "MCH: CS" && event.rawValue === "0800-1730CS"));

const andrewHardyMch = mchDoctors.find((doctor) => doctor.displayName === "Andrew HARDY");
const andrewHardyMchView = buildRosterView([], [], andrewHardyMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(andrewHardyMchView.events.some((event) => event.title === "MCH: CS Office" && event.rawValue === "0800-1730 OCS"));
assert.ok(andrewHardyMchView.events.some((event) => event.title === "Conference Leave" && event.rawValue === "CME/L" && event.allDay));
assert.ok(andrewHardyMchView.events.some((event) => event.title === "Exam Leave" && event.rawValue === "ME/L" && event.allDay));
assert.ok(andrewHardyMchView.events.some((event) => event.title === "Conference Leave" && event.rawValue === "CME/L" && event.start === "2026-06-08" && event.end === "2026-06-15"));

const adamWestMchWeek6 = adamWestMchView.events.filter((event) => event.rawValue === "PHNW 0800-1730");
assert.ok(adamWestMchWeek6.some((event) => event.title === "MCH: PHNW"));
const noSpacePhnwMchWorkbook = withWorkbookCell(mchWorkbook, "Week 6", "F21", { t: "s", v: "0800-1730PHNW", w: "0800-1730PHNW" });
const noSpacePhnwMchView = buildRosterView([], [], adamWestMch.key, undefined, {}, {}, [], [], noSpacePhnwMchWorkbook);
assert.ok(noSpacePhnwMchView.events.some((event) => event.title === "MCH: PHNW" && event.rawValue === "0800-1730PHNW"));
const bareNightMchWorkbook = withWorkbookCell(mchWorkbook, "Week 6", "F21", { t: "s", v: "NIGHT", w: "NIGHT" });
const bareNightMchView = buildRosterView([], [], adamWestMch.key, undefined, {}, {}, [], [], bareNightMchWorkbook);
assert.ok(bareNightMchView.events.some((event) => event.title === "MCH: Night shift" && event.rawValue === "NIGHT" && event.timeLabel === "23:00-08:30"), "bare NIGHT should be a timed MCH night shift");
const simDayMchWorkbook = withWorkbookCell(mchWorkbook, "Week 6", "G21", { t: "s", v: "(SIM day) 0800-1730", w: "(SIM day) 0800-1730" });
const simDayMchView = buildRosterView([], [], adamWestMch.key, undefined, {}, {}, [], [], simDayMchWorkbook);
assert.ok(simDayMchView.events.some((event) => event.title === "MCH: SIM day" && event.rawValue === "(SIM day) 0800-1730" && event.timeLabel === "08:00-17:30"), "timed SIM day should remain a timed teaching shift");

const markLimMch = mchDoctors.find((doctor) => doctor.displayName === "Mark LIM");
const markLimMchView = buildRosterView([], [], markLimMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(markLimMchView.events.some((event) => event.title === "MCH: Night shift" && event.rawValue === "2300-0830" && event.end.startsWith("2026-05-09")));
assert.equal(markLimMchView.events.filter((event) => event.title === "MCH: Night shift" && event.rawValue === "2300-0830" && event.start.startsWith("2026-05-08") && event.end.startsWith("2026-05-09")).length, 1);
assert.ok(markLimMchView.events.some((event) => event.title === "Conference Leave" && event.rawValue === "C/L" && event.allDay));

const firasMch = mchDoctors.find((doctor) => doctor.displayName === "Firas HAMDAN");
const firasMchView = buildRosterView([], [], firasMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(firasMchView.events.some((event) => event.title === "Annual Leave" && event.rawValue === "AL 0.5" && event.allDay));

const marianPanlilioMch = mchDoctors.find((doctor) => doctor.displayName === "Marian PANLILIO");
const marianPanlilioMchView = buildRosterView([], [], marianPanlilioMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(marianPanlilioMchView.events.some((event) => event.title === "Sick leave" && event.rawValue.trim() === "S/L PM" && event.allDay));

const houshmandMch = mchDoctors.find((doctor) => doctor.displayName === "Houshmand REFAEI");
const houshmandMchView = buildRosterView([], [], houshmandMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.equal(houshmandMchView.events.some((event) => String(event.rawValue || "").includes("EDO")), false);

const overlappingConferenceWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(overlappingConferenceWorkbook, XLSX.utils.aoa_to_sheet([
  ["TERM 2, 2026", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
  ["", "20-Jul", "21-Jul", "22-Jul", "23-Jul", "24-Jul", "25-Jul", "26-Jul"],
  ["Dr Michael Coman", "C/L", "C/L", "C/L", "C/L", "C/L", "C/L", "C/L"],
  ["Daily LEAVE", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave"],
]), "Week 1");
const michaelComan = doctorOptions([], [], overlappingConferenceWorkbook, mchWorkbook).find((doctor) => doctor.displayName === "Michael COMAN");
assert.ok(michaelComan);
const michaelComanView = buildRosterView([], [], michaelComan.key, undefined, {}, {}, michaelComan.aliases, overlappingConferenceWorkbook, mchWorkbook);
const michaelConferenceEvents = michaelComanView.events.filter((event) => event.title === "Conference Leave" && event.start === "2026-07-20");
assert.equal(michaelConferenceEvents.length, 1);
assert.equal(michaelConferenceEvents[0].end, "2026-07-27");
assert.equal(michaelConferenceEvents[0].rawValue, "C/L / CME/L");
assert.ok(michaelComanView.reviewItems.some((item) => item.id === michaelConferenceEvents[0].id));

const dailyLeave = doctorOptions([], [], overlappingConferenceWorkbook).find((doctor) => doctor.displayName === "Daily LEAVE");
assert.ok(dailyLeave);
const dailyLeaveView = buildRosterView([], [], dailyLeave.key, undefined, {}, {}, [], overlappingConferenceWorkbook);
const dailyAnnualLeave = dailyLeaveView.events.filter((event) => event.title === "Annual Leave");
assert.equal(dailyAnnualLeave.length, 1);
assert.equal(dailyAnnualLeave[0].start, "2026-07-20");
assert.equal(dailyAnnualLeave[0].end, "2026-07-27");
assert.ok(dailyLeaveView.reviewItems.some((item) => item.id === dailyAnnualLeave[0].id));

const michaelAnnualWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(michaelAnnualWorkbook, XLSX.utils.aoa_to_sheet([
  ["TERM 2, 2026", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
  ["", "27-Jul", "28-Jul", "29-Jul", "30-Jul", "31-Jul", "1-Aug", "2-Aug"],
  ["Michael COMAN", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave"],
]), "Week 1");
const michaelAnnualOption = doctorOptions([], [], michaelAnnualWorkbook, mchWorkbook).find((doctor) => doctor.displayName === "Michael COMAN");
assert.ok(michaelAnnualOption);
assert.deepEqual(michaelAnnualOption.sourceTypes, ["casey", "mch"]);
const michaelAnnualView = buildRosterView([], [], michaelAnnualOption.key, undefined, {}, {}, michaelAnnualOption.aliases, michaelAnnualWorkbook, mchWorkbook);
const michaelAnnualEvents = michaelAnnualView.events.filter((event) => ["Conference Leave", "Annual Leave"].includes(event.title));
// Distinct leave categories must remain distinct even when their date ranges touch.
assert.equal(michaelAnnualEvents.length, 3);
assert.ok(michaelAnnualEvents.some((event) => event.title === "Conference Leave" && event.start === "2026-07-20" && event.end === "2026-07-27"));
assert.ok(michaelAnnualEvents.some((event) => event.title === "Annual Leave" && event.start === "2026-07-27" && event.end === "2026-08-03"));
assert.ok(michaelAnnualView.events.some((event) => event.source === "MCH"));

const mergedAnnualWorkbook = XLSX.utils.book_new();
const mergedAnnualSheet = XLSX.utils.aoa_to_sheet([
  ["TERM 2, 2026", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
  ["", "27-Jul", "28-Jul", "29-Jul", "30-Jul", "31-Jul", "1-Aug", "2-Aug"],
  ["Merged LEAVE", "Annual Leave", "", "", "", "", "", ""],
]);
mergedAnnualSheet["!merges"] = [{ s: { r: 2, c: 1 }, e: { r: 2, c: 7 } }];
XLSX.utils.book_append_sheet(mergedAnnualWorkbook, mergedAnnualSheet, "Week 1");
const mergedAnnualDoctor = doctorOptions([], [], mergedAnnualWorkbook).find((doctor) => doctor.displayName === "Merged LEAVE");
assert.ok(mergedAnnualDoctor);
const mergedAnnualView = buildRosterView([], [], mergedAnnualDoctor.key, undefined, {}, {}, [], mergedAnnualWorkbook);
const mergedAnnualEvents = mergedAnnualView.events.filter((event) => event.title === "Annual Leave");
assert.equal(mergedAnnualEvents.length, 1);
assert.equal(mergedAnnualEvents[0].start, "2026-07-27");
assert.equal(mergedAnnualEvents[0].end, "2026-08-03");
assert.equal(mergedAnnualEvents[0].rawValue, "Annual Leave");

const annualSynonymWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(annualSynonymWorkbook, XLSX.utils.aoa_to_sheet([
  ["TERM 2, 2026", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
  ["", "3-Aug", "4-Aug", "5-Aug", "6-Aug", "7-Aug", "8-Aug", "9-Aug"],
  ["Paeds Annual", "Paeds AL", "", "", "", "", "", ""],
  ["Casey Annual", "Casey AL", "", "", "", "", "", ""],
  ["Paeds Sick", "Paeds S/L", "", "", "", "", "", ""],
]), "Week 1");
const paedsAnnual = doctorOptions([], [], annualSynonymWorkbook).find((doctor) => doctor.displayName === "Paeds ANNUAL");
const caseyAnnual = doctorOptions([], [], annualSynonymWorkbook).find((doctor) => doctor.displayName === "Casey ANNUAL");
const paedsSick = doctorOptions([], [], annualSynonymWorkbook).find((doctor) => doctor.displayName === "Paeds SICK");
assert.ok(buildRosterView([], [], paedsAnnual.key, undefined, {}, {}, paedsAnnual.aliases, annualSynonymWorkbook).events.some((event) => event.title === "Annual Leave" && event.rawValue === "Paeds AL"));
assert.ok(buildRosterView([], [], caseyAnnual.key, undefined, {}, {}, caseyAnnual.aliases, annualSynonymWorkbook).events.some((event) => event.title === "Annual Leave" && event.rawValue === "Casey AL"));
assert.ok(buildRosterView([], [], paedsSick.key, undefined, {}, {}, paedsSick.aliases, annualSynonymWorkbook).events.some((event) => event.title === "Sick leave" && event.rawValue === "Paeds S/L"));

const view = buildRosterView(mmcWorkbook, ddhWorkbook, richard.key);
const summary = previewSummary(view.events);

const aftabMmc = doctors.find((doctor) => doctor.displayName === "Aftab SAMDANI");
assert.ok(aftabMmc);
const aftabMmcView = buildRosterView(mmcWorkbook, [], aftabMmc.key);
assert.ok(aftabMmcView.events.some((event) => event.title === "Conference Leave" && event.rawValue.toUpperCase() === "CME LEAVE" && event.allDay));

assert.equal(view.events.length, 38);
assert.equal(summary.date_range, "2026-02-02 to 2026-05-02");
assert.ok(view.events.some((event) => event.title === "Conference Leave" && event.rawValue.includes("Dandenong CL")));
assert.ok(view.reviewItems.length >= view.events.length);
assert.ok(view.events.some((event) => event.rawValue.includes("Annual leave")));
assert.ok(view.events.some((event) => event.title === "DDH: Orange PM"));
assert.ok(view.events.some((event) => event.title === "Sick leave"));

const ddhFullWorkbook = XLSX.utils.book_new();
const ddhFullSheet = XLSX.utils.aoa_to_sheet([
  ["", "Mon. Feb. 02, 2026", "Tue. Feb. 03, 2026", "Wed. Feb. 04, 2026", "Thu. Feb. 05, 2026", "Fri. Feb. 06, 2026", "Sat. Feb. 07, 2026", "Sun. Feb. 08, 2026"],
  ["Richard Haydon", "", "", "", "", "", "", ""],
  ["SENIOR MEDICAL STAFF", "", "", "", "", "", "", ""],
  ["Jim BARTON", "AVAO AM", "", "Orange PM (on-call)", "AVAO PM", "Clinical Support", "", ""],
  ["", "07:30-17:00", "", "15:00-00:00", "14:30-00:00", "", "", ""],
  ["Caroline BOLT", "Orange PM (on-call)", "", "AVAO AM", "", "Orange AM IC", "", ""],
  ["", "15:00-00:00", "", "07:30-17:00", "", "08:00-18:00", "", ""],
  ["Di FLOOD", "CS AM", "SSU SMS", "Clinical Support ACEM OSCE", "", "HITH PM", "", ""],
  ["", "", "07:30-17:30", "", "", "", "", ""],
]);
XLSX.utils.book_append_sheet(ddhFullWorkbook, ddhFullSheet, "Sheet1");
const ddhFullDoctors = doctorOptions([], ddhFullWorkbook);
assert.ok(ddhFullDoctors.find((doctor) => doctor.displayName === "Jim BARTON"));
assert.ok(ddhFullDoctors.find((doctor) => doctor.displayName === "Caroline BOLT"));
assert.ok(ddhFullDoctors.find((doctor) => doctor.displayName === "Di FLOOD"));
assert.equal(ddhFullDoctors.find((doctor) => doctor.displayName === "SENIOR MEDICAL STAFF"), undefined);

const jim = ddhFullDoctors.find((doctor) => doctor.displayName === "Jim BARTON");
const jimView = buildRosterView([], ddhFullWorkbook, jim.key);
assert.ok(jimView.events.some((event) => event.title === "DDH: AVAO AM"));
assert.ok(jimView.events.some((event) => event.title === "DDH: Orange PM"));
assert.ok(jimView.events.some((event) => event.title === "DDH: AVAO PM"));
assert.ok(jimView.events.some((event) => event.title === "DDH: CS"));

const diFlood = ddhFullDoctors.find((doctor) => doctor.displayName === "Di FLOOD");
const diFloodView = buildRosterView([], ddhFullWorkbook, diFlood.key);
assert.ok(diFloodView.events.some((event) => event.title === "DDH: CS" && event.rawValue === "CS AM"));
assert.ok(diFloodView.events.some((event) => event.title === "DDH: SSU" && event.start.includes("07:30:00")));
assert.ok(diFloodView.events.some((event) => event.title === "DDH: CS Exam" && event.rawValue === "Clinical Support ACEM OSCE" && event.allDay), "DDH Clinical Support ACEM OSCE should be an all-day CS Exam event");
assert.equal(diFloodView.events.some((event) => event.rawValue === "HITH PM"), false, "other-hospital annotations must not become DDH calendar events");

const ddhDefaultTimesWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(ddhDefaultTimesWorkbook, XLSX.utils.aoa_to_sheet([
  ["", "Mon. Feb. 02, 2026", "Tue. Feb. 03, 2026", "Wed. Feb. 04, 2026", "Thu. Feb. 05, 2026", "Fri. Feb. 06, 2026", "Sat. Feb. 07, 2026", "Sun. Feb. 08, 2026"],
  ["Richard Haydon", "Orange PM (on-call)", "AVAO PM", "Silver AM IC", "SSU SMS", "Rover AM", "Silver PM IC", ""],
  ["SENIOR MEDICAL STAFF", "", "", "", "", "", "", ""],
  ["HMOS", "", "", "", "", "", "", ""],
  ["Default HMO", "INTERN SSU AM", "Orange AM4", "Orange PM4", "", "", "", ""],
  ["Explicit HMO", "Orange PM4", "", "", "", "", "", ""],
  ["", "16:00-23:00", "", "", "", "", "", ""],
]), "Sheet1");

const ddhDefaultDoctors = doctorOptions([], ddhDefaultTimesWorkbook);
const promotedSms = ddhDefaultDoctors.find((doctor) => doctor.displayName.toLowerCase() === "richard haydon");
const defaultHmo = ddhDefaultDoctors.find((doctor) => doctor.displayName.toLowerCase() === "default hmo");
const explicitHmo = ddhDefaultDoctors.find((doctor) => doctor.displayName.toLowerCase() === "explicit hmo");
assert.ok(promotedSms);
assert.ok(defaultHmo);
assert.ok(explicitHmo);

const promotedSmsView = buildRosterView([], ddhDefaultTimesWorkbook, promotedSms.key);
assert.ok(promotedSmsView.events.length > 0);
assert.ok(promotedSmsView.events.every((event) => event.seniority === "SMS"));
assert.ok(promotedSmsView.events.some((event) => event.title === "DDH: Orange PM" && event.start.includes("15:00:00") && event.end.includes("00:00:00")));
assert.ok(promotedSmsView.events.some((event) => event.title === "DDH: AVAO PM" && event.start.includes("14:30:00") && event.end.includes("00:00:00")));
assert.ok(promotedSmsView.events.some((event) => event.title === "DDH: Silver AM" && event.start.includes("08:00:00") && event.end.includes("18:00:00")));
assert.ok(promotedSmsView.events.some((event) => event.title === "DDH: SSU" && event.start.includes("07:30:00") && event.end.includes("17:30:00")));
assert.ok(promotedSmsView.events.some((event) => event.title === "DDH: Rover AM" && event.start.includes("08:00:00") && event.end.includes("18:00:00")));
assert.ok(promotedSmsView.events.some((event) => event.title === "DDH: Silver PM" && event.start.includes("15:00:00") && event.end.includes("00:00:00")));

const defaultHmoView = buildRosterView([], ddhDefaultTimesWorkbook, defaultHmo.key);
assert.ok(defaultHmoView.events.some((event) => event.title === "DDH: SSU AM" && event.start.includes("07:30:00") && event.end.includes("17:30:00")));
assert.ok(defaultHmoView.events.some((event) => event.title === "DDH: Orange AM" && event.start.includes("08:00:00") && event.end.includes("18:00:00")));
assert.ok(defaultHmoView.events.some((event) => event.title === "DDH: Orange PM" && event.start.includes("14:30:00") && event.end.includes("00:00:00")));

const explicitHmoView = buildRosterView([], ddhDefaultTimesWorkbook, explicitHmo.key);
assert.ok(explicitHmoView.events.some((event) => event.title === "DDH: Orange PM" && event.start.includes("16:00:00") && event.end.includes("23:00:00")));

const ddhAccuracyWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(ddhAccuracyWorkbook, XLSX.utils.aoa_to_sheet([
  ["", "Mon. Feb. 02, 2026", "Tue. Feb. 03, 2026", "Wed. Feb. 04, 2026", "Thu. Feb. 05, 2026", "Fri. Feb. 06, 2026", "Sat. Feb. 07, 2026", "Sun. Feb. 08, 2026"],
  ["HMOS", "", "", "", "", "", "", ""],
  ["Accuracy Doctor", "Night4", "Orange AM2", "Silver PM3", "AM Fast (1)", "HMO SSU PM", "Orientation 11-13", ""],
  ["", "", "", "", "", "", "", ""],
  ["Crisis Locum Doctor", "Crisis Locum", "OFF", "Sun", "", "", "", ""],
  ["Sick Leave Doctor", "S\\L AM", "OFF", "Sun", "", "", "", ""],
  ["Conference Leave Doctor", "CME 19hrs", "", "", "", "", "", ""],
  ["Family Leave Doctor", "F/L", "", "", "", "", "", ""],
  ["Parental Leave Doctor", "Parental/L", "", "", "", "", "", ""],
  ["Parental Variant Doctor", "JMS Parental", "", "", "", "", "", ""],
  ["Special Leave Doctor", "Special leave", "", "", "", "", "", ""],
  ["Exam Leave Doctor", "EL", "", "", "", "", "", ""],
  ["Annual Leave Doctor", "AL 0.75", "", "", "", "", "", ""],
  ["Annotation Doctor", "MMC AM", "Casey PM", "HITH AM", "Tox CS", "VHH PM", "ARV AM", "Swing PM"],
  ["GED Doctor", "GED Junior", "", "", "", "", "", ""],
  ["DDH Notes Doctor", "Intern", "unavailable", "--", "pm>am", "sec", "(2 ED shifts this week)", ""],
  ["Senior Medical Staff", "", "", "", "", "", "", ""],
  ["AED Doctor", "AED", "", "", "", "", "", ""],
  ["PED Doctor", "PED", "", "", "", "", "", ""],
  ["Warragul Doctor", "Warragul", "", "", "", "", "", ""],
  ["Availability Doctor", "N", "Y", "W", "", "", "", ""],
  ["CS Request Doctor", "C/S", "CS not onsite PLS", "", "", "", "", ""],
  ["Leave Without Pay Doctor", "JMS LWP", "", "", "", "", "", ""],
  ["Extra Note Doctor", "Extra N/D", "Extra 10/4/26", "", "", "", "", ""],
  ["Extra Shift Doctor", "Extra AM", "Extra 16:00-01:00", "", "", "", "", ""],
  ["Untimed Swing Doctor", "Extra Swing", "", "", "", "", "", ""],
  ["Timed Swing Doctor", "Extra Swing", "Extra Swing", "", "", "", "", ""],
  ["", "13:00-22:00", "16:00-01:00", "", "", "", "", ""],
  ["Frank Annotation Doctor", "MDC", "NA", "MDC NA", "Comm", "NA PM", "Comm for safety", ""],
  ["AMP", "", "", "", "", "", "", ""],
  ["AMP Roster Doctor", "Physiotherapist", "Jane Example", "Swing PM", "", "", "", ""],
  ["Orientation Variant Doctor", "Orienation 0800-1000", "", "", "", "", "", ""],
  ["Senior Medical Staff", "", "", "", "", "", "", ""],
  ["DDH Cleanup Doctor", "CS Tox", "clinical support", "clinical", "OSCE", "Swing FT", "CS -On-site", "CS not onsite pls"],
  ["DDH VHH Warning Doctor", "08H00", "10H00", "14H30", "", "", "", ""],
  ["", "VHH", "VHH", "VHH", "", "", "", ""],
  ["DDH Request Doctor", "Moved to Sat 11/4", "PM Austin instead of AM", "See Amy", "working 22/2", "pm onlu", "y but not 0730", "wweeeewe"],
]), "Sheet1");
const ddhAccuracyDoctors = doctorOptions([], ddhAccuracyWorkbook);
const ddhAccuracyDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "ACCURACY DOCTOR");
const crisisLocumDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "CRISIS LOCUM DOCTOR");
const sickLeaveDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "SICK LEAVE DOCTOR");
const conferenceLeaveDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "CONFERENCE LEAVE DOCTOR");
const ddhFamilyLeaveDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "FAMILY LEAVE DOCTOR");
const ddhParentalLeaveDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "PARENTAL LEAVE DOCTOR");
const ddhParentalVariantDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "PARENTAL VARIANT DOCTOR");
const ddhSpecialLeaveDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "SPECIAL LEAVE DOCTOR");
const ddhExamLeaveDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "EXAM LEAVE DOCTOR");
const ddhAnnualLeaveDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "ANNUAL LEAVE DOCTOR");
const ddhAnnotationDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "ANNOTATION DOCTOR");
const ddhGedDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "GED DOCTOR");
const ddhNotesDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "DDH NOTES DOCTOR");
const ddhAedDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "AED DOCTOR");
const ddhPedDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "PED DOCTOR");
const ddhWarragulDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "WARRAGUL DOCTOR");
const ddhAvailabilityDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "AVAILABILITY DOCTOR");
const ddhCsRequestDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "CS REQUEST DOCTOR");
const ddhLeaveWithoutPayDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "LEAVE WITHOUT PAY DOCTOR");
const ddhExtraNoteDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "EXTRA NOTE DOCTOR");
const ddhExtraShiftDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "EXTRA SHIFT DOCTOR");
const ddhUntimedSwingDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "UNTIMED SWING DOCTOR");
const ddhTimedSwingDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "TIMED SWING DOCTOR");
const ddhFrankAnnotationDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "FRANK ANNOTATION DOCTOR");
const ddhAmpRosterDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "AMP ROSTER DOCTOR");
const ddhOrientationVariantDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "ORIENTATION VARIANT DOCTOR");
const ddhSmsCleanupDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "DDH CLEANUP DOCTOR");
const ddhVhhWarningDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "DDH VHH WARNING DOCTOR");
const ddhRequestDoctor = ddhAccuracyDoctors.find((doctor) => doctor.key === "DDH REQUEST DOCTOR");
assert.ok(ddhAccuracyDoctor && crisisLocumDoctor && sickLeaveDoctor && conferenceLeaveDoctor && ddhFamilyLeaveDoctor && ddhParentalLeaveDoctor && ddhParentalVariantDoctor && ddhSpecialLeaveDoctor && ddhExamLeaveDoctor && ddhAnnualLeaveDoctor && ddhAnnotationDoctor && ddhGedDoctor && ddhNotesDoctor && ddhAedDoctor && ddhPedDoctor && ddhWarragulDoctor && ddhAvailabilityDoctor && ddhCsRequestDoctor && ddhLeaveWithoutPayDoctor && ddhExtraNoteDoctor && ddhExtraShiftDoctor && ddhUntimedSwingDoctor && ddhTimedSwingDoctor && ddhFrankAnnotationDoctor && ddhAmpRosterDoctor && ddhOrientationVariantDoctor && ddhSmsCleanupDoctor && ddhVhhWarningDoctor && ddhRequestDoctor, "DDH accuracy fixtures should expose their rostered doctors");
const ddhAccuracyView = buildRosterView([], ddhAccuracyWorkbook, ddhAccuracyDoctor.key);
assert.deepEqual(
  ddhAccuracyView.events.map((event) => [event.title, event.rawValue, event.timeLabel]),
  [
    ["DDH: Night", "Night4", "23:00-08:30"],
    ["DDH: Orange AM", "Orange AM2", "08:00-18:00"],
    ["DDH: Silver PM", "Silver PM3", "14:30-00:00"],
    ["DDH: FAST AM", "AM Fast (1)", "08:00-18:00"],
    ["DDH: SSU PM", "HMO SSU PM", "14:30-00:00"],
    ["DDH: Orientation", "Orientation 11-13", "11:00-13:00"],
  ],
  "numbered DDH roster slots should be recognised and orientation should retain its times",
);
assert.equal(ddhAccuracyView.issues.length, 0, "recognised DDH slot variants should not remain unresolved");
const crisisLocumView = buildRosterView([], ddhAccuracyWorkbook, crisisLocumDoctor.key);
assert.deepEqual(
  crisisLocumView.events.map((event) => [event.title, event.rawValue, event.allDay]),
  [
    ["DDH: Crisis Locum", "Crisis Locum", true],
  ],
  "Crisis Locum should be a conservative all-day shift while OFF and weekday annotations remain hidden",
);
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, sickLeaveDoctor.key).events.map((event) => event.title), ["Sick leave"], "backslash sick-leave notation should be recognised");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, conferenceLeaveDoctor.key).events.map((event) => event.title), ["Conference Leave"], "CME hour annotations should be recognised as conference leave");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhFamilyLeaveDoctor.key).events.map((event) => event.title), ["Family Leave"], "F/L should be recognised as family leave");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhParentalLeaveDoctor.key).events.map((event) => event.title), ["Parental Leave"], "Parental/L should be recognised as parental leave");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhParentalVariantDoctor.key).events.map((event) => event.title), ["Parental Leave"], "parental-leave variants should be recognised as parental leave");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhSpecialLeaveDoctor.key).events.map((event) => event.title), ["Special Leave"], "Special leave should be recognised");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhExamLeaveDoctor.key).events.map((event) => event.title), ["Exam Leave"], "EL should be recognised as exam leave");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhAnnualLeaveDoctor.key).events.map((event) => event.title), ["Annual Leave"], "fractional AL should be recognised as annual leave");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhAnnotationDoctor.key).events.map((event) => event.title), ["DDH: TOX Clinical Support", "DDH: Swing PM"], "DDH TOX Clinical Support should remain visible while other-hospital safety annotations stay hidden");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhGedDoctor.key).events.map((event) => event.title), ["DDH: GED shift"], "GED Junior should be a recognised GED shift");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhNotesDoctor.key).events.map((event) => event.title), ["DDH: TOX SEC"], "a standalone SMS SEC allocation should retain its roster abbreviation");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhNotesDoctor.key).issues, [], "DDH headings and availability notes should not become unresolved shift codes");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhCsRequestDoctor.key).events.map((event) => event.title), ["DDH: CS"], "C/S should resolve to Clinical Support while CS not onsite PLS remains a hidden request");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhAedDoctor.key).events, [], "AED is an MMC allocation annotation, not a DDH calendar shift");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhPedDoctor.key).events, [], "PED is an MMC allocation annotation, not a DDH calendar shift");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhAedDoctor.key).issues, [], "AED should not remain unresolved once hidden as an MMC allocation annotation");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhWarragulDoctor.key).events, [], "Warragul is an external-hospital annotation, not a DDH calendar shift");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhAvailabilityDoctor.key).events, [], "DDH availability notes should not become calendar events");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhLeaveWithoutPayDoctor.key).events.map((event) => event.title), ["Leave without pay"], "LWP variants should be recognised as leave without pay");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhExtraNoteDoctor.key).events, [], "untimed Extra payment annotations should not become shifts on the listed day");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhExtraShiftDoctor.key).events.map((event) => [event.title, event.timeLabel]), [["DDH: Extra AM", "08:00-18:00"], ["DDH: Extra PM", "16:00-01:00"]], "period-labelled or explicitly timed Extra entries should remain calendar shifts");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhUntimedSwingDoctor.key).events.map((event) => [event.title, event.allDay]), [["DDH: Swing shift", true]], "untimed Swing entries should remain visible as all-day Swing shifts");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhTimedSwingDoctor.key).events.map((event) => [event.title, event.timeLabel]), [["DDH: Swing AM", "13:00-22:00"], ["DDH: Swing PM", "16:00-01:00"]], "timed Swing entries should classify AM before 14:00 and PM after 15:00");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhFrankAnnotationDoctor.key).events, [], "Frank's DDH MDC, NA, and communication annotations should not become calendar events");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhFrankAnnotationDoctor.key).issues, [], "Frank's DDH MDC, NA, and communication annotations should not become unresolved shift codes");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhAmpRosterDoctor.key).events.map((event) => event.title), ["DDH: Swing PM"], "AMP supervision headings and names should not become calendar events while genuine AMP shifts remain visible");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhAmpRosterDoctor.key).issues, [], "AMP supervision headings and names should not become unresolved shift codes");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhOrientationVariantDoctor.key).events.map((event) => [event.title, event.timeLabel]), [["DDH: Orientation", "08:00-10:00"]], "orientation spelling variants should standardise to a timed Orientation shift");
assert.deepEqual(
  buildRosterView([], ddhAccuracyWorkbook, ddhSmsCleanupDoctor.key).events.map((event) => [event.title, event.allDay]),
  [["DDH: CS Tox", true], ["DDH: CS", true], ["DDH: Clinical", true], ["DDH: CS Exam", true], ["DDH: Swing shift", true], ["DDH: CS onsite", true]],
  "approved DDH SMS aliases should create calendar shifts while not-onsite requests remain hidden",
);
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhSmsCleanupDoctor.key).issues, [], "approved DDH SMS aliases should not remain unresolved");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhVhhWarningDoctor.key).events, [], "VHH late-early warning time fragments should stay off DDH calendars");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhVhhWarningDoctor.key).issues, [], "VHH late-early warning time fragments should not remain unresolved");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhRequestDoctor.key).events, [], "DDH roster-writer requests and annotations should stay off calendars");
assert.deepEqual(buildRosterView([], ddhAccuracyWorkbook, ddhRequestDoctor.key).issues, [], "DDH roster-writer requests and annotations should not remain unresolved");

const mmcPdfBytes = await readFile(fileURLToPath(new URL("../fixtures/AdultMMCTerm2.2026.Ver1.pdf", import.meta.url)));
const formData = new FormData();
formData.append("rosterFiles", new File([mmcPdfBytes], "AdultMMCTerm2.2026.Ver1.pdf", { type: "application/pdf" }));
const parsedPdf = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: formData }));
const pdfDoctors = doctorOptions(parsedPdf.sources.mmc, parsedPdf.sources.ddh);
assert.ok(pdfDoctors.length > 50);
assert.ok(pdfDoctors.find((doctor) => doctor.displayName === "Richard HAYDON"));
assert.ok(pdfDoctors.find((doctor) => doctor.displayName === "Abi THANIKASALAM"));
assert.ok(pdfDoctors.find((doctor) => doctor.displayName === "Titus HACKMAN"));
const pdfRichard = pdfDoctors.find((doctor) => doctor.displayName === "Richard HAYDON");
const pdfView = buildRosterView(parsedPdf.sources.mmc, parsedPdf.sources.ddh, pdfRichard.key);
assert.ok(pdfView.events.some((event) => event.title === "MMC: SSU PM"));
assert.ok(pdfView.events.some((event) => event.rawValue === "0800-1730" && event.title === "MMC: AM shift"));

class MemoryStore {
  constructor() {
    this.records = new Map();
    this.deletedKeys = [];
    this.accountListCalls = 0;
    this.accountGetCalls = 0;
    this.d1 = new MemoryD1();
    this.r2 = new MemoryR2();
  }

  async get(key, type) {
    if (String(key || "").startsWith("account:")) this.accountGetCalls += 1;
    const value = this.records.get(key);
    if (value === undefined) return null;
    return type === "json" ? JSON.parse(value) : value;
  }

  async put(key, value) {
    this.records.set(key, String(value));
  }

  async delete(key) {
    this.deletedKeys.push(key);
    this.records.delete(key);
  }

  async list(options = {}) {
    const prefix = options.prefix || "";
    if (prefix === "account:") this.accountListCalls += 1;
    return {
      keys: [...this.records.keys()].filter((name) => name.startsWith(prefix)).map((name) => ({ name })),
    };
  }

  resetMetrics() {
    this.accountListCalls = 0;
    this.accountGetCalls = 0;
  }
}

class MemoryR2 {
  constructor() {
    this.objects = new Map();
  }

  async put(key, value, options = {}) {
    const bytes = value instanceof Uint8Array ? value : new Uint8Array(value);
    this.objects.set(key, {
      bytes,
      httpMetadata: options.httpMetadata || {},
    });
  }

  async get(key) {
    const item = this.objects.get(key);
    if (!item) return null;
    return {
      httpMetadata: item.httpMetadata,
      arrayBuffer: async () => item.bytes.buffer.slice(item.bytes.byteOffset, item.bytes.byteOffset + item.bytes.byteLength),
    };
  }

  async delete(keyOrKeys) {
    const keys = Array.isArray(keyOrKeys) ? keyOrKeys : [keyOrKeys];
    for (const key of keys) {
      if (key) this.objects.delete(key);
    }
  }
}

class MemoryD1 {
  constructor() {
    this.files = new Map();
    this.doctors = new Map();
    this.fileDoctors = new Map();
    this.facilityStaffDesignations = new Map();
    this.facilitySmsMemberships = new Map();
    this.events = new Map();
    this.dailyPresence = new Map();
    this.issues = new Map();
    this.rawFiles = new Map();
    this.rosterSources = new Map();
    this.rosterSyncRuns = new Map();
    this.rosterDispatches = new Map();
    this.accountProfiles = new Map();
    this.accountClaims = new Map();
    this.rosterPeople = new Map();
    this.rosterPersonAliases = new Map();
    this.accountPeople = new Map();
    this.accountStates = new Map();
    this.accountHospitalLocations = new Map();
    this.canonicalDoctors = new Map();
    this.customEvents = new Map();
    this.subscriptionTokens = new Map();
    this.parserRules = new Map();
    this.parserRuleSuggestions = new Map();
    this.doctorProfiles = new Map();
    this.snapshotRegistry = new Map();
    this.consoleMessages = [];
    this.nextConsoleMessageId = 1;
    this.failNextEventInsert = false;
  }

  prepare(sql) {
    return new MemoryD1Statement(this, sql);
  }

  async batch(statements) {
    const snapshots = Object.fromEntries([
      "files",
      "doctors",
      "fileDoctors",
      "facilityStaffDesignations",
      "facilitySmsMemberships",
      "events",
      "dailyPresence",
      "issues",
      "rawFiles",
      "rosterSources",
      "rosterSyncRuns",
      "rosterDispatches",
      "accountProfiles",
      "accountClaims",
      "rosterPeople",
      "rosterPersonAliases",
      "accountPeople",
      "accountStates",
      "accountHospitalLocations",
      "canonicalDoctors",
      "customEvents",
      "subscriptionTokens",
      "parserRules",
      "parserRuleSuggestions",
      "doctorProfiles",
      "snapshotRegistry",
    ].map((key) => [key, new Map(this[key])]));
    try {
      const results = [];
      for (const statement of statements) results.push(await statement.run());
      return results;
    } catch (error) {
      for (const [key, value] of Object.entries(snapshots)) this[key] = value;
      throw error;
    }
  }
}

class MemoryD1Statement {
  constructor(db, sql) {
    this.db = db;
    this.sql = String(sql || "").replace(/\s+/g, " ").trim();
    this.args = [];
  }

  bind(...args) {
    this.args = args;
    return this;
  }

  async run() {
    const sql = this.sql;
    const args = this.args;
    if (sql.startsWith("CREATE ")) return { success: true };
    if (sql.startsWith("ALTER TABLE")) return { success: true };
    if (sql.startsWith("INSERT INTO roster_files")) {
      this.db.files.set(args[0], {
        id: args[0],
        name: args[1],
        source_type: args[2],
        source_id: args[3],
        active: args[4],
        size: args[5],
        last_modified: args[6],
        added_at: args[7],
        uploaded_at: args[8],
        uploaded_by: args[9],
        parsed_at: args[10],
        parser_version: args[11] || "",
      });
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO raw_roster_files")) {
      this.db.rawFiles.set(args[0], {
        file_id: args[0],
        name: args[1],
        source_type: args[2],
        size: args[3],
        last_modified: args[4],
        object_key: args[5],
        type: args[6],
        data_url: args[7],
        uploaded_at: args[8],
      });
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO roster_sources")) {
      this.db.rosterSources.set(args[0], {
        id: args[0], provider: args[1], source_type: args[2], label: args[3], enabled: args[4],
        config_json: args[5], cursor_json: args[6], provider_version: args[7], provider_modified_at: args[8],
        last_checked_at: args[9], last_success_at: args[10], last_error: args[11], active_file_id: args[12],
        created_at: args[13], updated_at: args[14],
      });
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO roster_sync_runs")) {
      this.db.rosterSyncRuns.set(args[0], {
        id: args[0], source_id: args[1], trigger_type: args[2], provider_version: args[3],
        content_hash: args[4], file_id: args[5], source_file_id: args[6], status: args[7], message: args[8],
        doctor_count: args[9], event_count: args[10], started_at: args[11], completed_at: args[12],
      });
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO roster_dispatches")) {
      this.db.rosterDispatches.set(args[0], {
        id: args[0], status: "requested", reason: args[1], github_run_id: "", requested_at: args[2],
        accepted_at: "", started_at: "", completed_at: "", retry_after: args[3], attempt_count: args[4], last_error: "",
      });
      return { success: true };
    }
    if (sql.startsWith("UPDATE roster_dispatches")) {
      const dispatch = this.db.rosterDispatches.get(args[8]);
      if (dispatch) Object.assign(dispatch, {
        status: args[0], github_run_id: args[1], accepted_at: args[2], started_at: args[3], completed_at: args[4],
        retry_after: args[5], attempt_count: args[6], last_error: args[7],
      });
      return { success: true };
    }
    if (sql.startsWith("UPDATE roster_sync_runs") && sql.includes("status = 'processing'")) {
      const run = this.db.rosterSyncRuns.get(args[1]);
      if (run) Object.assign(run, { status: "processing", message: args[0], completed_at: "" });
      return { success: true };
    }
    if (sql.startsWith("UPDATE roster_sync_runs") && sql.includes("status = 'superseded'")) {
      if (sql.includes("AS stale_run")) return { success: true, meta: { changes: 0 } };
      let changes = 0;
      for (const run of this.db.rosterSyncRuns.values()) {
        const raw = this.db.rawFiles.get(run.source_file_id || run.file_id);
        if (run.source_id !== args[2] || run.provider_version !== args[3] || run.id === args[4]) continue;
        if (!["queued", "processing"].includes(run.status)) continue;
        if (String(raw?.name || "").toLowerCase() !== String(args[5] || "").toLowerCase()) continue;
        Object.assign(run, { status: "superseded", message: args[0], completed_at: args[1] });
        changes += 1;
      }
      return { success: true, meta: { changes } };
    }
    if (sql.startsWith("UPDATE roster_sync_runs")) {
      const run = this.db.rosterSyncRuns.get(args[6]);
      if (run) Object.assign(run, {
        status: args[0], message: args[1], file_id: args[2], doctor_count: args[3], event_count: args[4], completed_at: args[5],
      });
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM roster_file_doctors")) {
      for (const [key, doctor] of [...this.db.fileDoctors.entries()]) {
        if (!key.startsWith(`${args[0]}|`)) continue;
        if (sql.includes("NOT EXISTS") && [...this.db.events.values()].some((event) => event.file_id === doctor.file_id && event.doctor_key === doctor.doctor_key)) continue;
        this.db.fileDoctors.delete(key);
      }
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM roster_events")) {
      for (const [key, value] of [...this.db.events.entries()]) {
        if (value.file_id !== args[0]) continue;
        if (sql.includes("start_date <= ?") && !(value.start_date <= args[1] && value.end_date >= args[2])) continue;
        this.db.events.delete(key);
      }
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM roster_issues")) {
      for (const [key, value] of [...this.db.issues.entries()]) {
        if (value.file_id !== args[0]) continue;
        if (sql.includes("start_date >= ?") && !(value.start_date >= args[1] && value.start_date <= args[2])) continue;
        this.db.issues.delete(key);
      }
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM roster_daily_presence")) {
      const prefix = String(args[0] || "").replace(/\\([%_\\])/g, "$1").replace(/%$/, "");
      for (const [key, value] of [...this.db.dailyPresence.entries()]) {
        if (String(value.event_id || "").startsWith(prefix)) this.db.dailyPresence.delete(key);
      }
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM roster_files")) {
      this.db.files.delete(args[0]);
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM raw_roster_files")) {
      this.db.rawFiles.delete(args[0]);
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM roster_doctors WHERE source_type = ?")) {
      const sourceType = args[0];
      const referencedKeys = new Set(
        [...this.db.fileDoctors.values()]
          .filter((doctor) => doctor.source_type === sourceType)
          .map((doctor) => doctor.doctor_key),
      );
      for (const [key, doctor] of [...this.db.doctors.entries()]) {
        if (doctor.source_type === sourceType && !referencedKeys.has(doctor.doctor_key)) this.db.doctors.delete(key);
      }
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO roster_doctors")) {
      for (let index = 0; index < args.length; index += 4) {
        this.db.doctors.set(`${args[index]}|${args[index + 1]}`, {
          source_type: args[index],
          doctor_key: args[index + 1],
          display_name: args[index + 2],
          updated_at: args[index + 3],
        });
      }
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO roster_file_doctors")) {
      const width = sql.includes("membership_source") ? 6 : 4;
      for (let index = 0; index < args.length; index += width) {
        this.db.fileDoctors.set(`${args[index]}|${args[index + 1]}|${args[index + 2]}`, {
          file_id: args[index],
          source_type: args[index + 1],
          doctor_key: args[index + 2],
          display_name: args[index + 3],
          seniority: args[index + 4] || "",
          membership_source: args[index + 5] || "roster",
        });
      }
      return { success: true };
    }
    if (sql.includes("INSERT INTO facility_sms_memberships")) {
      const fileId = sql.startsWith("WITH file_coverage") ? args[0] : "";
      const matchingDoctors = [...this.db.fileDoctors.values()]
        .filter((doctor) => !fileId || doctor.file_id === fileId)
        .filter((doctor) => String(doctor.seniority || "").toUpperCase() === "SMS");
      for (const doctor of matchingDoctors) {
        const dates = [...this.db.events.values()]
          .filter((event) => event.file_id === doctor.file_id && event.doctor_key === doctor.doctor_key)
          .map((event) => event.start_date)
          .filter(Boolean)
          .sort();
        if (!dates.length) continue;
        const key = `${doctor.source_type}|${doctor.doctor_key}`;
        const existing = this.db.facilitySmsMemberships.get(key);
        this.db.facilitySmsMemberships.set(key, {
          source_type: doctor.source_type, doctor_key: doctor.doctor_key, display_name: doctor.display_name,
          first_seen_date: existing ? [existing.first_seen_date, dates[0]].sort()[0] : dates[0],
          last_seen_date: existing ? [existing.last_seen_date, dates.at(-1)].sort().at(-1) : dates.at(-1),
        });
      }
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO roster_events")) {
      if (this.db.failNextEventInsert) {
        this.db.failNextEventInsert = false;
        throw new Error("Injected event insert failure.");
      }
      for (let index = 0; index < args.length; index += 16) {
        this.db.events.set(args[index], {
          id: args[index],
          file_id: args[index + 1],
          source_type: args[index + 2],
          doctor_key: args[index + 3],
          display_name: args[index + 4],
          start_date: args[index + 5],
          end_date: args[index + 6],
          start_ts: args[index + 7],
          end_ts: args[index + 8],
          title: args[index + 9],
          raw_value: args[index + 10],
          seniority: args[index + 11],
          location: args[index + 12],
          all_day: args[index + 13],
          time_label: args[index + 14],
          event_json: args[index + 15],
        });
      }
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO roster_issues")) {
      for (let index = 0; index < args.length; index += 14) {
        this.db.issues.set(args[index], {
          id: args[index],
          file_id: args[index + 1],
          source_type: args[index + 2],
          doctor_key: args[index + 3],
          display_name: args[index + 4],
          start_date: args[index + 5],
          raw_value: args[index + 6],
          seniority: args[index + 7],
          status: args[index + 8],
          message: args[index + 9],
          resolution_type: args[index + 10],
          suggested_title: args[index + 11],
          time_label: args[index + 12],
          issue_json: args[index + 13],
        });
      }
      return { success: true };
    }
    if (sql.startsWith("INSERT OR IGNORE INTO roster_daily_presence")) {
      for (let index = 0; index < args.length; index += 5) {
        const row = {
          date: args[index],
          source_type: args[index + 1],
          doctor_key: args[index + 2],
          display_name: args[index + 3],
          event_id: args[index + 4],
        };
        const key = `${row.date}|${row.source_type}|${row.doctor_key}|${row.event_id}`;
        if (!this.db.dailyPresence.has(key)) this.db.dailyPresence.set(key, row);
      }
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO account_profiles")) {
      const previous = this.db.accountProfiles.get(args[0]) || {};
      this.db.accountProfiles.set(args[0], {
        ...previous,
        email: args[0],
        real_name: args[1],
        role: args[2],
        insights_enabled: args[3],
        facility_overview_enabled: args[4],
        non_clinical: args[5],
        director_view_enabled: args[6],
        subscription_token: args[7],
        password_salt: args[8] || previous.password_salt || "",
        password_hash: args[9] || previous.password_hash || "",
        admin_issues_json: args[10] || "[]",
        local_parser_extensions_json: args[11] || "[]",
        created_at: args[12] || previous.created_at || "",
        updated_at: args[13] || args[8],
      });
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM account_claims")) {
      for (const key of [...this.db.accountClaims.keys()]) if (key.startsWith(`${args[0]}|`)) this.db.accountClaims.delete(key);
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM subscription_tokens WHERE email")) {
      for (const [token, row] of [...this.db.subscriptionTokens.entries()]) {
        if (row.email === args[0]) this.db.subscriptionTokens.delete(token);
      }
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO subscription_tokens")) {
      this.db.subscriptionTokens.set(args[0], {
        token: args[0],
        email: args[1],
        created_at: args[2],
        updated_at: args[3],
      });
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO parser_rules")) {
      this.db.parserRules.set(args[0], {
        id: args[0],
        scope: "global",
        source_type: args[1],
        seniority: args[2],
        code: args[3],
        title: args[4],
        rule_json: args[5],
        updated_at: args[6],
      });
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM parser_rules")) {
      for (const [key, rule] of [...this.db.parserRules.entries()]) {
        if (rule.scope === "global" && rule.source_type === args[0] && rule.seniority === args[1] && rule.code === args[2]) {
          this.db.parserRules.delete(key);
        }
      }
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM parser_rule_suggestions")) {
      this.db.parserRuleSuggestions.clear();
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO parser_rule_suggestions")) {
      this.db.parserRuleSuggestions.set(args[0], {
        id: args[0],
        email: args[1],
        status: "pending",
        suggestion_json: args[2],
        created_at: args[3],
        updated_at: args[4],
      });
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO account_claims")) {
      for (let index = 0; index < args.length; index += 6) {
        this.db.accountClaims.set(`${args[index]}|${args[index + 1]}|${args[index + 2]}`, {
          email: args[index],
          source_type: args[index + 1],
          doctor_key: args[index + 2],
          display_name: args[index + 3],
          matched_at: args[index + 4],
          updated_at: args[index + 5],
        });
      }
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO roster_people")) {
      const previous = this.db.rosterPeople.get(args[0]) || {};
      this.db.rosterPeople.set(args[0], { ...previous, person_id: args[0], preferred_display_name: args[1] || previous.preferred_display_name || "", updated_at: args[3] });
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO account_people")) {
      this.db.accountPeople.set(args[0], { email: args[0], person_id: args[1], created_at: args[2], updated_at: args[3] });
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO roster_person_aliases")) {
      this.db.rosterPersonAliases.set(`${args[0]}|${args[1]}`, { source_type: args[0], doctor_key: args[1], display_name: args[2], person_id: args[3], updated_at: args[5] });
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO account_states")) {
      this.db.accountStates.set(args[0], {
        email: args[0],
        session_json: args[1],
        updated_at: args[2],
      });
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM account_states")) {
      this.db.accountStates.delete(args[0]);
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM account_hospital_locations")) {
      for (const key of [...this.db.accountHospitalLocations.keys()]) if (key.startsWith(`${args[0]}|`)) this.db.accountHospitalLocations.delete(key);
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO account_hospital_locations")) {
      for (let index = 0; index < args.length; index += 4) {
        this.db.accountHospitalLocations.set(`${args[index]}|${args[index + 1]}`, {
          email: args[index],
          source_type: args[index + 1],
          location: args[index + 2],
          updated_at: args[index + 3],
        });
      }
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM canonical_doctors")) {
      this.db.canonicalDoctors.clear();
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO canonical_doctors")) {
      this.db.canonicalDoctors.set(args[0], {
        canonical_key: args[0],
        display_name: args[1],
        source_type: args[2],
        source_types_json: args[3],
        aliases_json: args[4],
        has_events: args[5],
      });
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM custom_events")) {
      for (const key of [...this.db.customEvents.keys()]) if (key.startsWith(`${args[0]}|`)) this.db.customEvents.delete(key);
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO custom_events")) {
      this.db.customEvents.set(`${args[0]}|${args[1]}`, {
        owner_email: args[0],
        id: args[1],
        title: args[2],
        start_date: args[3],
        end_date: args[4],
        all_day: args[5],
        start_time: args[6],
        end_time: args[7],
        location: args[8],
        include: args[9],
      });
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM account_profiles")) {
      this.db.accountProfiles.delete(args[0]);
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO doctor_profiles")) {
      this.db.doctorProfiles.set(args[0], {
        profile_id: args[0],
        doctor_key: args[1],
        display_name: args[2],
        source_types_json: args[3],
        state_json: args[4],
        created_at: args[5],
        updated_at: args[6],
      });
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO snapshot_registry")) {
      const row = {
        owner_type: args[0],
        owner_id: args[1],
        doctor_key: args[2],
        range_key: args[3],
        requested_revision: args[4],
        built_revision: args[5],
        status: args[6],
        artifact_key: args[7],
        built_at: args[8],
        size_bytes: args[9],
        build_ms: args[10],
        last_error: args[11],
        updated_at: args[12],
      };
      this.db.snapshotRegistry.set(`${row.owner_type}|${row.owner_id}|${row.doctor_key}|${row.range_key}`, row);
      return { success: true };
    }
    if (sql.startsWith("INSERT INTO console_messages")) {
      this.db.consoleMessages.push({
        id: this.db.nextConsoleMessageId++,
        actor_email: args[0],
        message: args[1],
        is_error: args[2],
        created_at: args[3],
      });
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM console_messages")) {
      this.db.consoleMessages = [...this.db.consoleMessages]
        .sort((left, right) => String(right.created_at).localeCompare(String(left.created_at)) || right.id - left.id)
        .slice(0, 50);
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM doctor_profiles")) {
      this.db.doctorProfiles.delete(args[0]);
      return { success: true };
    }
    if (sql.startsWith("DELETE FROM snapshot_registry")) {
      for (const [key, row] of [...this.db.snapshotRegistry.entries()]) {
        if (row.owner_type === args[0] && row.owner_id === args[1]) this.db.snapshotRegistry.delete(key);
      }
      return { success: true };
    }
    if (sql.startsWith("UPDATE roster_files SET active")) {
      const file = this.db.files.get(args[1]);
      if (file) {
        file.active = sql.includes("active = 1") ? 1 : args[0];
        if (sql.includes("parser_version")) file.parser_version = args[0];
      }
      return { success: true };
    }
    if (sql.startsWith("UPDATE facility_staff_designations")) {
      return { success: true, meta: { changes: 0 } };
    }
    throw new Error(`Unsupported MemoryD1 run SQL: ${sql}`);
  }

  async all() {
    const sql = this.sql;
    const args = this.args;
    if (sql.startsWith("PRAGMA table_info(roster_files)")) {
      return {
        results: ["id", "name", "source_type", "source_id", "active", "size", "last_modified", "added_at", "uploaded_at", "uploaded_by", "parsed_at", "parser_version"].map((name) => ({ name })),
      };
    }
    if (sql.startsWith("PRAGMA table_info(roster_sync_runs)")) {
      return {
        results: ["id", "source_id", "trigger_type", "provider_version", "content_hash", "file_id", "source_file_id", "status", "message", "doctor_count", "event_count", "started_at", "completed_at"].map((name) => ({ name })),
      };
    }
    if (sql.startsWith("PRAGMA table_info(roster_file_doctors)")) {
      return { results: ["file_id", "source_type", "doctor_key", "display_name", "seniority", "membership_source", "provider_staff_id"].map((name) => ({ name })) };
    }
    if (sql.startsWith("PRAGMA table_info(roster_events)")) {
      return { results: ["id", "file_id", "source_type", "doctor_key", "display_name", "start_date", "end_date", "start_ts", "end_ts", "title", "raw_value", "seniority", "provider_staff_id", "location", "all_day", "time_label", "event_json"].map((name) => ({ name })) };
    }
    if (sql.startsWith("PRAGMA table_info(roster_sources)")) {
      return {
        results: ["id", "provider", "source_type", "label", "enabled", "config_json", "cursor_json", "provider_version", "provider_modified_at", "last_checked_at", "last_success_at", "last_error", "active_file_id", "created_at", "updated_at"].map((name) => ({ name })),
      };
    }
    if (sql.startsWith("PRAGMA table_info(account_profiles)")) {
      return {
        results: [
          "email",
          "real_name",
          "role",
          "insights_enabled",
          "facility_overview_enabled",
          "non_clinical",
          "director_view_enabled",
          "subscription_token",
          "password_salt",
          "password_hash",
          "admin_issues_json",
          "local_parser_extensions_json",
          "created_at",
          "updated_at",
        ].map((name) => ({ name })),
      };
    }
    if (sql.startsWith("PRAGMA table_info(raw_roster_files)")) {
      return {
        results: ["file_id", "name", "source_type", "size", "last_modified", "object_key", "type", "data_url", "uploaded_at"].map((name) => ({ name })),
      };
    }
    if (sql.startsWith("PRAGMA table_info(roster_issues)")) {
      return {
        results: [
          "id",
          "file_id",
          "source_type",
          "doctor_key",
          "display_name",
          "start_date",
          "raw_value",
          "seniority",
          "status",
          "message",
          "resolution_type",
          "suggested_title",
          "time_label",
          "issue_json",
        ].map((name) => ({ name })),
      };
    }
    if (sql.startsWith("SELECT * FROM roster_sources ORDER BY")) {
      return { results: [...this.db.rosterSources.values()].sort((left, right) => String(left.label).localeCompare(String(right.label)) || String(left.id).localeCompare(String(right.id))) };
    }
    if (sql.includes("FROM facility_staff_designations")) {
      return { results: [] };
    }
    if (sql.includes("FROM facility_sms_memberships")) {
      const termEnd = args[0] || "9999-12-31";
      const sourceType = args.length > 1 ? args[1] : "";
      return { results: [...this.db.facilitySmsMemberships.values()]
        .filter((row) => row.first_seen_date <= termEnd)
        .filter((row) => !sourceType || row.source_type === sourceType) };
    }
    if (sql.startsWith("SELECT * FROM roster_sync_runs") && !sql.includes("WHERE source_id")) {
      return { results: [...this.db.rosterSyncRuns.values()].sort((left, right) => String(right.started_at).localeCompare(String(left.started_at)) || String(right.id).localeCompare(String(left.id))).slice(0, Number(args[0] || 50)) };
    }
    if (sql.startsWith("SELECT * FROM roster_dispatches ORDER BY")) {
      return { results: [...this.db.rosterDispatches.values()].sort((left, right) => String(right.requested_at).localeCompare(String(left.requested_at))).slice(0, 1) };
    }
    if (sql.includes("FROM roster_file_doctors") && sql.includes("file_name") && sql.includes("event_count")) {
      const requestedKeys = sql.includes("roster_file_doctors.doctor_key IN") ? new Set(args) : null;
      return {
        results: [...this.db.fileDoctors.values()]
          .filter((doctor) => this.db.files.get(doctor.file_id)?.active === 1)
          .filter((doctor) => !requestedKeys || requestedKeys.has(doctor.doctor_key))
          .map((doctor) => {
            const file = this.db.files.get(doctor.file_id);
            return {
              file_id: doctor.file_id,
              file_name: file?.name || "",
              file_source_type: file?.source_type || doctor.source_type,
              active: file?.active ?? 0,
              source_type: doctor.source_type,
              doctor_key: doctor.doctor_key,
              display_name: doctor.display_name,
              event_count: [...this.db.events.values()].filter((event) => event.file_id === doctor.file_id && event.doctor_key === doctor.doctor_key).length,
            };
          })
          .sort((left, right) => String(left.file_id).localeCompare(String(right.file_id)) || String(left.display_name).localeCompare(String(right.display_name))),
      };
    }
    if (sql.includes("FROM canonical_doctors")) {
      return {
        results: [...this.db.canonicalDoctors.values()]
          .filter((doctor) => !sql.includes("WHERE has_events = 1") || doctor.has_events === 1)
          .sort((left, right) => left.display_name.localeCompare(right.display_name) || left.source_type.localeCompare(right.source_type)),
      };
    }
    if (sql.includes("FROM custom_events")) {
      return {
        results: [...this.db.customEvents.values()]
          .filter((event) => event.owner_email === args[0])
          .sort((left, right) => left.start_date.localeCompare(right.start_date) || left.start_time.localeCompare(right.start_time) || left.title.localeCompare(right.title) || left.id.localeCompare(right.id)),
      };
    }
    if (sql.includes("FROM account_hospital_locations")) {
      return {
        results: [...this.db.accountHospitalLocations.values()]
          .filter((row) => row.email === args[0])
          .sort((left, right) => left.source_type.localeCompare(right.source_type)),
      };
    }
    if (sql.includes("FROM roster_events") && sql.includes("GROUP BY file_id, doctor_key")) {
      const pairs = new Set();
      for (let index = 0; index < args.length; index += 2) pairs.add(`${args[index]}:${args[index + 1]}`);
      const grouped = new Map();
      for (const event of this.db.events.values()) {
        const key = `${event.file_id}:${event.doctor_key}`;
        if (!pairs.has(key)) continue;
        grouped.set(key, (grouped.get(key) || 0) + 1);
      }
      return {
        results: [...grouped.entries()].map(([key, count]) => {
          const [file_id, doctor_key] = key.split(":");
          return { file_id, doctor_key, count };
        }),
      };
    }
    if (sql.includes("FROM roster_events") && sql.includes("GROUP BY file_id")) {
      const ids = new Set(args);
      const grouped = new Map();
      for (const event of this.db.events.values()) {
        if (!ids.has(event.file_id)) continue;
        grouped.set(event.file_id, (grouped.get(event.file_id) || 0) + 1);
      }
      return { results: [...grouped.entries()].map(([file_id, count]) => ({ file_id, count })) };
    }
    if (sql.includes("FROM roster_file_doctors") && sql.includes("GROUP BY file_id")) {
      const ids = new Set(args);
      const grouped = new Map();
      for (const doctor of this.db.fileDoctors.values()) {
        if (!ids.has(doctor.file_id)) continue;
        grouped.set(doctor.file_id, (grouped.get(doctor.file_id) || 0) + 1);
      }
      return { results: [...grouped.entries()].map(([file_id, count]) => ({ file_id, count })) };
    }
    if (sql.includes("FROM roster_events") && sql.includes("(roster_events.file_id = ? AND roster_events.doctor_key = ?)")) {
      const end = args[args.length - 2];
      const start = args[args.length - 1];
      const pairs = new Set();
      for (let index = 0; index < args.length - 2; index += 2) {
        pairs.add(`${args[index]}:${args[index + 1]}`);
      }
      return {
        results: [...this.db.events.values()]
          .filter((event) => this.db.files.get(event.file_id)?.active === 1)
          .filter((event) => pairs.has(`${event.file_id}:${event.doctor_key}`))
          .filter((event) => event.start_date <= end && event.end_date >= start)
          .sort((left, right) => left.start_ts.localeCompare(right.start_ts))
          .map((event) => ({ event_json: event.event_json })),
      };
    }
    if (sql.includes("FROM roster_issues") && sql.includes("(roster_issues.file_id = ? AND roster_issues.doctor_key = ?)")) {
      const end = args[args.length - 2];
      const start = args[args.length - 1];
      const pairs = new Set();
      for (let index = 0; index < args.length - 2; index += 2) {
        pairs.add(`${args[index]}:${args[index + 1]}`);
      }
      return {
        results: [...this.db.issues.values()]
          .filter((issue) => this.db.files.get(issue.file_id)?.active === 1)
          .filter((issue) => pairs.has(`${issue.file_id}:${issue.doctor_key}`))
          .filter((issue) => issue.start_date <= end && issue.start_date >= start)
          .sort((left, right) => left.start_date.localeCompare(right.start_date))
          .map((issue) => ({ issue_json: issue.issue_json })),
      };
    }
    if (sql.includes("FROM roster_daily_presence AS mine")) {
      const overlapKeyCount = (sql.match(/mine\.doctor_key IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1;
      const overlapKeys = new Set(args.slice(0, overlapKeyCount));
      const start = args[overlapKeyCount];
      const end = args[overlapKeyCount + 1];
      const sourceOffset = overlapKeyCount + 4;
      const hasSourceFilter = sql.includes("p.source_type IN");
      const hasDoctorFilter = sql.includes("p.doctor_key IN");
      const hasExcludedDoctorFilter = sql.includes("p.doctor_key NOT IN");
      const sourceCount = hasSourceFilter ? (sql.match(/p\.source_type IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const doctorCount = hasDoctorFilter ? (sql.match(/p\.doctor_key IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const excludedDoctorCount = hasExcludedDoctorFilter ? (sql.match(/p\.doctor_key NOT IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const sourceTypes = new Set(args.slice(sourceOffset, sourceOffset + sourceCount));
      const doctorKeys = new Set(args.slice(sourceOffset + sourceCount, sourceOffset + sourceCount + doctorCount));
      const excludedDoctorKeys = new Set(args.slice(sourceOffset + sourceCount + doctorCount, sourceOffset + sourceCount + doctorCount + excludedDoctorCount));
      const presence = [...this.db.dailyPresence.values()];
      const minePresence = presence.filter((row) => overlapKeys.has(row.doctor_key) && row.date >= start && row.date <= end);
      const resultPresence = presence
        .filter((row) => row.date >= start && row.date <= end)
        .filter((row) => minePresence.some((mine) => mine.date === row.date && mine.source_type === row.source_type))
        .filter((row) => !sourceTypes.size || sourceTypes.has(row.source_type))
        .filter((row) => !doctorKeys.size || doctorKeys.has(row.doctor_key))
        .filter((row) => !excludedDoctorKeys.has(row.doctor_key))
        .filter((row, index, list) => list.findIndex((item) => item.event_id === row.event_id && item.doctor_key === row.doctor_key && item.source_type === row.source_type) === index);
      const results = resultPresence
        .map((row) => ({ presence: row, event: this.db.events.get(row.event_id) }))
        .filter(({ event }) => event && this.db.files.get(event.file_id)?.active === 1)
        .sort((left, right) => left.presence.display_name.localeCompare(right.presence.display_name) || left.event.start_ts.localeCompare(right.event.start_ts));
      if (!sql.includes("event_json")) {
        return {
          results: results
            .map(({ presence }) => ({ doctor_key: presence.doctor_key, display_name: presence.display_name, source_type: presence.source_type }))
            .filter((event, index, list) => list.findIndex((item) => item.doctor_key === event.doctor_key && item.source_type === event.source_type) === index),
        };
      }
      return {
        results: results.map(({ presence, event }) => ({
          doctor_key: presence.doctor_key,
          display_name: presence.display_name,
          source_type: presence.source_type,
          event_json: event.event_json,
          start_ts: event.start_ts,
        })),
      };
    }
    if (sql.includes("FROM roster_daily_presence AS p") && sql.includes("INNER JOIN roster_events AS ev")) {
      const start = args[0];
      const end = args[1];
      let offset = 2;
      const sourceCount = sql.includes("p.source_type IN") ? (sql.match(/p\.source_type IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const doctorCount = sql.includes("p.doctor_key IN") ? (sql.match(/p\.doctor_key IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const excludedDoctorCount = sql.includes("p.doctor_key NOT IN") ? (sql.match(/p\.doctor_key NOT IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const sourceTypes = new Set(args.slice(offset, offset + sourceCount));
      offset += sourceCount;
      const doctorKeys = new Set(args.slice(offset, offset + doctorCount));
      offset += doctorCount;
      const excludedDoctorKeys = new Set(args.slice(offset, offset + excludedDoctorCount));
      const results = [...this.db.dailyPresence.values()]
        .filter((row) => row.date >= start && row.date <= end)
        .filter((row) => !sourceTypes.size || sourceTypes.has(row.source_type))
        .filter((row) => !doctorKeys.size || doctorKeys.has(row.doctor_key))
        .filter((row) => !excludedDoctorKeys.has(row.doctor_key))
        .filter((row, index, list) => list.findIndex((item) => item.event_id === row.event_id) === index)
        .map((row) => ({ presence: row, event: this.db.events.get(row.event_id) }))
        .filter(({ event }) => event && this.db.files.get(event.file_id)?.active === 1)
        .sort((left, right) => left.presence.display_name.localeCompare(right.presence.display_name) || left.event.start_ts.localeCompare(right.event.start_ts))
        .map(({ presence, event }) => ({
          doctor_key: presence.doctor_key,
          display_name: presence.display_name,
          source_type: presence.source_type,
          event_json: event.event_json,
          start_ts: event.start_ts,
        }));
      return { results };
    }
    if (sql.includes("FROM roster_events AS mine")) {
      const overlapKeyCount = (sql.match(/mine\.doctor_key IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1;
      const overlapKeys = new Set(args.slice(0, overlapKeyCount));
      const end = args[overlapKeyCount];
      const start = args[overlapKeyCount + 1];
      const sourceOffset = overlapKeyCount + 4;
      const hasSourceFilter = sql.includes("other_events.source_type IN");
      const hasDoctorFilter = sql.includes("other_events.doctor_key IN");
      const hasExcludedDoctorFilter = sql.includes("other_events.doctor_key NOT IN");
      const sourceCount = hasSourceFilter ? (sql.match(/other_events\.source_type IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const doctorCount = hasDoctorFilter ? (sql.match(/other_events\.doctor_key IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const excludedDoctorCount = hasExcludedDoctorFilter ? (sql.match(/other_events\.doctor_key NOT IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const sourceTypes = new Set(args.slice(sourceOffset, sourceOffset + sourceCount));
      const doctorKeys = new Set(args.slice(sourceOffset + sourceCount, sourceOffset + sourceCount + doctorCount));
      const excludedDoctorKeys = new Set(args.slice(sourceOffset + sourceCount + doctorCount, sourceOffset + sourceCount + doctorCount + excludedDoctorCount));
      const myEvents = [...this.db.events.values()]
        .filter((event) => this.db.files.get(event.file_id)?.active === 1)
        .filter((event) => overlapKeys.has(event.doctor_key))
        .filter((event) => event.start_date <= end && event.end_date >= start);
      const results = [...this.db.events.values()]
          .filter((event) => this.db.files.get(event.file_id)?.active === 1)
          .filter((event) => event.start_date <= end && event.end_date >= start)
          .filter((event) => myEvents.some((mine) => mine.source_type === event.source_type && event.start_date <= mine.end_date && event.end_date >= mine.start_date))
          .filter((event) => !sourceTypes.size || sourceTypes.has(event.source_type))
          .filter((event) => !doctorKeys.size || doctorKeys.has(event.doctor_key))
          .filter((event) => !excludedDoctorKeys.has(event.doctor_key))
          .filter((event, index, events) => events.findIndex((item) => item.id === event.id) === index)
          .sort((left, right) => left.display_name.localeCompare(right.display_name) || left.start_ts.localeCompare(right.start_ts));
      if (!sql.includes("event_json")) {
        return {
          results: results
            .map((event) => ({ doctor_key: event.doctor_key, display_name: event.display_name, source_type: event.source_type }))
            .filter((event, index, events) => events.findIndex((item) => item.doctor_key === event.doctor_key && item.source_type === event.source_type) === index),
        };
      }
      return { results };
    }
    if (sql.includes("SELECT DISTINCT roster_events.seniority AS seniority")) {
      const keys = new Set(args);
      return {
        results: [...new Set([...this.db.events.values()]
          .filter((event) => this.db.files.get(event.file_id)?.active === 1)
          .filter((event) => keys.has(event.doctor_key))
          .map((event) => event.seniority)
          .filter(Boolean))]
          .sort()
          .map((seniority) => ({ seniority })),
      };
    }
    if (sql.includes("SELECT id FROM roster_files WHERE active = 1")) {
      const limit = Number(args[0] || 10);
      const offset = Number(args[1] || 0);
      return {
        results: [...this.db.files.values()]
          .filter((file) => file.active === 1)
          .sort((left, right) => String(left.added_at || "").localeCompare(String(right.added_at || "")) || String(left.name || "").localeCompare(String(right.name || "")) || String(left.id || "").localeCompare(String(right.id || "")))
          .slice(offset, offset + limit)
          .map((file) => ({ id: file.id })),
      };
    }
    if (sql.includes("SELECT") && sql.includes("roster_events.id AS id") && sql.includes("WHERE roster_events.file_id = ?")) {
      return {
        results: [...this.db.events.values()]
          .filter((event) => event.file_id === args[0])
          .filter((event) => this.db.files.get(event.file_id)?.active === 1)
          .map((event) => ({
            id: event.id,
            source_type: event.source_type,
            doctor_key: event.doctor_key,
            display_name: event.display_name,
            start_date: event.start_date,
            end_date: event.end_date,
            start_ts: event.start_ts,
            end_ts: event.end_ts,
          })),
      };
    }
    if (sql.includes("FROM roster_events") && sql.includes("doctor_key IN")) {
      const end = args[args.length - 2];
      const start = args[args.length - 1];
      const keys = new Set(args.slice(0, -2));
      return {
        results: [...this.db.events.values()]
          .filter((event) => this.db.files.get(event.file_id)?.active === 1)
          .filter((event) => keys.has(event.doctor_key))
          .filter((event) => event.start_date <= end && event.end_date >= start)
          .sort((left, right) => left.start_ts.localeCompare(right.start_ts))
          .map((event) => ({ event_json: event.event_json })),
      };
    }
    if (sql.includes("FROM roster_issues") && sql.includes("doctor_key IN")) {
      const end = args[args.length - 2];
      const start = args[args.length - 1];
      const keys = new Set(args.slice(0, -2));
      return {
        results: [...this.db.issues.values()]
          .filter((issue) => this.db.files.get(issue.file_id)?.active === 1)
          .filter((issue) => keys.has(issue.doctor_key))
          .filter((issue) => issue.start_date <= end && issue.start_date >= start)
          .sort((left, right) => left.start_date.localeCompare(right.start_date))
          .map((issue) => ({ issue_json: issue.issue_json })),
      };
    }
    if (sql.includes("FROM roster_events") && sql.includes("display_name")) {
      const end = args[0];
      const start = args[1];
      const hasSourceFilter = sql.includes("roster_events.source_type IN");
      const hasDoctorFilter = sql.includes("roster_events.doctor_key IN");
      const hasExcludedDoctorFilter = sql.includes("roster_events.doctor_key NOT IN");
      const sourceCount = hasSourceFilter ? (sql.match(/roster_events\.source_type IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const doctorCount = hasDoctorFilter ? (sql.match(/roster_events\.doctor_key IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const excludedDoctorCount = hasExcludedDoctorFilter ? (sql.match(/roster_events\.doctor_key NOT IN \(([^)]*)\)/)?.[1].split("?").length || 0) - 1 : 0;
      const sourceTypes = new Set(args.slice(2, 2 + sourceCount));
      const doctorKeys = new Set(args.slice(2 + sourceCount, 2 + sourceCount + doctorCount));
      const excludedDoctorKeys = new Set(args.slice(2 + sourceCount + doctorCount, 2 + sourceCount + doctorCount + excludedDoctorCount));
      return {
        results: [...this.db.events.values()]
          .filter((event) => this.db.files.get(event.file_id)?.active === 1)
          .filter((event) => event.start_date <= end && event.end_date >= start)
          .filter((event) => !sourceTypes.size || sourceTypes.has(event.source_type))
          .filter((event) => !doctorKeys.size || doctorKeys.has(event.doctor_key))
          .filter((event) => !excludedDoctorKeys.has(event.doctor_key))
          .sort((left, right) => left.display_name.localeCompare(right.display_name) || left.start_ts.localeCompare(right.start_ts)),
      };
    }
    if (sql.includes("FROM facility_staff_seniority_overrides")) {
      return { results: [] };
    }
    if (sql.includes("FROM roster_file_doctors") && sql.includes("roster_file_doctors.file_id AS file_id")) {
      const hasKeyFilter = sql.includes("roster_file_doctors.doctor_key IN");
      const keys = new Set(args);
      return {
        results: [...this.db.fileDoctors.values()]
          .filter((doctor) => this.db.files.get(doctor.file_id)?.active === 1)
          .filter((doctor) => !hasKeyFilter || keys.has(doctor.doctor_key))
          .map((doctor) => {
            const file = this.db.files.get(doctor.file_id);
            return {
              file_id: doctor.file_id,
              file_name: file?.name || "",
              file_source_type: file?.source_type || doctor.source_type,
              active: file?.active ?? 0,
              source_type: doctor.source_type,
              doctor_key: doctor.doctor_key,
              display_name: doctor.display_name,
              event_count: [...this.db.events.values()].filter((event) => event.file_id === doctor.file_id && event.doctor_key === doctor.doctor_key).length,
            };
          })
          .sort((left, right) => String(left.file_name).localeCompare(String(right.file_name)) || left.display_name.localeCompare(right.display_name)),
      };
    }
    if (sql.includes("FROM roster_file_doctors") && sql.includes("DISTINCT")) {
      return {
        results: [...this.db.fileDoctors.values()]
          .filter((doctor) => this.db.files.get(doctor.file_id)?.active === 1)
          .map((doctor) => ({
            source_type: doctor.source_type,
            doctor_key: doctor.doctor_key,
            display_name: doctor.display_name,
          }))
          .filter((doctor, index, doctors) => doctors.findIndex((item) => item.source_type === doctor.source_type && item.doctor_key === doctor.doctor_key) === index)
          .sort((left, right) => left.display_name.localeCompare(right.display_name) || left.source_type.localeCompare(right.source_type)),
      };
    }
    if (sql.includes("FROM roster_files") && sql.includes("doctor_count") && sql.includes("event_count")) {
      const includeInactive = !sql.includes("WHERE roster_files.active = 1");
      return {
        results: [...this.db.files.values()]
          .filter((file) => includeInactive || file.active === 1)
          .sort((left, right) => String(left.added_at || "").localeCompare(String(right.added_at || "")) || left.name.localeCompare(right.name))
          .map((file) => ({
            id: file.id,
            name: file.name,
            source_type: file.source_type,
            source_id: file.source_id,
            active: file.active,
            size: file.size,
            last_modified: file.last_modified,
            added_at: file.added_at,
            uploaded_at: file.uploaded_at,
            uploaded_by: file.uploaded_by,
            doctor_count: [...this.db.fileDoctors.values()].filter((doctor) => doctor.file_id === file.id).length,
            event_count: [...this.db.events.values()].filter((event) => event.file_id === file.id).length,
          })),
      };
    }
    if (sql.includes("FROM raw_roster_files") && !sql.includes("WHERE file_id = ?")) {
      return {
        results: [...this.db.rawFiles.values()].sort((left, right) => String(left.uploaded_at || "").localeCompare(String(right.uploaded_at || "")) || String(left.file_id || "").localeCompare(String(right.file_id || ""))),
      };
    }
    if (sql.includes("MIN(roster_events.start_date)") && sql.includes("FROM roster_files")) {
      const includeInactive = !sql.includes("WHERE roster_files.active = 1");
      return {
        results: [...this.db.files.values()]
          .filter((file) => includeInactive || file.active === 1)
          .map((file) => {
            const events = [...this.db.events.values()].filter((event) => event.file_id === file.id);
            const starts = events.map((event) => event.start_date).filter(Boolean).sort();
            const ends = events.map((event) => event.end_date).filter(Boolean).sort();
            return {
            id: file.id,
            name: file.name,
            source_type: file.source_type,
            source_id: file.source_id,
              active: file.active,
              last_modified: file.last_modified,
              added_at: file.added_at,
              uploaded_at: file.uploaded_at,
              start_date: starts[0] || "",
              coverage_end_date: starts[starts.length - 1] || "",
              end_date: ends[ends.length - 1] || "",
              event_count: events.length,
            };
          })
          .sort((left, right) => String(left.source_type).localeCompare(String(right.source_type)) || String(left.start_date).localeCompare(String(right.start_date))),
      };
    }
    if (sql.includes("FROM roster_files") && sql.includes("INNER JOIN roster_file_doctors") && sql.includes("roster_file_doctors.doctor_key IN")) {
      const keys = new Set(args);
      const seen = new Set();
      const results = [];
      for (const doctor of [...this.db.fileDoctors.values()]) {
        const file = this.db.files.get(doctor.file_id);
        if (!file || file.active !== 1 || !keys.has(doctor.doctor_key) || seen.has(file.id)) continue;
        seen.add(file.id);
        results.push({
          id: file.id,
          name: file.name,
          source_type: file.source_type,
          active: file.active,
          size: file.size,
          last_modified: file.last_modified,
          added_at: file.added_at,
          uploaded_at: file.uploaded_at,
          uploaded_by: file.uploaded_by,
        });
      }
      return {
        results: results.sort((left, right) => String(left.added_at || "").localeCompare(String(right.added_at || "")) || left.name.localeCompare(right.name)),
      };
    }
    if (sql.startsWith("SELECT file_id, source_type, doctor_key, display_name FROM roster_file_doctors")) {
      return {
        results: [...this.db.fileDoctors.values()]
          .sort((left, right) => left.display_name.localeCompare(right.display_name) || left.source_type.localeCompare(right.source_type)),
      };
    }
    if (sql.includes("FROM account_profiles") && sql.includes("LEFT JOIN account_claims")) {
      const results = [];
      const tokenFilter = sql.includes("WHERE account_profiles.subscription_token = ?");
      const emailFilter = sql.includes("WHERE account_profiles.email = ?");
      for (const profile of [...this.db.accountProfiles.values()].sort((left, right) => left.email.localeCompare(right.email))) {
        if (tokenFilter && profile.subscription_token !== args[0]) continue;
        if (emailFilter && profile.email !== args[0]) continue;
        const state = this.db.accountStates.get(profile.email) || null;
        const claims = [...this.db.accountClaims.values()]
          .filter((claim) => claim.email === profile.email)
          .sort((left, right) => left.source_type.localeCompare(right.source_type) || left.display_name.localeCompare(right.display_name));
        if (!claims.length) {
          results.push({
            email: profile.email,
            real_name: profile.real_name,
            role: profile.role,
            insights_enabled: profile.insights_enabled,
            facility_overview_enabled: profile.facility_overview_enabled,
            non_clinical: profile.non_clinical,
            director_view_enabled: profile.director_view_enabled,
            subscription_token: profile.subscription_token,
            password_salt: profile.password_salt,
            password_hash: profile.password_hash,
            admin_issues_json: profile.admin_issues_json,
            local_parser_extensions_json: profile.local_parser_extensions_json,
            created_at: profile.created_at,
            updated_at: profile.updated_at,
            source_type: null,
            doctor_key: null,
            display_name: null,
            matched_at: null,
            session_json: state?.session_json || null,
          });
          continue;
        }
        for (const claim of claims) {
          results.push({
            email: profile.email,
            real_name: profile.real_name,
            role: profile.role,
            insights_enabled: profile.insights_enabled,
            facility_overview_enabled: profile.facility_overview_enabled,
            non_clinical: profile.non_clinical,
            director_view_enabled: profile.director_view_enabled,
            subscription_token: profile.subscription_token,
            password_salt: profile.password_salt,
            password_hash: profile.password_hash,
            admin_issues_json: profile.admin_issues_json,
            local_parser_extensions_json: profile.local_parser_extensions_json,
            created_at: profile.created_at,
            updated_at: profile.updated_at,
            source_type: claim.source_type,
            doctor_key: claim.doctor_key,
            display_name: claim.display_name,
            matched_at: claim.matched_at,
            session_json: state?.session_json || null,
          });
        }
      }
      return { results };
    }
    if (sql.startsWith("SELECT rule_json FROM parser_rules")) {
      return {
        results: [...this.db.parserRules.values()]
          .filter((rule) => rule.scope === "global")
          .map((rule) => ({ rule_json: rule.rule_json })),
      };
    }
    if (sql.startsWith("SELECT suggestion_json FROM parser_rule_suggestions")) {
      return {
        results: [...this.db.parserRuleSuggestions.values()]
          .filter((suggestion) => suggestion.status === "pending")
          .sort((left, right) => right.updated_at.localeCompare(left.updated_at))
          .map((suggestion) => ({ suggestion_json: suggestion.suggestion_json })),
      };
    }
    if (sql.startsWith("SELECT token_hash, email FROM account_invites")) {
      return { results: [] };
    }
    if (sql.startsWith("SELECT * FROM doctor_profiles ORDER BY")) {
      return {
        results: [...this.db.doctorProfiles.values()]
          .sort((left, right) => left.display_name.localeCompare(right.display_name) || left.profile_id.localeCompare(right.profile_id)),
      };
    }
    if (sql.includes("FROM snapshot_registry") && sql.includes("WHERE owner_type = ? AND owner_id = ?")) {
      return {
        results: [...this.db.snapshotRegistry.values()]
          .filter((row) => row.owner_type === args[0] && row.owner_id === args[1])
          .sort((left, right) => String(right.updated_at || "").localeCompare(String(left.updated_at || ""))),
      };
    }
    if (sql.includes("FROM snapshot_registry") && sql.includes("owner_type IN")) {
      const ownerTypeCount = (sql.match(/owner_type IN \(([^)]*)\)/)?.[1].match(/\?/g) || []).length;
      const statusCount = (sql.match(/status IN \(([^)]*)\)/)?.[1].match(/\?/g) || []).length;
      const ownerTypes = new Set(args.slice(0, ownerTypeCount));
      const statuses = new Set(args.slice(ownerTypeCount, ownerTypeCount + statusCount));
      let offset = ownerTypeCount + statusCount;
      const hasRange = sql.includes("AND range_key = ?");
      const rangeKey = hasRange ? args[offset++] : "";
      const limit = Number(args[offset] || 25);
      return {
        results: [...this.db.snapshotRegistry.values()]
          .filter((row) => ownerTypes.has(row.owner_type))
          .filter((row) => statuses.has(row.status))
          .filter((row) => !hasRange || row.range_key === rangeKey)
          .sort((left, right) => String(right.updated_at || "").localeCompare(String(left.updated_at || "")))
          .slice(0, limit),
      };
    }
    if (sql.includes("FROM console_messages")) {
      return {
        results: [...this.db.consoleMessages]
          .sort((left, right) => String(right.created_at).localeCompare(String(left.created_at)) || right.id - left.id)
          .slice(0, Number(args[0] || 50)),
      };
    }
    throw new Error(`Unsupported MemoryD1 all SQL: ${sql}`);
  }

  async first() {
    const sql = this.sql;
    const args = this.args;
    if (sql.startsWith("SELECT person_id FROM roster_person_aliases")) {
      return this.db.rosterPersonAliases.get(`${args[0]}|${args[1]}`) || null;
    }
    if (sql.startsWith("SELECT * FROM roster_sources WHERE id = ?")) {
      return this.db.rosterSources.get(args[0]) || null;
    }
    if (sql.startsWith("SELECT id FROM roster_sync_runs WHERE status IN")) {
      return [...this.db.rosterSyncRuns.values()].find((row) => ["queued", "processing"].includes(row.status)) || null;
    }
    if (sql.startsWith("SELECT * FROM roster_dispatches WHERE status IN")) {
      return [...this.db.rosterDispatches.values()]
        .filter((row) => ["requested", "accepted", "running", "failed"].includes(row.status) && String(row.retry_after || "") > String(args[0] || ""))
        .sort((left, right) => String(right.requested_at).localeCompare(String(left.requested_at)))[0] || null;
    }
    if (sql.startsWith("SELECT * FROM roster_dispatches WHERE id = ?")) {
      return this.db.rosterDispatches.get(args[0]) || null;
    }
    if (sql.startsWith("SELECT * FROM roster_dispatches ORDER BY")) {
      return [...this.db.rosterDispatches.values()].sort((left, right) => String(right.requested_at).localeCompare(String(left.requested_at)))[0] || null;
    }
    if (sql.includes("FROM roster_sync_runs") && sql.includes("WHERE roster_sync_runs.source_id = ?") && sql.includes("provider_version = ?")) {
      const statusOrder = new Map([["success", 0], ["processing", 1], ["queued", 2], ["failed", 3]]);
      return [...this.db.rosterSyncRuns.values()]
        .filter((row) => row.source_id === args[0] && row.provider_version === args[1])
        .filter((row) => String(this.db.rawFiles.get(row.source_file_id || row.file_id)?.name || "").toLowerCase() === String(args[2] || "").toLowerCase())
        .filter((row) => statusOrder.has(row.status))
        .sort((left, right) => (statusOrder.get(left.status) - statusOrder.get(right.status))
          || String(right.started_at).localeCompare(String(left.started_at)))[0] || null;
    }
    if (sql.includes("FROM roster_sync_runs") && sql.includes("WHERE source_id = ?")) {
      let rows = [...this.db.rosterSyncRuns.values()]
        .filter((row) => row.source_id === args[0] && row.content_hash === args[1]);
      if (sql.includes("status = 'success'")) rows = rows.filter((row) => row.status === "success");
      if (sql.includes("status IN ('queued', 'processing')")) rows = rows.filter((row) => ["queued", "processing"].includes(row.status));
      return rows.sort((left, right) => String(right.completed_at || right.started_at).localeCompare(String(left.completed_at || left.started_at)))[0] || null;
    }
    if (sql.includes("COUNT(*) AS active_file_count") && sql.includes("FROM roster_files")) {
      const activeFiles = [...this.db.files.values()].filter((file) => file.active === 1);
      return {
        active_file_count: activeFiles.length,
        max_parsed_at: activeFiles.map((file) => String(file.parsed_at || "")).sort().at(-1) || "",
        max_uploaded_at: activeFiles.map((file) => String(file.uploaded_at || "")).sort().at(-1) || "",
        max_last_modified: Math.max(0, ...activeFiles.map((file) => Number(file.last_modified || 0))),
      };
    }
    if (sql.startsWith("SELECT id, source_type, parsed_at FROM roster_files WHERE id = ? AND active = 1")) {
      const file = this.db.files.get(args[0]);
      return file?.active === 1 ? { id: file.id, source_type: file.source_type, parsed_at: file.parsed_at } : null;
    }
    if (sql.startsWith("SELECT MIN(start_date) AS start_date, MAX(start_date) AS end_date FROM roster_events WHERE file_id = ?")) {
      const rows = [...this.db.events.values()].filter((event) => event.file_id === args[0]);
      return { start_date: rows.map((event) => event.start_date).sort()[0] || null, end_date: rows.map((event) => event.start_date).sort().at(-1) || null };
    }
    if (sql.includes("FROM facility_sms_memberships") && sql.includes("WHERE source_type = ? AND doctor_key = ?")) {
      return this.db.facilitySmsMemberships.get(`${args[0]}|${args[1]}`) || null;
    }
    if (sql.includes("FROM facility_staff_designations") && sql.includes("COUNT(*) AS count")) {
      const rows = [...this.db.facilityStaffDesignations.values()];
      return { count: rows.length, max_updated_at: rows.map((row) => String(row.updated_at || "")).sort().at(-1) || "" };
    }
    if (sql.includes("FROM facility_staff_seniority_overrides") && sql.includes("COUNT(*) AS count")) {
      return { count: 0, max_updated_at: "" };
    }
    if (sql.includes("FROM facility_sms_memberships") && sql.includes("COUNT(*) AS count")) {
      const rows = [...this.db.facilitySmsMemberships.values()];
      return { count: rows.length, max_updated_at: rows.map((row) => String(row.updated_at || "")).sort().at(-1) || "" };
    }
    if (sql.includes("FROM custom_events") && sql.includes("COUNT(*) AS count")) {
      const rows = [...this.db.customEvents.values()].filter((event) => event.owner_email === args[0]);
      return {
        count: rows.length,
        max_updated_at: rows.map((row) => String(row.updated_at || "")).sort().at(-1) || "",
      };
    }
    if (sql.includes("FROM account_claims") && sql.includes("COUNT(*) AS count") && sql.includes("WHERE email = ?")) {
      const rows = [...this.db.accountClaims.values()].filter((claim) => claim.email === args[0]);
      return {
        count: rows.length,
        max_updated_at: rows.map((row) => String(row.updated_at || "")).sort().at(-1) || "",
      };
    }
    if (sql.includes("FROM account_hospital_locations") && sql.includes("COUNT(*) AS count") && sql.includes("WHERE email = ?")) {
      const rows = [...this.db.accountHospitalLocations.values()].filter((row) => row.email === args[0]);
      return {
        count: rows.length,
        max_updated_at: rows.map((row) => String(row.updated_at || "")).sort().at(-1) || "",
      };
    }
    if (sql.includes("FROM doctor_profiles") && sql.includes("INNER JOIN account_claims")) {
      const claimKeys = new Set([...this.db.accountClaims.values()].filter((claim) => claim.email === args[0]).map((claim) => claim.doctor_key));
      const rows = [...this.db.doctorProfiles.values()].filter((profile) => claimKeys.has(profile.doctor_key));
      return {
        count: new Set(rows.map((row) => row.profile_id)).size,
        max_updated_at: rows.map((row) => String(row.updated_at || "")).sort().at(-1) || "",
      };
    }
    if (sql.includes("FROM parser_rules") && sql.includes("COUNT(*) AS count")) {
      const rows = [...this.db.parserRules.values()].filter((rule) => rule.scope === "global");
      return {
        count: rows.length,
        max_updated_at: rows.map((row) => String(row.updated_at || "")).sort().at(-1) || "",
      };
    }
    if (sql.includes("account_profiles.local_parser_extensions_json") && sql.includes("LEFT JOIN account_states")) {
      const profile = this.db.accountProfiles.get(args[0]) || null;
      const state = this.db.accountStates.get(args[0]) || null;
      if (!profile && !state) return null;
      return {
        session_json: state?.session_json || null,
        local_parser_extensions_json: profile?.local_parser_extensions_json || "[]",
      };
    }
    if (sql.startsWith("SELECT COUNT(*) AS count FROM roster_events WHERE file_id = ? AND doctor_key = ?")) {
      return {
        count: [...this.db.events.values()].filter((event) => event.file_id === args[0] && event.doctor_key === args[1]).length,
      };
    }
    if (sql.startsWith("SELECT COUNT(*) AS count FROM roster_events WHERE file_id")) {
      return {
        count: [...this.db.events.values()].filter((event) => event.file_id === args[0]).length,
      };
    }
    if (sql.startsWith("SELECT COUNT(*) AS count FROM roster_file_doctors WHERE file_id")) {
      return {
        count: [...this.db.fileDoctors.values()].filter((doctor) => doctor.file_id === args[0]).length,
      };
    }
    if (sql.startsWith("SELECT session_json FROM account_states WHERE email")) {
      return this.db.accountStates.get(args[0]) || null;
    }
    if (sql.startsWith("SELECT email FROM subscription_tokens WHERE token")) {
      return this.db.subscriptionTokens.get(args[0]) || null;
    }
    if (sql.startsWith("SELECT COUNT(*) AS count FROM account_profiles WHERE subscription_token")) {
      return { count: [...this.db.accountProfiles.values()].filter((profile) => profile.subscription_token).length };
    }
    if (sql.startsWith("SELECT COUNT(*) AS count FROM account_profiles")) {
      return { count: this.db.accountProfiles.size };
    }
    if (sql.startsWith("SELECT COUNT(*) AS count FROM account_claims")) {
      return { count: this.db.accountClaims.size };
    }
    if (sql.startsWith("SELECT COUNT(*) AS count FROM account_states")) {
      return { count: this.db.accountStates.size };
    }
    if (sql.startsWith("SELECT COUNT(*) AS count FROM doctor_profiles")) {
      return { count: this.db.doctorProfiles.size };
    }
    if (sql.startsWith("SELECT * FROM doctor_profiles WHERE profile_id")) {
      return this.db.doctorProfiles.get(args[0]) || null;
    }
    if (sql.includes("FROM snapshot_registry") && sql.includes("WHERE owner_type = ? AND owner_id = ? AND doctor_key = ? AND range_key = ?")) {
      return this.db.snapshotRegistry.get(`${args[0]}|${args[1]}|${args[2]}|${args[3]}`) || null;
    }
    if (sql.includes("FROM raw_roster_files") && sql.includes("WHERE file_id = ?")) {
      return this.db.rawFiles.get(args[0]) || null;
    }
    if (sql.startsWith("SELECT id FROM roster_files WHERE id = ?")) {
      const file = this.db.files.get(args[0]);
      return file ? { id: file.id } : null;
    }
    if (sql.startsWith("SELECT source_type FROM roster_files WHERE id = ?")) {
      const file = this.db.files.get(args[0]);
      return file ? { source_type: file.source_type } : null;
    }
    throw new Error(`Unsupported MemoryD1 first SQL: ${sql}`);
  }
}

function repositoryFile(id, overrides = {}) {
  return {
    repoId: id,
    id,
    name: `${id}.xlsx`,
    sourceType: "mmc",
    active: true,
    size: 12,
    lastModified: 1,
    doctors: [{
      key: "TITUS HACKMAN",
      displayName: "Titus HACKMAN",
      sourceType: "mmc",
    }],
    ...overrides,
  };
}

async function seedRepository(store, files) {
  await store.put("repository:index", JSON.stringify({ version: 1, files }));
  for (const file of files) {
    await store.put(`repository:file:${file.id}`, JSON.stringify({
      ...file,
      dataUrl: `data:application/octet-stream;base64,${Buffer.from(file.id).toString("base64")}`,
    }));
  }
}

function seedD1Repository(db, files) {
  db.canonicalDoctors.clear();
  for (const file of files) {
    db.files.set(file.id, {
      id: file.id,
      name: file.name || `${file.id}.xlsx`,
      source_type: file.sourceType || "mmc",
      active: file.active === false ? 0 : 1,
      size: file.size || 0,
      last_modified: file.lastModified || 0,
      added_at: file.addedAt || "",
      uploaded_at: file.uploadedAt || "",
      uploaded_by: file.uploadedBy || "",
      parsed_at: file.parsedAt || "",
    });
    for (const doctor of file.doctors || []) {
      db.fileDoctors.set(`${file.id}|${doctor.sourceType || file.sourceType || "mmc"}|${doctor.key}`, {
        file_id: file.id,
        source_type: doctor.sourceType || file.sourceType || "mmc",
        doctor_key: doctor.key,
        display_name: doctor.displayName || doctor.key,
      });
    }
  }
}

function seedMinimalD1DoctorEvent(db, fileId, doctorKey, sourceType, displayName = doctorKey) {
  const id = `${fileId}:${doctorKey}`;
  const event = {
    id,
    source: String(sourceType || "mmc").toUpperCase(),
    title: "Shift",
    start: "2026-01-01T09:00:00",
    end: "2026-01-01T17:00:00",
    allDay: false,
    rawValue: "Shift",
  };
  db.events.set(id, {
    id,
    file_id: fileId,
    source_type: sourceType,
    doctor_key: doctorKey,
    display_name: displayName,
    start_date: "2026-01-01",
    end_date: "2026-01-01",
    start_ts: "2026-01-01T09:00:00",
    end_ts: "2026-01-01T17:00:00",
    title: "Shift",
    raw_value: "Shift",
    seniority: "Registrar",
    location: "",
    all_day: 0,
    time_label: "0900-1700",
    event_json: JSON.stringify(event),
  });
}

async function seedUser(store, email, password, realName = "Titus Hackman", db = null) {
  await postState(store, {
    action: "login",
    email,
    password,
    mode: "create",
    realName,
  }, db);
}

async function postState(store, payload, db = null) {
  const { response, body } = await postStateRaw(store, payload, db);
  assert.equal(response.ok, true, body.error || "state request failed");
  return body;
}

async function postStateRaw(store, payload, db = null, options = {}) {
  const rosterDb = db || store?.d1 || new MemoryD1();
  const waitUntilPromises = [];
  const requestContext = {
    request: new Request("http://fixture.test/api/state", {
      method: "POST",
      body: JSON.stringify(payload),
      headers: { "content-type": "application/json" },
    }),
    env: {
      ROSTER_DB: rosterDb,
      ROSTER_FILES: store?.r2 || new MemoryR2(),
      ROSTER_CACHE: store?.cacheR2 || store?.r2 || new MemoryR2(),
    },
  };
  if (options.captureWaitUntil === true) {
    requestContext.waitUntil = (promise) => waitUntilPromises.push(Promise.resolve(promise));
  }
  const response = await handleStatePost(requestContext);
  const body = await response.json();
  return { response, body, waitUntilPromises };
}

function memoryD1AccountRecord(db, email) {
  const profile = db.accountProfiles.get(email);
  if (!profile) return null;
  return {
    email,
    realName: profile.real_name,
    role: profile.role,
    adminIssues: JSON.parse(profile.admin_issues_json || "[]"),
    localParserExtensions: JSON.parse(profile.local_parser_extensions_json || "[]"),
    claims: [...db.accountClaims.values()].filter((claim) => claim.email === email).map((claim) => ({
      sourceType: claim.source_type,
      key: claim.doctor_key,
      displayName: claim.display_name,
    })),
  };
}

const dispatchDb = new MemoryD1();
dispatchDb.rosterSyncRuns.set("queued-dispatch", {
  id: "queued-dispatch", source_id: "monash-adults", trigger_type: "sharepoint", provider_version: "1.0",
  content_hash: "dispatch-hash", file_id: "dispatch-file", status: "queued", message: "Queued", doctor_count: 0, event_count: 0,
  started_at: "2026-07-29T04:00:00.000Z", completed_at: "",
});
const originalFetch = globalThis.fetch;
const dispatchRequests = [];
globalThis.fetch = async (url, options = {}) => {
  dispatchRequests.push({ url: String(url), options });
  return new Response(null, { status: 204 });
};
const dispatched = await requestQueuedRosterProcessing({ ROSTER_DB: dispatchDb, GITHUB_ACTIONS_TOKEN: "fixture-token" }, {
  reason: "fixture", now: new Date("2026-07-29T04:01:00.000Z"),
});
assert.equal(dispatched.dispatched, true, "a queued roster should immediately dispatch GitHub processing");
assert.equal(dispatchRequests.length, 1, "a new queue should make one GitHub dispatch request");
assert.match(dispatchRequests[0].url, /actions\/workflows\/monash-roster-sync\.yml\/dispatches$/, "dispatch should target the roster workflow");
const duplicateDispatch = await requestQueuedRosterProcessing({ ROSTER_DB: dispatchDb, GITHUB_ACTIONS_TOKEN: "fixture-token" }, {
  reason: "fixture-repeat", now: new Date("2026-07-29T04:02:00.000Z"),
});
assert.equal(duplicateDispatch.dispatched, false, "an accepted dispatch lease should prevent duplicate GitHub runs");
assert.equal(dispatchRequests.length, 1, "the duplicate queue check should not call GitHub again");
const lifecycle = await recordRosterDispatchLifecycle({ ROSTER_DB: dispatchDb }, {
  dispatchId: dispatched.dispatch.id, event: "started", githubRunId: "12345",
});
assert.equal(lifecycle.dispatch.status, "running", "the workflow start callback should make dispatch state observable");
globalThis.fetch = originalFetch;

const rejectedDispatchDb = new MemoryD1();
rejectedDispatchDb.rosterSyncRuns.set("rejected-dispatch", {
  id: "rejected-dispatch", source_id: "monash-adults", trigger_type: "sharepoint", provider_version: "2.0",
  content_hash: "rejected-hash", file_id: "rejected-file", status: "queued", message: "Queued", doctor_count: 0, event_count: 0,
  started_at: "2026-07-29T04:00:00.000Z", completed_at: "",
});
let rejectedRequestCount = 0;
globalThis.fetch = async () => {
  rejectedRequestCount += 1;
  return Response.json({ message: "Resource not accessible by personal access token" }, { status: 403 });
};
const rejectedDispatch = await requestQueuedRosterProcessing({ ROSTER_DB: rejectedDispatchDb, GITHUB_ACTIONS_TOKEN: "rejected-token" }, {
  reason: "fixture-rejected", now: new Date("2026-07-29T04:01:00.000Z"),
});
assert.equal(rejectedDispatch.reason, "github-rejected", "a rejected GitHub token should be visible to the caller");
const rejectedRetry = await requestQueuedRosterProcessing({ ROSTER_DB: rejectedDispatchDb, GITHUB_ACTIONS_TOKEN: "rejected-token" }, {
  reason: "fixture-rejected-repeat", now: new Date("2026-07-29T04:02:00.000Z"),
});
assert.equal(rejectedRetry.dispatched, false, "a rejected token should honour its retry lease");
assert.equal(rejectedRequestCount, 1, "a rejected GitHub token must not be retried for every watchdog tick");
globalThis.fetch = originalFetch;

const stateStore = new MemoryStore();
stateStore.d1 = new MemoryD1();
const creatorPassword = "fixture-password";
const wranglerConfig = await readFile(new URL("../wrangler.toml", import.meta.url), "utf8");
assert.ok(/binding\s*=\s*"ROSTER_DB"/.test(wranglerConfig), "wrangler.toml must bind D1 ROSTER_DB");
assert.ok(!/ROSTER_STORE/.test(wranglerConfig), "wrangler.toml must not reintroduce KV ROSTER_STORE");
await postState(stateStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
await seedRepository(stateStore, [repositoryFile("fixture-roster", {
  name: "AdultMMCTerm2.2026.Ver1.pdf",
  sourceType: "mmc",
})]);

const creatorImports = await postStateRaw(stateStore, {
  action: "loadImports",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.equal(creatorImports.response.status, 410);
const emptyD1Status = await postState(stateStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.equal(emptyD1Status.total, 0, "KV repository metadata must not be used when D1 roster_files is empty");

const d1StateStore = new MemoryStore();
const d1Store = new MemoryD1();
const d1Doctor = doctorOptions(parsedMmcUpload.sources.mmc, [], [], [])[0];
assert.ok(d1Doctor?.key, "fixture should expose at least one MMC doctor");
await postState(d1StateStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, d1Store);
const d1Doctors = doctorOptions(parsedMmcUpload.sources.mmc, [], [], []);
const d1EventsByDoctor = Object.fromEntries(d1Doctors.map((doctor) => [
  doctor.key,
  buildRosterView(parsedMmcUpload.sources.mmc, [], doctor.key, {}, {}, {}, [], [], []).events,
]));
await postState(d1StateStore, {
  action: "uploadRawRosterFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: {
    id: "d1-mmc",
    name: "AdultTerm1.2026.xlsx",
    size: 123,
    lastModified: 1,
    addedAt: "2026-01-01T00:00:00.000Z",
    sourceType: "mmc",
  },
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  dataUrl: workbookDataUrl(mmcWorkbook),
}, d1Store);
await postState(d1StateStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: {
    id: "d1-mmc",
    name: "AdultTerm1.2026.xlsx",
    size: 123,
    lastModified: 1,
    addedAt: "2026-01-01T00:00:00.000Z",
    sourceType: "mmc",
  },
  doctors: d1Doctors,
  eventsByDoctor: d1EventsByDoctor,
}, d1Store);
const d1CaseyDoctorOptions = doctorOptions([], [], caseyWorkbook);
const d1CaseyDoctors = [];
const seenCaseyDoctorKeys = new Set();
for (const doctor of d1CaseyDoctorOptions) {
  const aliases = Array.isArray(doctor.aliases) && doctor.aliases.length
    ? doctor.aliases.filter((alias) => String(alias.sourceType || "").toLowerCase() === "casey")
    : [{ key: doctor.key, displayName: doctor.displayName, sourceType: "casey" }];
  for (const alias of aliases) {
    const key = String(alias.key || doctor.key || "").trim();
    if (!key || seenCaseyDoctorKeys.has(key)) continue;
    seenCaseyDoctorKeys.add(key);
    d1CaseyDoctors.push({
      key,
      displayName: alias.displayName || doctor.displayName,
      sourceType: "casey",
    });
  }
}
const d1CaseyEventsByDoctor = Object.fromEntries(d1CaseyDoctors.map((doctor) => [
  doctor.key,
  buildRosterView([], [], doctor.key, undefined, {}, {}, [], caseyWorkbook).events,
]));
const d1CaseyStore = new MemoryStore();
const d1CaseyDb = new MemoryD1();
await postState(d1CaseyStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, d1CaseyDb);
const caseySave = await postState(d1CaseyStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: {
    id: "d1-casey",
    name: "Casey_Term_2_2026_DRAFT.xlsm",
    size: caseyBytes.length,
    lastModified: 1,
    addedAt: "2026-01-01T00:00:00.000Z",
    sourceType: "casey",
  },
  doctors: d1CaseyDoctors,
  eventsByDoctor: d1CaseyEventsByDoctor,
  skipStatus: true,
}, d1CaseyDb);
assert.equal(caseySave.indexing, "complete");
assert.ok(Number(caseySave.fileStatus?.eventCount || 0) > 1000, "Casey roster save should persist a large event set");
assert.ok(d1CaseyDb.dailyPresence.size > 1000, "Casey roster save should index daily presence rows");
const caseyStatus = await postState(d1CaseyStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  includeAvailableDoctors: true,
}, d1CaseyDb);
assert.equal(caseyStatus.files.find((file) => file.id === "d1-casey")?.status, "populated");
assert.ok(
  (caseyStatus.availableDoctors || []).some((doctor) => doctor.displayName === "Andrew DYALL" && doctor.sourceType === "casey"),
  "Casey roster save should expose Casey-only doctors for switcher refresh",
);
const retainedRaw = await postState(d1StateStore, {
  action: "fetchRawRosterFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  fileId: "d1-mmc",
}, d1Store);
assert.match(retainedRaw.dataUrl, /^data:/, "retained raw files should be fetchable for browser-side reparsing");
const legacyRawStore = new MemoryStore();
const legacyRawDb = new MemoryD1();
await postState(legacyRawStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, legacyRawDb);
legacyRawDb.rawFiles.set("legacy-raw", {
  file_id: "legacy-raw",
  object_key: "",
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  data_url: workbookDataUrl(mmcWorkbook),
  uploaded_at: "2026-01-01T00:00:00.000Z",
});
const migratedRaw = await postState(legacyRawStore, {
  action: "fetchRawRosterFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  fileId: "legacy-raw",
}, legacyRawDb);
assert.match(migratedRaw.dataUrl, /^data:/, "legacy raw files should remain fetchable during lazy migration");
assert.equal(legacyRawDb.rawFiles.get("legacy-raw")?.object_key, "rosters/legacy-raw", "legacy raw files should be promoted to R2 when fetched");
assert.equal(legacyRawDb.rawFiles.get("legacy-raw")?.data_url, "", "lazy migration should clear the inline D1 payload after promotion");
legacyRawDb.rawFiles.set("retained-only:1:1", {
  file_id: "retained-only:1:1",
  name: "Dandenong retained only.xlsx",
  source_type: "ddh",
  size: 1,
  last_modified: 1,
  object_key: "rosters/retained-only:1:1",
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  data_url: "",
  uploaded_at: "2026-01-01T00:01:00.000Z",
});
const retainedOnlyStatus = await postState(legacyRawStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, legacyRawDb);
const retainedOnlyFile = retainedOnlyStatus.files.find((file) => file.id === "retained-only:1:1");
assert.equal(retainedOnlyFile?.retainedSourceOnly, true, "calendar status should list retained R2 source files without derived rows");
assert.equal(retainedOnlyFile?.status, "retained");
const reparsedD1Status = await postState(d1StateStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  selectedDoctorKey: d1Doctor.key,
}, d1Store);
assert.equal(reparsedD1Status.files.find((file) => file.id === "d1-mmc")?.rawSourceAvailable, true, "durably stored raw files should be visible in status");

const missingRawDb = new MemoryD1();
await postState(new MemoryStore(), {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, missingRawDb);
await postState(new MemoryStore(), {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "missing-raw", name: "missing-raw.xlsx", sourceType: "mmc", active: true },
  doctors: [{ key: "RICHARD HAYDON", displayName: "Richard HAYDON", sourceType: "mmc" }],
  eventsByDoctor: {
    "RICHARD HAYDON": [{ id: "missing-raw-shift", source: "MMC", title: "MMC shift", allDay: true, start: "2026-02-03", end: "2026-02-03", rawValue: "MMC shift" }],
  },
}, missingRawDb);
const missingRawReparse = await postStateRaw(new MemoryStore(), {
  action: "fetchRawRosterFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  fileId: "missing-raw",
}, missingRawDb);
assert.equal(missingRawReparse.response.status, 404, "files without retained raw bytes should fail source fetch explicitly");
await postState(d1StateStore, {
  action: "save",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  state: {
    version: 1,
    imports: [{
      id: "d1-mmc",
      name: "AdultTerm1.2026.xlsx",
      size: 123,
      lastModified: 1,
      addedAt: "2026-01-01T00:00:00.000Z",
      sourceType: "mmc",
    }],
    session: {
      doctorKey: d1Doctor.key,
      settings: {},
    },
  },
}, d1Store);
assert.ok(d1Store.events.size > 0, "D1 should contain derived roster events after creator save");
const transactionalDb = new MemoryD1();
const transactionalStore = new MemoryStore();
await postState(transactionalStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, transactionalDb);
await postState(transactionalStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "safe-refresh", name: "safe-refresh.xlsx", sourceType: "mmc" },
  doctors: [{ key: "RICHARD HAYDON", displayName: "Richard HAYDON", sourceType: "mmc" }],
  eventsByDoctor: {
    "RICHARD HAYDON": [{ id: "original", source: "MMC", title: "Original", allDay: true, start: "2026-02-03", end: "2026-02-03", rawValue: "Original" }],
  },
}, transactionalDb);
transactionalDb.failNextEventInsert = true;
const failedRefresh = await postStateRaw(transactionalStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "safe-refresh", name: "safe-refresh.xlsx", sourceType: "mmc" },
  doctors: [{ key: "RICHARD HAYDON", displayName: "Richard HAYDON", sourceType: "mmc" }],
  eventsByDoctor: {
    "RICHARD HAYDON": [{ id: "replacement", source: "MMC", title: "Replacement", allDay: true, start: "2026-02-04", end: "2026-02-04", rawValue: "Replacement" }],
  },
}, transactionalDb);
assert.equal(failedRefresh.response.ok, false, "injected replacement failure should surface");
assert.deepEqual(
  [...transactionalDb.events.values()].map((event) => event.title),
  ["Original"],
  "failed derived-file refresh should preserve the previous events transactionally",
);
d1Store.accountProfiles.get("rhaydon@gmail.com").facility_overview_enabled = 0;
const d1CreatorLogin = await postState(d1StateStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, d1Store);
assert.equal(d1CreatorLogin.snapshot?.preview?.derivedFromD1, true, "creator login should return an inline D1-derived snapshot");
assert.equal(d1CreatorLogin.viewedAccountType, "creator", "creator login should identify the viewed account as the creator context");
assert.equal(d1CreatorLogin.isImpersonating, false, "creator login should not impersonate another account");
assert.equal(d1CreatorLogin.facilityOverviewEnabled, true, "Creator At a glance access should remain unconditional even if its stored flag is false");
assert.equal(d1CreatorLogin.state.session.doctorKey, d1Doctor.key, "creator login should keep selected doctor metadata");
assert.equal(d1CreatorLogin.subscription.enabled, true, "creator account should expose subscription URL capability");
const d1DisableCreatorFacilityOverview = await postStateRaw(d1StateStore, {
  action: "setUserFacilityOverviewEnabled",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "rhaydon@gmail.com",
  facilityOverviewEnabled: false,
}, d1Store);
assert.equal(d1DisableCreatorFacilityOverview.response.status, 400, "Creator At a glance access should not be disableable");
const d1CreatorCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, d1Store);
assert.equal(d1CreatorCalendar.snapshot?.preview?.derivedFromD1, true);
assert.ok(d1CreatorCalendar.snapshot.preview.events.length > 0);
assert.ok(d1CreatorCalendar.snapshot.detectedSources.mmc.length > 0, "creator D1 snapshots should retain detected roster sources");
const d1CreatorFeedResponse = await handleFeedGet({
  request: new Request(`http://fixture.test/api/feed?token=${d1CreatorLogin.subscription.token}`),
  env: { ROSTER_DB: d1Store },
});
assert.equal(d1CreatorFeedResponse.ok, true, "creator subscription feed should resolve from the creator-selected doctor");
const d1CreatedUser = await postState(d1StateStore, {
  action: "adminCreateUser",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "d1-user@example.com",
  targetRealName: d1Doctor.displayName,
  targetPassword: "d1-password",
}, d1Store);
assert.ok(d1CreatedUser.user.claims.length > 0, "admin-created account should immediately claim exact roster matches");
assert.equal(d1CreatedUser.user.facilityOverviewEnabled, false, "new standard accounts should require an explicit At a glance grant");
const d1NonClinicalDirector = await postState(d1StateStore, {
  action: "adminCreateUser",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "director@example.com",
  targetRealName: d1Doctor.displayName,
  targetPassword: "director-password",
  nonClinical: true,
  directorViewEnabled: true,
}, d1Store);
assert.equal(d1NonClinicalDirector.user.nonClinical, true, "Creator should be able to create a non-clinical account");
assert.equal(d1NonClinicalDirector.user.directorViewEnabled, true, "Creator should be able to grant Director access on creation");
assert.equal(d1NonClinicalDirector.user.claims.length, 0, "non-clinical accounts must not auto-claim a matching clinician");
const d1NonClinicalDirectorLogin = await postState(d1StateStore, {
  action: "login",
  email: "director@example.com",
  password: "director-password",
}, d1Store);
assert.equal(d1NonClinicalDirectorLogin.nonClinical, true, "non-clinical classification should persist on login");
assert.equal(d1NonClinicalDirectorLogin.directorViewEnabled, true, "Director access should persist on login");
assert.equal(d1NonClinicalDirectorLogin.claims.length, 0, "non-clinical login must remain claim-free");
const d1SelfGrantDirector = await postStateRaw(d1StateStore, {
  action: "setUserDirectorViewEnabled",
  email: "director@example.com",
  password: "director-password",
  targetEmail: "director@example.com",
  directorViewEnabled: false,
}, d1Store);
assert.equal(d1SelfGrantDirector.response.status, 403, "standard users should not be able to change Director access");
const d1RevokedDirector = await postState(d1StateStore, {
  action: "setUserDirectorViewEnabled",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "director@example.com",
  directorViewEnabled: false,
}, d1Store);
assert.equal(d1RevokedDirector.user.directorViewEnabled, false, "Creator should be able to revoke Director access");
const d1DirectLogin = await postState(d1StateStore, {
  action: "login",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(d1DirectLogin.snapshot?.preview?.derivedFromD1, true, "claimed login should return an inline D1-derived snapshot");
assert.equal(d1DirectLogin.viewedAccountType, "claimed-user", "claimed direct login should identify the viewed account type");
assert.equal(d1DirectLogin.isImpersonating, false, "claimed direct login should not be marked as creator impersonation");
assert.equal(d1DirectLogin.state.session.doctorKey, d1Doctor.key, "claimed login should default to the claimed doctor");
assert.equal(d1DirectLogin.facilityOverviewEnabled, false, "full login should deny At a glance until the Creator grants access");
const d1SelfGrantFacilityOverview = await postStateRaw(d1StateStore, {
  action: "setUserFacilityOverviewEnabled",
  email: "d1-user@example.com",
  password: "d1-password",
  targetEmail: "d1-user@example.com",
  facilityOverviewEnabled: true,
}, d1Store);
assert.equal(d1SelfGrantFacilityOverview.response.status, 403, "standard users should not be able to grant themselves At a glance access");
const d1DeniedFacilityOverview = await postStateRaw(d1StateStore, {
  action: "queryFacilityOverviewOnShift",
  email: "d1-user@example.com",
  password: "d1-password",
  facilityKey: "mmc",
  date: "2026-02-03",
}, d1Store);
assert.equal(d1DeniedFacilityOverview.response.status, 403, "At a glance API data should be denied without a per-user grant");
const d1GrantedFacilityOverview = await postState(d1StateStore, {
  action: "setUserFacilityOverviewEnabled",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "d1-user@example.com",
  facilityOverviewEnabled: true,
}, d1Store);
assert.equal(d1GrantedFacilityOverview.user.facilityOverviewEnabled, true, "the Creator should be able to grant At a glance per user");
const d1GrantedUserLogin = await postState(d1StateStore, {
  action: "login",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(d1GrantedUserLogin.facilityOverviewEnabled, true, "an explicit At a glance grant should persist across login");
const d1FastCachedLogin = await postState(d1StateStore, {
  action: "login",
  email: "d1-user@example.com",
  password: "d1-password",
  responseMode: "fast",
}, d1Store);
assert.equal(d1FastCachedLogin.snapshotSource, "server-cache", "fast login should serve a ready server snapshot as current");
assert.equal(d1FastCachedLogin.snapshotStale, false, "fast login should not mark revision-matched server cache stale");
assert.equal(d1FastCachedLogin.snapshot?.preview?.derivedFromD1, true, "fast login should return the cached server snapshot");
const d1FastBrowserRevisionLogin = await postState(d1StateStore, {
  action: "login",
  email: "d1-user@example.com",
  password: "d1-password",
  responseMode: "fast",
  cachedRevision: d1FastCachedLogin.snapshotRevision,
}, d1Store);
assert.equal(d1FastBrowserRevisionLogin.snapshotCurrent, true, "fast login should accept a current browser revision");
assert.equal(d1FastBrowserRevisionLogin.snapshot, null, "fast login should not resend an unchanged browser snapshot");
const d1CurrentRevisionCheck = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
  cachedRevision: d1FastCachedLogin.snapshotRevision,
  allowInlineBuild: false,
}, d1Store);
assert.equal(d1CurrentRevisionCheck.snapshotCurrent, true, "cachedRevision should let the server confirm the visible calendar is current");
assert.equal(d1CurrentRevisionCheck.snapshot, null, "current-revision checks should not resend or replace the snapshot");
const d1UserRegistryKey = [...d1Store.snapshotRegistry.keys()].find((key) => key.startsWith(`user-account|d1-user@example.com|${d1Doctor.key}|`));
assert.ok(d1UserRegistryKey, "claimed D1 account should have a snapshot registry entry");
let d1UserRegistry = d1Store.snapshotRegistry.get(d1UserRegistryKey);
const fastStaleRevision = "fast-stale-revision";
d1UserRegistry.built_revision = fastStaleRevision;
d1UserRegistry.status = "ready";
const fastStaleWhileRevalidate = await postStateRaw(d1StateStore, {
  action: "login",
  email: "d1-user@example.com",
  password: "d1-password",
  responseMode: "fast",
}, d1Store, { captureWaitUntil: true });
assert.equal(fastStaleWhileRevalidate.response.ok, true);
assert.equal(fastStaleWhileRevalidate.body.snapshotRevision, fastStaleRevision, "fast login should immediately serve the latest ready R2 revision");
assert.equal(fastStaleWhileRevalidate.body.diagnostics.login.skippedRevision, true, "fast login should report deferred revision validation");
assert.equal(fastStaleWhileRevalidate.body.diagnostics.login.validationDeferred, false, "fast login should leave snapshot validation to the calendar load request");
assert.match(fastStaleWhileRevalidate.response.headers.get("server-timing") || "", /auth;dur=[\d.]+[\s\S]*r2;dur=[\d.]+/, "fast login should emit standard Server-Timing phases");
assert.equal(fastStaleWhileRevalidate.waitUntilPromises.length, 0, "fast login should not schedule background snapshot work");
await Promise.all(fastStaleWhileRevalidate.waitUntilPromises);
assert.equal(d1Store.snapshotRegistry.get(d1UserRegistryKey)?.built_revision, fastStaleRevision, "fast login should leave the ready snapshot untouched until calendar validation");
d1UserRegistry = d1Store.snapshotRegistry.get(d1UserRegistryKey);
d1UserRegistry.built_revision = "outdated-revision";
d1UserRegistry.status = "ready";
const d1StaleServerCache = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
  allowInlineBuild: false,
}, d1Store);
assert.equal(d1StaleServerCache.snapshotSource, "stale-server-cache", "stale server snapshots should be returned as stale cache");
assert.equal(d1StaleServerCache.snapshotStale, true, "revision-mismatched server cache should be marked stale");
d1UserRegistry.built_revision = d1FastCachedLogin.snapshotRevision;
d1UserRegistry.status = "building";
d1UserRegistry.updated_at = new Date().toISOString();
d1UserRegistry.size_bytes = 0;
d1StateStore.r2.objects.delete(d1UserRegistry.artifact_key);
const d1RecentBuilding = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
  allowInlineBuild: true,
}, d1Store);
assert.equal(d1RecentBuilding.snapshotSource, "server-cache-building", "recent building rows should suppress duplicate inline builds");
assert.equal(d1RecentBuilding.snapshot, null, "recent building rows should not replace the visible snapshot");
d1UserRegistry.updated_at = "2000-01-01T00:00:00.000Z";
const d1ExpiredBuildingRetry = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
  allowInlineBuild: true,
}, d1Store);
assert.equal(d1ExpiredBuildingRetry.snapshotSource, "d1-build", "expired building rows should be retried by the next real request");
assert.equal(d1ExpiredBuildingRetry.snapshot?.preview?.derivedFromD1, true);
const d1CacheMissNoInline = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
  allowInlineBuild: false,
  doctorKey: "NOT A REAL DOCTOR",
}, d1Store);
assert.equal(d1CacheMissNoInline.snapshotSource, "server-cache-miss", "non-inline cache misses should not build synchronously");
assert.equal(d1CacheMissNoInline.snapshot, null);
const d1DirectCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.ok(d1DirectCalendar.snapshot.detectedSources.mmc.length > 0, "claimed D1 snapshots should retain linked roster sources");
const staleResolvedDdhIssue = {
  id: "DDH::Unknown::2026-06-08::PHNW CS",
  source: "DDH",
  seniority: "Unknown",
  startDay: "2026-06-08",
  rawValue: "PHNW CS",
  code: "PHNW CS",
  status: "unknown",
  message: "DDH shift label not recognised.",
  resolutionType: "shift_code",
  suggestedTitle: "DDH: PHNW",
  timeLabel: "",
};
const currentD1UserRegistry = d1Store.snapshotRegistry.get(d1UserRegistryKey);
await storeCachedSnapshot(d1StateStore.r2, currentD1UserRegistry.artifact_key, {
  ...d1DirectCalendar.snapshot,
  preview: {
    ...d1DirectCalendar.snapshot.preview,
    issues: [...(d1DirectCalendar.snapshot.preview.issues || []), staleResolvedDdhIssue],
  },
}, {
  revision: d1DirectCalendar.snapshotRevision,
  ownerType: currentD1UserRegistry.owner_type,
  ownerId: currentD1UserRegistry.owner_id,
  doctorKey: currentD1UserRegistry.doctor_key,
  rangeKey: currentD1UserRegistry.range_key,
});
await postState(d1StateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  source: "DDH",
  rawValue: "PHNW CS",
  rule: {
    source: "DDH",
    seniority: "SMS",
    code: "PHNW CS",
    kind: "shift",
    base: "PHNW",
    period: "",
    suffix: "",
    allDay: true,
    startTime: "",
    endTime: "",
    location: "",
    includeAsShift: true,
  },
}, d1Store);
const d1ParserRevisionCheck = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
  cachedRevision: d1DirectCalendar.snapshotRevision,
  allowInlineBuild: false,
}, d1Store);
assert.notEqual(d1ParserRevisionCheck.snapshotCurrent, true, "parser-rule changes should invalidate cachedRevision checks");
assert.equal(d1ParserRevisionCheck.snapshotSource, "stale-server-cache", "parser-rule revision changes may reuse stale server cache without inline build");
assert.doesNotMatch(
  stateApiSource,
  /if \(calendarRevision && String\(body\?\.cachedRevision \|\| ""\) === calendarRevision\) \{[\s\S]{0,1000}?snapshotCurrent: true/,
  "doctor-profile cache validation must verify its own snapshot registry rather than accepting a browser-wide revision alone",
);
assert.equal(
  d1ParserRevisionCheck.snapshot.preview.issues.some((issue) => issue.rawValue === "PHNW CS"),
  false,
  "resolved parser warnings should be filtered out of stale server snapshots before return",
);
await postState(d1StateStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: {
    id: "ddh-rover-diagnostics",
    name: "Dandenong 2026.xlsx",
    sourceType: "ddh",
    active: true,
  },
  doctors: [{ key: d1Doctor.key, displayName: d1Doctor.displayName, sourceType: "ddh" }],
  eventsByDoctor: {
    [d1Doctor.key]: [{
      id: "ddh-rover-event",
      source: "DDH",
      seniority: "Unknown",
      title: "DDH: Rover AM",
      allDay: false,
      start: "2026-05-14T08:00:00",
      end: "2026-05-14T18:00:00",
      rawValue: "Rover AM",
      timeLabel: "08:00-18:00",
    }],
  },
  issuesByDoctor: {
    [d1Doctor.key]: [
      {
        id: "DDH::Unknown::2026-05-14::Rover AM",
        source: "DDH",
        seniority: "Unknown",
        startDay: "2026-05-14",
        rawValue: "Rover AM",
        status: "unknown",
        message: "DDH shift code not recognised; using explicit roster time.",
        resolutionType: "shift_code",
        suggestedTitle: "DDH: Rover AM",
        timeLabel: "08:00-18:00",
      },
      {
        id: "DDH::Unknown::2026-06-08::PHNW CS",
        source: "DDH",
        seniority: "Unknown",
        startDay: "2026-06-08",
        rawValue: "PHNW CS",
        code: "PHNW CS",
        status: "unknown",
        message: "DDH shift label not recognised.",
        resolutionType: "shift_code",
        suggestedTitle: "DDH: PHNW",
        timeLabel: "",
      },
      {
        id: "DDH::Unknown::2026-05-15::Mystery AM",
        source: "DDH",
        seniority: "Unknown",
        startDay: "2026-05-15",
        rawValue: "Mystery AM",
        status: "unknown",
        message: "DDH shift code not recognised; using explicit roster time.",
        resolutionType: "shift_code",
        suggestedTitle: "DDH: Mystery AM",
        timeLabel: "08:00-18:00",
      },
    ],
  },
}, d1Store);
assert.equal(
  [...d1Store.issues.values()].filter((issue) => issue.file_id === "ddh-rover-diagnostics").length,
  3,
  "derived roster saves should persist parser diagnostics transactionally",
);
const d1StoredIssueCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.ok(
  d1StoredIssueCalendar.snapshot.preview.events.some((event) => event.title === "DDH: Rover AM"),
  "DDH Rover shifts should render from indexed events",
);
assert.equal(
  d1StoredIssueCalendar.snapshot.preview.issues.some((issue) => issue.rawValue === "Rover AM"),
  false,
  "stale DDH Rover diagnostics should be hidden on first D1 calendar load",
);
assert.equal(
  d1StoredIssueCalendar.snapshot.preview.issues.some((issue) => issue.rawValue === "PHNW CS"),
  false,
  "server-built snapshots should hide roster issues resolved by active parser rules",
);
assert.equal(
  d1StoredIssueCalendar.snapshot.preview.issues.some((issue) => issue.rawValue === "Mystery AM"),
  true,
  "calendar loads should show unresolved diagnostics stored during import",
);
await postState(d1StateStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: {
    id: "ddh-rover-diagnostics",
    name: "Dandenong 2026.xlsx",
    sourceType: "ddh",
    active: true,
  },
  doctors: [{ key: d1Doctor.key, displayName: d1Doctor.displayName, sourceType: "ddh" }],
  eventsByDoctor: {
    [d1Doctor.key]: [{
      id: "ddh-rover-event",
      source: "DDH",
      seniority: "Unknown",
      title: "DDH: Rover AM",
      allDay: false,
      start: "2026-05-14T08:00:00",
      end: "2026-05-14T18:00:00",
      rawValue: "Rover AM",
      timeLabel: "08:00-18:00",
    }],
  },
  issuesByDoctor: {},
}, d1Store);
assert.equal(
  [...d1Store.issues.values()].some((issue) => issue.file_id === "ddh-rover-diagnostics"),
  false,
  "reparsed roster files should replace stale diagnostics instead of accumulating them",
);
const d1ClearedIssueCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(
  d1ClearedIssueCalendar.snapshot.preview.issues.some((issue) => issue.rawValue === "Mystery AM"),
  false,
  "cleared import diagnostics should not reappear on later D1 calendar loads",
);
for (const [key, event] of [...d1Store.events.entries()]) if (event.file_id === "ddh-rover-diagnostics") d1Store.events.delete(key);
for (const [key, issue] of [...d1Store.issues.entries()]) if (issue.file_id === "ddh-rover-diagnostics") d1Store.issues.delete(key);
for (const [key, doctor] of [...d1Store.fileDoctors.entries()]) if (doctor.file_id === "ddh-rover-diagnostics") d1Store.fileDoctors.delete(key);
d1Store.files.delete("ddh-rover-diagnostics");
await seedUser(d1StateStore, "admin-enter-match@example.com", "admin-enter-password", d1Doctor.displayName, d1Store);
for (const key of [...d1Store.accountClaims.keys()]) {
  if (key.startsWith("admin-enter-match@example.com|")) d1Store.accountClaims.delete(key);
}
const adminEnteredResolution = await postState(d1StateStore, {
  action: "resolveAccountClaims",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "admin-enter-match@example.com",
}, d1Store);
assert.equal(
  adminEnteredResolution.claims.some((claim) => claim.key === d1Doctor.key && claim.sourceType === "mmc"),
  false,
  "automatic matching must not give a second account an already-claimed clinician identity",
);
assert.equal(
  adminEnteredResolution.availableDoctors.find((doctor) => doctor.key === d1Doctor.key)?.claimedBy,
  "d1-user@example.com",
  "an unclaimed account should see who already owns its matching roster identity",
);
for (const key of [...d1Store.accountClaims.keys()]) {
  if (key.startsWith("admin-enter-match@example.com|")) d1Store.accountClaims.delete(key);
}
const d1OnsiteMmcEvent = d1DirectCalendar.snapshot.preview.events.find((event) => event.source === "MMC" && !/\\b(CS|leave|conference|PHNW)\\b/i.test(`${event.title} ${event.rawValue}`));
assert.ok(d1OnsiteMmcEvent?.location, "D1 account calendar load should apply SQL-backed hospital defaults to onsite shifts");
await postState(d1StateStore, {
  action: "save",
  email: "d1-user@example.com",
  password: "d1-password",
  state: {
    version: 1,
    imports: d1DirectLogin.state.imports,
    session: {
      doctorKey: d1Doctor.key,
      settings: {
        defaultLocationMmc: "User One MMC Location",
      },
    },
  },
}, d1Store);
for (const row of d1Store.events.values()) {
  const event = JSON.parse(row.event_json);
  if (event.source !== "MMC" || /\b(CS|leave|conference|PHNW)\b/i.test(`${event.title} ${event.rawValue}`)) continue;
  event.location = "";
  row.location = "";
  row.event_json = JSON.stringify(event);
  break;
}
const d1CustomLocationCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.ok(
  d1CustomLocationCalendar.snapshot.preview.events.some((event) => event.source === "MMC" && event.location === "User One MMC Location"),
  "account SQL hospital location should override shared roster-event location",
);
const d1UserFeedResponse = await handleFeedGet({
  request: new Request(`http://fixture.test/api/feed?token=${d1DirectLogin.subscription.token}`),
  env: { ROSTER_DB: d1Store },
});
assert.equal(d1UserFeedResponse.ok, true, "claimed user subscription feed should resolve");
assert.match(await d1UserFeedResponse.text(), /LOCATION:User One MMC Location/, "subscription feed should apply account hospital defaults");
const d1OverrideEvent = d1DirectCalendar.snapshot.preview.events[0];
await postState(d1StateStore, {
  action: "save",
  email: "d1-user@example.com",
  password: "d1-password",
  state: {
    version: 1,
    imports: d1DirectLogin.state.imports,
    session: {
      doctorKey: d1Doctor.key,
      exportRange: { startDate: "2026-02-01", endDate: "2026-02-28", allFuture: false },
      settings: {},
      overrides: {
        [d1OverrideEvent.id]: {
          title: "D1 Edited Shift",
        },
      },
      customEvents: [{
        id: "d1-custom-event",
        ownerEmail: "d1-user@example.com",
        title: "D1 Custom Event stale",
        startDate: "2026-02-12",
        endDate: "2026-02-12",
        allDay: true,
        include: true,
      }, {
        id: "d1-custom-event",
        ownerEmail: "d1-user@example.com",
        title: "D1 Custom Event",
        startDate: "2026-02-13",
        endDate: "2026-02-13",
        allDay: true,
        include: true,
      }],
    },
  },
}, d1Store);
const d1SessionLogin = await postState(d1StateStore, {
  action: "login",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(d1SessionLogin.state.session.exportRange.startDate, "2026-02-01", "D1 account session should load without KV state");
assert.deepEqual(
  d1SessionLogin.state.session.customEvents.map((event) => [event.id, event.title, event.startDate]),
  [["d1-custom-event", "D1 Custom Event", "2026-02-13"]],
  "D1 session save should collapse duplicate custom event ids with the latest value winning",
);
await postState(d1StateStore, {
  action: "save",
  email: "d1-user@example.com",
  password: "d1-password",
  state: {
    ...d1SessionLogin.state,
    session: {
      ...d1SessionLogin.state.session,
      customEvents: [
        ...d1SessionLogin.state.session.customEvents,
        { id: "different-id-same-event", ownerEmail: "d1-user@example.com", title: "D1 Custom Event", startDate: "2026-02-13", endDate: "2026-02-13", allDay: true, include: true },
      ],
    },
  },
}, d1Store);
const d1IdentityDedupedLogin = await postState(d1StateStore, {
  action: "login",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(d1IdentityDedupedLogin.state.session.customEvents.length, 1, "matching logical custom events should collapse even when ids differ");
const d1SessionCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.ok(d1SessionCalendar.snapshot.preview.events.some((event) => event.title === "D1 Edited Shift"), "D1 calendar load should apply session overrides");
assert.ok(d1SessionCalendar.snapshot.preview.events.some((event) => event.title === "D1 Custom Event"), "D1 calendar load should include session custom events");
assert.equal(d1SessionCalendar.snapshot.preview.customEventsMaterialized, true, "D1 calendar snapshots should declare custom events already materialized");
assert.equal(
  d1SessionCalendar.snapshot.preview.events.filter((event) => event.title === "D1 Custom Event").length,
  1,
  "D1 calendar snapshots should carry each logical custom event once",
);
const d1AdminLoad = await postState(d1StateStore, {
  action: "adminLoadUser",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "d1-user@example.com",
}, d1Store);
assert.equal(d1AdminLoad.snapshot?.preview?.derivedFromD1, true, "admin user switch should return an inline D1-derived snapshot");
assert.equal(d1AdminLoad.viewedAccountId, "d1-user@example.com", "creator switch should identify the viewed account");
assert.equal(d1AdminLoad.viewedAccountType, "claimed-user", "creator switch should retain the target user's account type");
assert.equal(d1AdminLoad.isImpersonating, true, "creator switch should explicitly enter impersonation mode");
assert.equal(d1AdminLoad.returnToCreatorAvailable, true, "creator switch should advertise the Back to creator affordance");
assert.equal(d1AdminLoad.snapshot.ownerType, d1DirectLogin.snapshot.ownerType, "creator switch and direct login should reuse the same canonical snapshot ownership");
const d1AdminRename = await postState(d1StateStore, {
  action: "updateAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "d1-user@example.com",
  realName: "Corrected User Name",
}, d1Store);
assert.equal(d1AdminRename.realName, "Corrected User Name", "creator should be able to correct another account's real name");
assert.equal(d1AdminRename.user.realName, "Corrected User Name", "admin name updates should return a refreshed Current users summary");
const d1AdminCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "d1-user@example.com",
}, d1Store);
assert.deepEqual(
  d1AdminCalendar.snapshot.preview.events.map((event) => event.id),
  d1SessionCalendar.snapshot.preview.events.map((event) => event.id),
  "direct and creator-switched user loads should use the same D1 calendar events",
);

const recreateDb = new MemoryD1();
const recreateStore = new MemoryStore();
seedD1Repository(recreateDb, [{
  id: "recreate-roster",
  name: "recreate.xlsx",
  sourceType: "mmc",
  active: true,
  doctors: [{ key: "TITUS HACKMAN", displayName: "Titus Hackman", sourceType: "mmc" }],
  eventsByDoctor: {
    "TITUS HACKMAN": [{ id: "recreate-shift", source: "MMC", title: "Roster shift", seniority: "Intern", allDay: true, start: "2026-02-03", end: "2026-02-03", rawValue: "Roster shift" }],
  },
}]);
recreateDb.events.set("recreate-event", {
  id: "recreate-event",
  file_id: "recreate-roster",
  source_type: "mmc",
  doctor_key: "TITUS HACKMAN",
  display_name: "Titus Hackman",
  start_date: "2026-02-03",
  end_date: "2026-02-03",
  start_ts: "2026-02-03",
  end_ts: "2026-02-03",
  title: "Roster shift",
  raw_value: "Roster shift",
  seniority: "Intern",
  location: "",
  all_day: 1,
  time_label: "",
  event_json: JSON.stringify({ id: "recreate-shift", source: "MMC", title: "Roster shift", seniority: "Intern", allDay: true, start: "2026-02-03", end: "2026-02-03", rawValue: "Roster shift" }),
});
await postState(recreateStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, recreateDb);
const firstMatchedCreate = await postState(recreateStore, {
  action: "adminCreateUser",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "recreated@example.com",
  targetRealName: "Titus Hackman",
  targetPassword: "recreated-password",
}, recreateDb);
assert.equal(firstMatchedCreate.user.claims.length, 1, "first matched create should auto-claim the roster doctor");
assert.deepEqual(firstMatchedCreate.user.seniorities, ["Intern"], "user summaries should include roster-derived seniorities");
await postState(recreateStore, {
  action: "deleteAccount",
  email: "recreated@example.com",
  password: "recreated-password",
}, recreateDb);
const recreatedMatchedUser = await postState(recreateStore, {
  action: "adminCreateUser",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "recreated@example.com",
  targetRealName: "Titus Hackman",
  targetPassword: "recreated-password-2",
}, recreateDb);
assert.equal(recreatedMatchedUser.user.claims.length, 1, "recreating a self-deleted matched user should auto-claim cleanly");
const madeUpUser = await postState(recreateStore, {
  action: "adminCreateUser",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "made-up@example.com",
  targetRealName: "Made Up",
  targetPassword: "made-up-password",
}, recreateDb);
assert.equal(madeUpUser.user.claims.length, 0, "creating a no-roster user should remain lightweight and unclaimed");
const d1Insights = await postState(d1StateStore, {
  action: "queryRosterInsights",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  startDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
  endDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
}, d1Store);
assert.ok(Array.isArray(d1Insights.coworkers));
assert.equal(d1Insights.source, "roster-events", "coworker lookup should use the proven direct roster_events path");
assert.ok(d1Store.dailyPresence.size > 0, "daily presence should be populated during derived roster storage");
assert.ok(
  [...d1Store.dailyPresence.values()].some((row) => d1Store.events.has(row.event_id)),
  "daily presence rows should include references to stored roster event ids",
);
const d1InsightsExcludingSelectedDoctor = await postState(d1StateStore, {
  action: "queryRosterInsights",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  startDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
  endDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
  excludeDoctorKeys: [d1Doctor.key],
}, d1Store);
assert.ok(
  d1InsightsExcludingSelectedDoctor.coworkers.every((row) => row.doctorKey !== d1Doctor.key),
  "coworker lookup should exclude the selected doctor in SQL",
);
const d1OverlapInsights = await postState(d1StateStore, {
  action: "queryRosterInsights",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  startDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
  endDate: d1CreatorCalendar.snapshot.preview.events.at(-1).start.slice(0, 10),
  overlapDoctorKeys: [d1Doctor.key],
  excludeDoctorKeys: [d1Doctor.key],
}, d1Store);
assert.ok(Array.isArray(d1OverlapInsights.coworkers), "overlap coworker lookup should return SQL-derived rows");
assert.ok(
  d1OverlapInsights.coworkers.every((row) => row.doctorKey !== d1Doctor.key),
  "overlap coworker lookup should exclude the selected doctor in SQL",
);
const d1OverlapDoctors = await postState(d1StateStore, {
  action: "queryRosterOverlapDoctors",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  startDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
  endDate: d1CreatorCalendar.snapshot.preview.events.at(-1).start.slice(0, 10),
  overlapDoctorKeys: [d1Doctor.key],
  excludeDoctorKeys: [d1Doctor.key],
}, d1Store);
assert.ok(Array.isArray(d1OverlapDoctors.doctors), "overlap doctor lookup should return compact doctor rows");
assert.equal(d1OverlapDoctors.source, "roster-events", "overlap doctor lookup should use the proven direct roster_events path");
const savedEventJsonWithoutDoctorMetadata = [...d1Store.events.values()].find((event) => {
  const parsed = JSON.parse(event.event_json || "{}");
  return !parsed.doctorKey && !parsed.displayName;
});
assert.ok(savedEventJsonWithoutDoctorMetadata, "fixture should cover stored events without doctor metadata in event_json");
d1Store.dailyPresence.clear();
const d1WarmupStyleInsights = await postState(d1StateStore, {
  action: "queryRosterInsights",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  startDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
  endDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
  excludeDoctorKeys: [d1Doctor.key],
  allowFallback: false,
}, d1Store);
assert.equal(d1WarmupStyleInsights.source, "roster-events", "insight requests should no longer depend on daily presence");
assert.ok(d1WarmupStyleInsights.coworkers.length, "direct roster_events insight lookup should still return rows when daily presence is empty");
assert.equal(d1Store.dailyPresence.size, 0, "warmup insight requests should not repair daily presence");
const d1FallbackInsights = await postState(d1StateStore, {
  action: "queryRosterInsights",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  startDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
  endDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
  excludeDoctorKeys: [d1Doctor.key],
  allowFallback: true,
}, d1Store);
assert.equal(d1FallbackInsights.source, "roster-events", "explicit insight requests should use roster_events directly");
assert.ok(d1FallbackInsights.coworkers.length, "direct coworker lookup should still return roster rows");
assert.equal(d1Store.dailyPresence.size, 0, "explicit insight requests should not run a full daily-presence repair");
const d1FallbackOverlapDoctors = await postState(d1StateStore, {
  action: "queryRosterOverlapDoctors",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  startDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
  endDate: d1CreatorCalendar.snapshot.preview.events.at(-1).start.slice(0, 10),
  overlapDoctorKeys: [d1Doctor.key],
  excludeDoctorKeys: [d1Doctor.key],
  allowFallback: true,
}, d1Store);
assert.equal(d1FallbackOverlapDoctors.source, "roster-events", "explicit overlap requests should use roster_events directly");
const repairedPresence = await postState(d1StateStore, {
  action: "repairRosterDailyPresence",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  limit: 1,
  offset: 0,
}, d1Store);
assert.ok(repairedPresence.repaired.files > 0, "daily presence repair should process active files");
assert.ok(Object.hasOwn(repairedPresence.repaired, "done"), "daily presence repair should return bounded batch state");
assert.ok(d1Store.dailyPresence.size > 0, "daily presence repair should rebuild rows from roster_events");
const d1RepairedOverlapDoctors = await postState(d1StateStore, {
  action: "queryRosterOverlapDoctors",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  startDate: d1CreatorCalendar.snapshot.preview.events[0].start.slice(0, 10),
  endDate: d1CreatorCalendar.snapshot.preview.events.at(-1).start.slice(0, 10),
  overlapDoctorKeys: [d1Doctor.key],
  excludeDoctorKeys: [d1Doctor.key],
}, d1Store);
assert.equal(d1RepairedOverlapDoctors.source, "roster-events", "repaired daily presence should not change the direct overlap path");
const d1RepositoryFile = [...d1Store.files.keys()][0];
await postState(d1StateStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: {
    id: d1RepositoryFile,
    name: "AdultTerm1.2026.xlsx",
    sourceType: "mmc",
    active: true,
  },
  doctors: [...d1Store.fileDoctors.values()].map((doctor) => ({
    key: doctor.doctor_key,
    displayName: doctor.display_name,
    sourceType: doctor.source_type,
  })),
  eventsByDoctor: Object.fromEntries(
    [...d1Store.fileDoctors.values()].map((doctor) => [
      doctor.doctor_key,
      [...d1Store.events.values()]
        .filter((event) => event.doctor_key === doctor.doctor_key)
        .map((event) => JSON.parse(event.event_json)),
    ]),
  ),
}, d1Store);
const d1Status = await postState(d1StateStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  expectedFileIds: ["d1-mmc", "missing-d1-mmc"],
}, d1Store);
assert.equal(d1Status.total, 1);
assert.equal(d1Status.populated, 1);
assert.equal(d1Status.remaining, 0);
assert.equal(d1Status.expectedFiles.expectedCount, 2);
assert.equal(d1Status.expectedFiles.persistedCount, 1);
assert.equal(d1Status.expectedFiles.populatedCount, 1);
assert.deepEqual(d1Status.expectedFiles.persistedFileIds, ["d1-mmc"]);
assert.deepEqual(d1Status.expectedFiles.populatedFileIds, ["d1-mmc"]);
assert.deepEqual(d1Status.expectedFiles.missingFileIds, ["missing-d1-mmc"]);
assert.ok(d1Status.accounts.profiles >= 2);
assert.ok(d1Status.accounts.claims >= 1);
assert.ok(d1Status.accounts.states >= 2);
assert.ok(d1Status.accounts.subscriptionTokens >= 2);
await d1StateStore.delete("repository:index");
const d1OnlyRepositoryStatus = await postState(d1StateStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, d1Store);
assert.equal(d1OnlyRepositoryStatus.total, 1, "D1 roster_files should supply calendar status without KV repository index");
assert.equal(d1OnlyRepositoryStatus.populated, 1);
const d1UserList = await postState(d1StateStore, {
  action: "listUsers",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, d1Store);
assert.ok(d1UserList.availableDoctors.some((doctor) => doctor.key === d1Doctor.key));
assert.equal(
  d1UserList.availableDoctors.find((doctor) => doctor.key === d1Doctor.key)?.claimedBy,
  "d1-user@example.com",
  "D1 doctor directory should include claimed account metadata for the Creator switcher",
);
const d1NoKvIndexLogin = await postState(d1StateStore, {
  action: "login",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(d1NoKvIndexLogin.snapshot?.preview?.derivedFromD1, true, "D1 account login should return a D1-derived snapshot without KV repository index");
const d1NoKvIndexCalendar = await postState(d1StateStore, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(d1NoKvIndexCalendar.snapshot?.preview?.derivedFromD1, true, "D1 account calendar load should work without KV repository index");
assert.ok(d1NoKvIndexCalendar.snapshot.fileRefs.some((ref) => ref.id === d1RepositoryFile), "D1 claimed-account calendar snapshots should include source file refs for the Account modal");
const d1NoKvIndexEnrichment = await postState(d1StateStore, {
  action: "resolveAccountClaims",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.deepEqual(d1NoKvIndexEnrichment.availableDoctors, [], "claimed-account enrichment should not reload the full doctor directory");
await seedUser(d1StateStore, "d1-unmatched@example.com", "d1-unmatched-password", "Unmatched Person", d1Store);
const d1UnmatchedEnrichment = await postState(d1StateStore, {
  action: "resolveAccountClaims",
  email: "d1-unmatched@example.com",
  password: "d1-unmatched-password",
}, d1Store);
assert.ok(d1UnmatchedEnrichment.availableDoctors.some((doctor) => doctor.key === d1Doctor.key), "D1 doctor directory should load only when claim resolution still leaves the account unclaimed");
const d1ClaimResolution = await postState(d1StateStore, {
  action: "resolveDoctorAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctor: {
    key: d1Doctor.key,
    displayName: d1Doctor.displayName,
    sourceTypes: ["mmc"],
  },
}, d1Store);
assert.equal(d1ClaimResolution.mode, "claimed-account");
assert.equal(d1ClaimResolution.email, "d1-user@example.com");
const d1DoctorProfile = await postState(d1StateStore, {
  action: "loadDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: `${d1Doctor.key}::mmc`,
  doctorKey: d1Doctor.key,
  displayName: d1Doctor.displayName,
  sourceTypes: ["mmc"],
}, d1Store);
assert.equal(d1DoctorProfile.snapshot?.preview?.derivedFromD1, true);
assert.equal(d1DoctorProfile.snapshotStale, false);
assert.ok(d1DoctorProfile.snapshot.preview.events.length > 0);
assert.ok(d1DoctorProfile.snapshot.fileRefs.some((ref) => ref.id === d1RepositoryFile), "D1 doctor profile should derive file refs without KV repository index");
const d1DoctorProfileServerCache = await postState(d1StateStore, {
  action: "loadDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: `${d1Doctor.key}::mmc`,
  doctorKey: d1Doctor.key,
  displayName: d1Doctor.displayName,
  sourceTypes: ["mmc"],
  allowInlineBuild: false,
}, d1Store);
assert.equal(d1DoctorProfileServerCache.snapshotSource, "server-cache", "doctor-profile loads should reuse ready server cache generically");
assert.equal(d1DoctorProfileServerCache.snapshotStale, false);
const d1DoctorProfileCurrent = await postState(d1StateStore, {
  action: "loadDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: `${d1Doctor.key}::mmc`,
  doctorKey: d1Doctor.key,
  displayName: d1Doctor.displayName,
  sourceTypes: ["mmc"],
  cachedRevision: d1DoctorProfileServerCache.snapshotRevision,
  allowInlineBuild: false,
}, d1Store);
assert.notEqual(d1DoctorProfileCurrent.snapshotCurrent, true, "doctor-profile cachedRevision must not bypass its own snapshot-registry validation");
assert.ok(
  !d1DoctorProfileCurrent.snapshot || d1DoctorProfileCurrent.snapshot.preview?.derivedFromD1,
  "doctor-profile cachedRevision must never substitute another profile's calendar data",
);
const d1RepositoryDoctors = [...d1Store.fileDoctors.values()].map((doctor) => ({
  key: doctor.doctor_key,
  displayName: doctor.display_name,
  sourceType: doctor.source_type,
}));
const d1RepositoryEventsByDoctor = Object.fromEntries(
  [...d1Store.fileDoctors.values()].map((doctor) => [
    doctor.doctor_key,
    [...d1Store.events.values()]
      .filter((event) => event.doctor_key === doctor.doctor_key)
      .map((event) => JSON.parse(event.event_json)),
  ]),
);
await postState(d1StateStore, {
  action: "resetDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  fileId: d1RepositoryFile,
}, d1Store);
const d1InactiveRepositoryStatus = await postState(d1StateStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, d1Store);
assert.equal(d1InactiveRepositoryStatus.total, 0, "inactive D1 roster files should not drive active calendar status");
await postState(d1StateStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: {
    id: d1RepositoryFile,
    name: "AdultTerm1.2026.xlsx",
    sourceType: "mmc",
    active: true,
  },
  doctors: d1RepositoryDoctors,
  eventsByDoctor: d1RepositoryEventsByDoctor,
}, d1Store);
const d1ProfileOverrideEvent = d1DoctorProfile.snapshot.preview.events[0];
await postState(d1StateStore, {
  action: "saveDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: `${d1Doctor.key}::mmc`,
  doctorKey: d1Doctor.key,
  displayName: d1Doctor.displayName,
  sourceTypes: ["mmc"],
  state: {
    version: 1,
    imports: [],
    session: {
      doctorKey: d1Doctor.key,
      hadPreview: true,
      settings: {},
      overrides: {
        [d1ProfileOverrideEvent.id]: { title: "D1 Profile Edited Shift" },
      },
      customEvents: [{
        id: "d1-profile-custom-event",
        title: "D1 Profile Custom Event",
        startDate: "2026-02-13",
        endDate: "2026-02-13",
        allDay: true,
        include: true,
      }],
    },
  },
}, d1Store);
await d1StateStore.delete(`doctor-profile:${d1Doctor.key}::mmc`);
await d1StateStore.delete(`snapshot:doctor-profile:${d1Doctor.key}::mmc`);
const d1OnlyDoctorProfile = await postState(d1StateStore, {
  action: "loadDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: `${d1Doctor.key}::mmc`,
  doctorKey: d1Doctor.key,
  displayName: d1Doctor.displayName,
  sourceTypes: ["mmc"],
}, d1Store);
assert.equal(d1OnlyDoctorProfile.snapshot?.preview?.derivedFromD1, true);
assert.equal(d1OnlyDoctorProfile.snapshot?.preview?.customEventsMaterialized, true);
assert.ok(d1OnlyDoctorProfile.snapshot.preview.events.some((event) => event.title === "D1 Profile Edited Shift"), "D1 doctor profile should apply stored overrides without KV profile state");
assert.ok(d1OnlyDoctorProfile.snapshot.preview.events.some((event) => event.title === "D1 Profile Custom Event"), "D1 doctor profile should include stored custom events without KV profile state");
const d1ProfileStatus = await postState(d1StateStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, d1Store);
assert.ok(d1ProfileStatus.accounts.doctorProfiles >= 1);
const leaveMergeStore = new MemoryStore();
const leaveMergeDb = new MemoryD1();
const leaveDoctor = { key: "LEAVE DOCTOR", displayName: "Leave Doctor", sourceType: "mmc" };
await seedRepository(leaveMergeStore, [
  repositoryFile("leave-mmc", { sourceType: "mmc", doctors: [leaveDoctor] }),
  repositoryFile("leave-ddh", { sourceType: "ddh", doctors: [{ ...leaveDoctor, sourceType: "ddh" }] }),
]);
await postState(leaveMergeStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, leaveMergeDb);
for (const [fileId, sourceType] of [["leave-mmc", "mmc"], ["leave-ddh", "ddh"]]) {
  await postState(leaveMergeStore, {
    action: "saveDerivedCalendarFile",
    email: "rhaydon@gmail.com",
    password: creatorPassword,
    file: { id: fileId, name: `${fileId}.xlsx`, sourceType, active: true },
    doctors: [{ ...leaveDoctor, sourceType }],
    eventsByDoctor: {
      [leaveDoctor.key]: [{
        id: `${fileId}-leave`,
        source: sourceType.toUpperCase() === "DDH" ? "DDH" : "MMC",
        title: "Annual Leave",
        allDay: true,
        start: "2026-04-06",
        end: "2026-04-13",
        rawValue: "Annual Leave",
        monthKey: "2026-04",
      }],
    },
  }, leaveMergeDb);
}
await seedUser(leaveMergeStore, "leave@example.com", "leave-password", "Leave Doctor", leaveMergeDb);
await postState(leaveMergeStore, {
  action: "setAccountRosterClaims",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "leave@example.com",
  claims: [{ sourceType: "mmc", key: leaveDoctor.key }, { sourceType: "ddh", key: leaveDoctor.key }],
}, leaveMergeDb);
const leaveLogin = await postState(leaveMergeStore, {
  action: "login",
  email: "leave@example.com",
  password: "leave-password",
}, leaveMergeDb);
assert.equal(leaveLogin.snapshot?.preview?.derivedFromD1, true, "duplicate leave account login should return a D1-derived snapshot");
const leaveCalendar = await postState(leaveMergeStore, {
  action: "loadCalendarEvents",
  email: "leave@example.com",
  password: "leave-password",
}, leaveMergeDb);
const mergedLeave = leaveCalendar.snapshot.preview.events.filter((event) => event.title === "Annual Leave");
assert.equal(mergedLeave.length, 1);
assert.deepEqual(mergedLeave[0].sources.sort(), ["DDH", "MMC"]);
const aliasClaimsStore = new MemoryStore();
const aliasClaimsDb = new MemoryD1();
await postState(aliasClaimsStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, aliasClaimsDb);
for (const [fileId, sourceType, key, title] of [
  ["alias-mmc", "mmc", "ALIAS DOCTOR", "MMC Alias Shift"],
  ["alias-ddh", "ddh", "DR ALIAS DOCTOR", "DDH Alias Shift"],
]) {
  await postState(aliasClaimsStore, {
    action: "saveDerivedCalendarFile",
    email: "rhaydon@gmail.com",
    password: creatorPassword,
    file: { id: fileId, name: `${fileId}.xlsx`, sourceType, active: true },
    doctors: [{ key, displayName: "Alias Doctor", sourceType }],
    eventsByDoctor: {
      [key]: [{
        id: `${fileId}-shift`,
        source: sourceType.toUpperCase(),
        title,
        allDay: true,
        start: "2026-05-18",
        end: "2026-05-19",
        rawValue: title,
        monthKey: "2026-05",
      }],
    },
  }, aliasClaimsDb);
}
await seedUser(aliasClaimsStore, "alias@example.com", "alias-password", "Alias Doctor", aliasClaimsDb);
await postState(aliasClaimsStore, {
  action: "setAccountRosterClaims",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "alias@example.com",
  claims: [
    { sourceType: "mmc", key: "ALIAS DOCTOR" },
    { sourceType: "ddh", key: "DR ALIAS DOCTOR" },
  ],
}, aliasClaimsDb);
const aliasCalendar = await postState(aliasClaimsStore, {
  action: "loadCalendarEvents",
  email: "alias@example.com",
  password: "alias-password",
  doctorKey: "ALIAS DOCTOR",
}, aliasClaimsDb);
assert.deepEqual(
  aliasCalendar.snapshot.preview.events.map((event) => event.title).sort(),
  ["DDH Alias Shift", "MMC Alias Shift"],
  "calendar load should include all selected doctor alias keys across hospitals",
);
const typoAliasStore = new MemoryStore();
const typoAliasDb = new MemoryD1();
await postState(typoAliasStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, typoAliasDb);
await postState(typoAliasStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "aeshan-mch", name: "aeshan-mch.xlsx", sourceType: "mch", active: true },
  doctors: [{ key: "AESHAN KULARATNE", displayName: "Aeshan KULARATNE", sourceType: "mch" }],
  eventsByDoctor: { "AESHAN KULARATNE": [{ id: "aeshan-mch-shift", source: "MCH", title: "MCH Shift", allDay: false, start: "2026-02-03T08:00:00", end: "2026-02-03T17:00:00", rawValue: "MCH Shift" }] },
}, typoAliasDb);
seedD1Repository(typoAliasDb, [
  repositoryFile("aeshan-ddh", { sourceType: "ddh", doctors: [{ key: "AESHAN KULURATNE", displayName: "Aeshan KULURATNE", sourceType: "ddh" }] }),
  repositoryFile("zero-only", { sourceType: "ddh", doctors: [{ key: "ZERO PERSON", displayName: "Zero PERSON", sourceType: "ddh" }] }),
]);
const typoDoctors = await postState(typoAliasStore, {
  action: "listUsers",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, typoAliasDb);
const aeshanOption = typoDoctors.availableDoctors.find((doctor) => doctor.displayName === "Aeshan KULARATNE");
assert.ok(aeshanOption, "event-backed typo variants should keep the event-backed display name");
assert.deepEqual(aeshanOption.aliases.map((alias) => alias.key).sort(), ["AESHAN KULARATNE", "AESHAN KULURATNE"]);
assert.equal(typoDoctors.availableDoctors.some((doctor) => doctor.key === "ZERO PERSON"), false, "zero-event standalone identities should be hidden from the picker");
const typoProfile = await postState(typoAliasStore, {
  action: "loadDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: "AESHAN KULARATNE::ddh+mch",
  doctorKey: "AESHAN KULARATNE",
  displayName: "Aeshan KULARATNE",
  sourceTypes: ["ddh", "mch"],
}, typoAliasDb);
assert.deepEqual(typoProfile.snapshot.preview.events.map((event) => event.title), ["MCH Shift"]);
assert.equal(typoProfile.snapshot.profileCoverage.zeroEventAliases.length, 1);
assert.deepEqual(typoProfile.snapshot.profileCoverage.absentSources, []);
const conflictingAliasDb = new MemoryD1();
const conflictingAliasStore = new MemoryStore();
await postState(conflictingAliasStore, { action: "login", email: "rhaydon@gmail.com", password: creatorPassword }, conflictingAliasDb);
for (const [fileId, sourceType, key, displayName] of [
  ["aeshan-a", "mmc", "AESHAN KULARATNE", "Aeshan KULARATNE"],
  ["aeshan-b", "mch", "AESHAN KULURATNE", "Aeshan KULURATNE"],
]) {
  await postState(conflictingAliasStore, {
    action: "saveDerivedCalendarFile",
    email: "rhaydon@gmail.com",
    password: creatorPassword,
    file: { id: fileId, name: `${fileId}.xlsx`, sourceType, active: true },
    doctors: [{ key, displayName, sourceType }],
    eventsByDoctor: { [key]: [{ id: `${fileId}-shift`, source: sourceType.toUpperCase(), title: `${displayName} Shift`, allDay: false, start: "2026-05-05T08:00:00", end: "2026-05-05T17:00:00", rawValue: "Shift" }] },
  }, conflictingAliasDb);
}
const conflictingDoctors = await postState(conflictingAliasStore, {
  action: "listUsers",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, conflictingAliasDb);
assert.equal(conflictingDoctors.availableDoctors.filter((doctor) => doctor.key === "AESHAN KULARATNE" || doctor.key === "AESHAN KULURATNE").length, 2, "overlapping working shifts should block typo merges");
const fourRosterStore = new MemoryStore();
const fourRosterDb = new MemoryD1();
await postState(fourRosterStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, fourRosterDb);
for (const [fileId, sourceType, key, displayName, title, start] of [
  ["creator-mmc-t1", "mmc", "RICHARD HAYDON", "Richard HAYDON", "MMC Term 1", "2026-02-03"],
  ["creator-ddh-t1", "ddh", "RICHARD HAYDON", "Richard Haydon", "DDH Term 1", "2026-02-04"],
  ["creator-mmc-t2", "mmc", "RICHARD HAYDON", "Richard HAYDON", "MMC Term 2", "2026-05-05"],
  ["creator-ddh-t2", "ddh", "HAYDON RICHARD", "HAYDON, Richard", "DDH Term 2", "2026-05-06"],
]) {
  await postState(fourRosterStore, {
    action: "saveDerivedCalendarFile",
    email: "rhaydon@gmail.com",
    password: creatorPassword,
    selectedDoctorKey: "RICHARD HAYDON",
    file: { id: fileId, name: `${fileId}.xlsx`, sourceType, active: true },
    doctors: [{ key, displayName, sourceType }],
    eventsByDoctor: {
      [key]: [{
        id: `${fileId}-shift`,
        source: sourceType.toUpperCase(),
        title,
        allDay: true,
        start,
        end: start,
        rawValue: title,
        monthKey: start.slice(0, 7),
      }],
    },
  }, fourRosterDb);
}
const fourRosterStatus = await postState(fourRosterStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  selectedDoctorKey: "RICHARD HAYDON",
}, fourRosterDb);
assert.equal(fourRosterStatus.total, 4, "all four active roster files should be reported");
assert.equal(fourRosterStatus.files.filter((file) => file.eventCount > 0).length, 4, "each roster file should have D1 event rows");
assert.equal(fourRosterStatus.files.filter((file) => file.selectedDoctorEventCount > 0).length, 4, "each roster file should have selected creator event rows");
const fourRosterCalendar = await postState(fourRosterStore, {
  action: "loadCalendarEvents",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctorKey: "RICHARD HAYDON",
}, fourRosterDb);
assert.deepEqual(
  fourRosterCalendar.snapshot.preview.events.map((event) => event.title).sort(),
  ["DDH Term 1", "DDH Term 2", "MMC Term 1", "MMC Term 2"],
  "creator calendar load should include both hospitals across both terms",
);
assert.equal(fourRosterCalendar.diagnostics?.selectedDoctorFiles, undefined, "default calendar loads should avoid file/doctor diagnostics by default");
const fourRosterDiagnosticCalendar = await postState(fourRosterStore, {
  action: "loadCalendarEvents",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctorKey: "RICHARD HAYDON",
  diagnostics: true,
}, fourRosterDb);
assert.equal(fourRosterDiagnosticCalendar.diagnostics.queryMode, "file-doctor-pairs");
assert.equal(fourRosterDiagnosticCalendar.diagnostics.selectedDoctorFiles.length, 4, "diagnostics should include each resolved file/doctor pair when requested");
const fourRosterExpectedStatus = await postState(fourRosterStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  selectedDoctorKey: "RICHARD HAYDON",
  expectedFileIds: ["creator-mmc-t1", "creator-ddh-t1", "creator-mmc-t2", "creator-ddh-t2"],
}, fourRosterDb);
assert.equal(fourRosterExpectedStatus.expectedFiles.expectedCount, 4);
assert.equal(fourRosterExpectedStatus.expectedFiles.persistedCount, 4);
assert.equal(fourRosterExpectedStatus.expectedFiles.activeCount, 4);
assert.deepEqual(fourRosterExpectedStatus.expectedFiles.missingFileIds, []);
for (let index = 1; index <= 55; index += 1) {
  await postState(fourRosterStore, {
    action: "appendConsoleMessage",
    email: "rhaydon@gmail.com",
    password: creatorPassword,
    message: `Console message ${index}`,
    isError: index === 55,
  }, fourRosterDb);
}
const consoleHistory = await postState(fourRosterStore, {
  action: "consoleMessages",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, fourRosterDb);
assert.equal(consoleHistory.messages.length, 50, "console history should keep only the latest 50 messages");
assert.equal(consoleHistory.messages[0].message, "Console message 55");
assert.equal(consoleHistory.messages[0].isError, true);
assert.equal(consoleHistory.messages.at(-1).message, "Console message 6");

const partialUploadStore = new MemoryStore();
const partialUploadDb = new MemoryD1();
await postState(partialUploadStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, partialUploadDb);
await postState(partialUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  expectedFileIds: ["partial-mmc", "partial-ddh"],
  file: { id: "partial-mmc", name: "partial-mmc.xlsx", sourceType: "mmc", active: true },
  doctors: [{ key: "RICHARD HAYDON", displayName: "Richard HAYDON", sourceType: "mmc" }],
  eventsByDoctor: {
    "RICHARD HAYDON": [{
      id: "partial-mmc-shift",
      source: "MMC",
      title: "Persisted MMC Shift",
      allDay: true,
      start: "2026-02-03",
      end: "2026-02-03",
      rawValue: "Persisted MMC Shift",
      monthKey: "2026-02",
    }],
  },
}, partialUploadDb);
await postState(partialUploadStore, {
  action: "save",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  state: {
    version: 1,
    imports: [
      { repoId: "partial-mmc", id: "partial-mmc", sourceType: "mmc", name: "partial-mmc.xlsx" },
      { repoId: "partial-ddh", id: "partial-ddh", sourceType: "ddh", name: "partial-ddh.xlsx" },
    ],
    session: { doctorKey: "RICHARD HAYDON", settings: {} },
  },
}, partialUploadDb);
const partialUploadStatus = await postState(partialUploadStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  selectedDoctorKey: "RICHARD HAYDON",
  expectedFileIds: ["partial-mmc", "partial-ddh"],
}, partialUploadDb);
assert.equal(partialUploadStatus.expectedFiles.persistedCount, 1, "partial upload status should report only persisted D1 files");
assert.equal(partialUploadStatus.expectedFiles.populatedCount, 1, "partial upload status should count only populated roster files as synced");
assert.deepEqual(partialUploadStatus.expectedFiles.missingFileIds, ["partial-ddh"]);
const partialUploadCalendar = await postState(partialUploadStore, {
  action: "loadCalendarEvents",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctorKey: "RICHARD HAYDON",
}, partialUploadDb);
assert.deepEqual(
  partialUploadCalendar.snapshot.preview.events.map((event) => event.title),
  ["Persisted MMC Shift"],
  "post-login calendar rebuild should only use D1-persisted roster rows",
);
const invalidReplacement = await postStateRaw(partialUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "partial-mmc", name: "partial-mmc.xlsx", sourceType: "mmc", active: true },
  doctors: [{ key: "RICHARD HAYDON", displayName: "Richard HAYDON", sourceType: "mmc" }],
  eventsByDoctor: { "RICHARD HAYDON": [] },
}, partialUploadDb);
assert.equal(invalidReplacement.response.status, 422, "empty derived uploads should be rejected before replacing D1 rows");
assert.equal(
  [...partialUploadDb.events.values()].filter((event) => event.file_id === "partial-mmc").length,
  1,
  "rejected derived uploads must preserve previously indexed D1 events",
);
const sparseReplacement = await postStateRaw(partialUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "partial-mmc", name: "partial-mmc.xlsx", sourceType: "mmc", active: true },
  doctors: [
    { key: "RICHARD HAYDON", displayName: "Richard HAYDON", sourceType: "mmc" },
    { key: "SECOND DOCTOR", displayName: "Second Doctor", sourceType: "mmc" },
  ],
  eventsByDoctor: {
    "RICHARD HAYDON": [{
      id: "partial-mmc-shift",
      source: "MMC",
      title: "Persisted MMC Shift",
      allDay: true,
      start: "2026-02-03",
      end: "2026-02-03",
      rawValue: "Persisted MMC Shift",
      monthKey: "2026-02",
    }],
  },
}, partialUploadDb);
assert.equal(sparseReplacement.response.status, 422, "suspiciously sparse derived uploads should be rejected before replacing D1 rows");
const sharedUploadStore = new MemoryStore();
const sharedUploadDb = new MemoryD1();
await postState(sharedUploadStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, sharedUploadDb);
await seedUser(sharedUploadStore, "shared-user@example.com", "shared-password", "Shared User", sharedUploadDb);
const sharedUserSave = await postStateRaw(sharedUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "shared-user@example.com",
  password: "shared-password",
  file: { id: "shared-ddh", name: "Shared_DDH_04-05-2026_to_02-08-2026.xlsx", sourceType: "ddh", active: true, lastModified: 20 },
  doctors: [{ key: "SHARED USER", displayName: "Shared User", sourceType: "ddh" }],
  eventsByDoctor: {
    "SHARED USER": [{
      id: "shared-ddh-shift",
      source: "DDH",
      title: "Shared DDH Shift",
      allDay: true,
      start: "2026-05-06",
      end: "2026-05-06",
      rawValue: "Shared DDH Shift",
      monthKey: "2026-05",
    }],
  },
}, sharedUploadDb);
assert.equal(sharedUserSave.response.status, 403, "non-creator users must not add roster files to D1");
assert.equal(sharedUploadDb.files.has("shared-ddh"), false);
await postState(sharedUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "shared-ddh", name: "Shared_DDH_04-05-2026_to_02-08-2026.xlsx", sourceType: "ddh", active: true, lastModified: 20 },
  doctors: [{ key: "SHARED USER", displayName: "Shared User", sourceType: "ddh" }],
  eventsByDoctor: {
    "SHARED USER": [{
      id: "shared-ddh-shift", source: "DDH", title: "Shared DDH Shift", allDay: true,
      start: "2026-05-06", end: "2026-05-06", rawValue: "Shared DDH Shift", monthKey: "2026-05",
    }],
  },
}, sharedUploadDb);
assert.equal(sharedUploadDb.files.get("shared-ddh")?.uploaded_by, "rhaydon@gmail.com");
const sharedUserReset = await postStateRaw(sharedUploadStore, {
  action: "resetDerivedCalendarFile",
  email: "shared-user@example.com",
  password: "shared-password",
  fileId: "shared-ddh",
}, sharedUploadDb);
assert.equal(sharedUserReset.response.status, 403, "non-creator users must not remove D1 roster files");
const detailedSharedStatus = await postState(sharedUploadStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  selectedDoctorKey: "SHARED USER",
  lightweight: false,
}, sharedUploadDb);
const sharedDdhStatus = detailedSharedStatus.files.find((file) => file.id === "shared-ddh");
assert.equal(sharedDdhStatus?.selectedDoctorEventCount, 1, "full roster status should count the selected doctor's events per file");
assert.equal(sharedDdhStatus?.selectedDoctor?.displayName, "Shared User", "full roster status should report the matched roster identity");
assert.equal(sharedDdhStatus?.selectedDoctor?.shifts?.[0]?.title, "Shared DDH Shift", "full roster status should return the selected doctor's shifts");
assert.equal(detailedSharedStatus.rosterSourceStatuses?.find((source) => source.id === "monash-adults")?.state, "not-configured", "source status should distinguish an unconnected automation source");
for (const [fileId, name] of [
  ["automation:monash-paeds:failed", "Paeds - Term 1 2026.xlsx"],
  ["automation:monash-adults:queued", "AdultTerm3.2026.xlsx"],
]) {
  sharedUploadDb.rawFiles.set(fileId, {
    file_id: fileId,
    name,
    source_type: fileId.includes("paeds") ? "mch" : "mmc",
    size: 100,
    last_modified: 1,
    object_key: `automation/${fileId}`,
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    data_url: "",
    uploaded_at: "2026-07-29T03:00:00.000Z",
  });
}
sharedUploadDb.rosterSyncRuns.set("failed-paeds", {
  id: "failed-paeds", source_id: "monash-paeds", trigger_type: "sharepoint", provider_version: "4.0",
  content_hash: "failed-hash", file_id: "automation:monash-paeds:failed", status: "failed", message: "Could not parse",
  doctor_count: 0, event_count: 0, started_at: "2026-07-29T02:00:00.000Z", completed_at: "2026-07-29T02:01:00.000Z",
});
sharedUploadDb.rosterSyncRuns.set("queued-adults", {
  id: "queued-adults", source_id: "monash-adults", trigger_type: "sharepoint", provider_version: "19.0",
  content_hash: "queued-hash", file_id: "automation:monash-adults:queued", status: "queued", message: "Queued",
  doctor_count: 0, event_count: 0, started_at: "2026-07-29T03:00:00.000Z", completed_at: "",
});
sharedUploadDb.rosterSources.set("monash-adults", {
  id: "monash-adults", provider: "sharepoint", source_type: "mmc", label: "Monash Adults", enabled: 1,
  config_json: "{}", cursor_json: "{}", provider_version: "19.0", provider_modified_at: "2026-07-29T02:59:00.000Z",
  last_checked_at: "2026-07-29T03:00:00.000Z", last_success_at: "2026-07-29T02:30:00.000Z", last_error: "",
  active_file_id: "", created_at: "2026-07-29T01:00:00.000Z", updated_at: "2026-07-29T03:00:00.000Z",
});
const matchingProviderVersion = await findRosterSyncByProviderVersion(
  sharedUploadDb,
  "monash-adults",
  "19.0",
  "AdultTerm3.2026.xlsx",
);
assert.equal(matchingProviderVersion?.id, "queued-adults", "the same SharePoint file version should resolve to its existing queue run");
const automatedPendingStatus = await postState(sharedUploadStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  selectedDoctorKey: "SHARED USER",
  lightweight: false,
}, sharedUploadDb);
assert.equal(
  automatedPendingStatus.files.some((file) => String(file.id).startsWith("automation:")),
  false,
  "queued and failed automation payloads should not inflate the active roster-file total",
);
assert.equal(
  automatedPendingStatus.rosterSourceStatuses.find((source) => source.id === "monash-adults")?.state,
  "queued",
  "an update waiting for GitHub should be reported as queued even when an older version imported successfully",
);
for (const [fileId, name, lastModified, title] of [
  ["supersede-old", "MMC_Term2_2026_old.xlsx", 10, "Old MMC Shift"],
  ["supersede-new", "MMC_Term2_2026_new.xlsx", 30, "New MMC Shift"],
]) {
  await postState(sharedUploadStore, {
    action: "saveDerivedCalendarFile",
    email: "rhaydon@gmail.com",
    password: creatorPassword,
    file: { id: fileId, name, sourceType: "mmc", active: true, lastModified },
    doctors: [{ key: "SHARED USER", displayName: "Shared User", sourceType: "mmc" }],
    eventsByDoctor: {
      "SHARED USER": [{
        id: `${fileId}-shift`,
        source: "MMC",
        title,
        allDay: true,
        start: "2026-05-07",
        end: "2026-05-07",
        rawValue: title,
        monthKey: "2026-05",
      }],
    },
  }, sharedUploadDb);
}
assert.equal(sharedUploadDb.files.has("supersede-old"), false, "older overlapping same-source roster should be removed after its replacement completes");
assert.equal(
  [...sharedUploadDb.events.values()].some((event) => event.file_id === "supersede-old"),
  false,
  "superseded roster events should be removed with their inactive derived file",
);
assert.equal(sharedUploadDb.files.get("supersede-new")?.active, 1, "latest overlapping same-source roster should remain active");
for (const [fileId, name, sourceId, title] of [
  ["ddh-manual-term", "Dandenong_Emergency_Doctors_Roster_03-08-2026_to_01-11-2026.xlsx", "", "Manual DDH"],
  ["ddh-findmyshift-term", "Dandenong-FindMyShift-stream-paired-v1-2026-08-03-to-2026-11-01.xlsx", "dandenong-findmyshift", "Automated DDH"],
]) {
  await postState(sharedUploadStore, {
    action: "saveDerivedCalendarFile",
    email: "rhaydon@gmail.com",
    password: creatorPassword,
    file: { id: fileId, name, sourceType: "ddh", sourceId, active: true },
    doctors: [{ key: "DDH DOCTOR", displayName: "DDH Doctor", sourceType: "ddh" }],
    eventsByDoctor: {
      "DDH DOCTOR": [{
        id: `${fileId}-shift`,
        source: "DDH",
        title,
        allDay: false,
        start: "2026-08-03T08:00:00+10:00",
        end: "2026-11-01T17:00:00+11:00",
        rawValue: title,
        monthKey: "2026-08",
      }],
    },
  }, sharedUploadDb);
}
assert.equal(sharedUploadDb.files.has("ddh-manual-term"), false, "a FindMyShift source should replace a manual DDH file with identical term coverage");
assert.equal(sharedUploadDb.files.get("ddh-findmyshift-term")?.active, 1, "the automated DDH replacement should remain active");
sharedUploadDb.customEvents.set("history@example.com|historical-note", {
  owner_email: "history@example.com",
  id: "historical-note",
  title: "Historical custom note",
  start_date: "2028-01-10",
  end_date: "2028-01-10",
  all_day: 1,
});
await postState(sharedUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "ddh-partial-manual", name: "Dandenong-manual-2028-01-01-to-2028-02-28.xlsx", sourceType: "ddh", active: true },
  doctors: [{ key: "PARTIAL HISTORY DOCTOR", displayName: "Partial History Doctor", sourceType: "ddh" }],
  eventsByDoctor: {
    "PARTIAL HISTORY DOCTOR": [
      { id: "partial-history", source: "DDH", title: "Historical manual shift", allDay: false, start: "2028-01-10T08:00:00+11:00", end: "2028-01-10T17:00:00+11:00", rawValue: "Historical manual shift", monthKey: "2028-01" },
      { id: "partial-overlap", source: "DDH", title: "Superseded manual shift", allDay: false, start: "2028-02-10T08:00:00+11:00", end: "2028-02-10T17:00:00+11:00", rawValue: "Superseded manual shift", monthKey: "2028-02" },
    ],
  },
}, sharedUploadDb);
await postState(sharedUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "ddh-partial-automatic", name: "Dandenong-FindMyShift-2028-02-01-to-2028-02-28.xlsx", sourceType: "ddh", sourceId: "dandenong-findmyshift-partial", active: true },
  doctors: [{ key: "PARTIAL HISTORY DOCTOR", displayName: "Partial History Doctor", sourceType: "ddh" }],
  eventsByDoctor: {
    "PARTIAL HISTORY DOCTOR": [
      { id: "partial-authoritative", source: "DDH", title: "Latest synced shift", allDay: false, start: "2028-02-10T09:00:00+11:00", end: "2028-02-10T18:00:00+11:00", rawValue: "Latest synced shift", monthKey: "2028-02" },
    ],
  },
}, sharedUploadDb);
assert.equal(sharedUploadDb.files.has("ddh-partial-manual"), true, "a partially overlapping manual roster should remain as historical storage");
assert.ok([...sharedUploadDb.events.values()].some((event) => event.title === "Historical manual shift"), "manual events before the synced window should remain historical");
assert.equal([...sharedUploadDb.events.values()].some((event) => event.title === "Superseded manual shift"), false, "the latest synced window should replace overlapping manual shifts");
assert.ok([...sharedUploadDb.events.values()].some((event) => event.title === "Latest synced shift"), "the latest synced roster should remain authoritative inside its window");
assert.equal(sharedUploadDb.customEvents.has("history@example.com|historical-note"), true, "roster supersession must not remove account-owned custom events");
await postState(sharedUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: {
    id: "findmyshift-window-old",
    name: "Dandenong-FindMyShift-2026-07-01-to-2026-08-07.xlsx",
    sourceType: "ddh",
    sourceId: "dandenong-findmyshift-window",
    active: true,
    uploadedAt: "2026-08-07T00:00:00.000Z",
  },
  doctors: [{ key: "WINDOW DOCTOR", displayName: "Window Doctor", sourceType: "ddh" }],
  eventsByDoctor: {
    "WINDOW DOCTOR": [
      { id: "window-history", source: "DDH", title: "Historical truth", allDay: false, start: "2026-07-01T08:00:00+10:00", end: "2026-07-01T17:00:00+10:00", rawValue: "Historical truth", monthKey: "2026-07" },
      { id: "window-old-overlap", source: "DDH", title: "Old overlapping value", allDay: false, start: "2026-08-07T08:00:00+10:00", end: "2026-08-07T17:00:00+10:00", rawValue: "Old overlapping value", monthKey: "2026-08" },
      { id: "window-old-now-absent", source: "DDH", title: "Old row absent from new truth", allDay: false, start: "2026-09-01T08:00:00+10:00", end: "2026-09-01T17:00:00+10:00", rawValue: "Old row absent from new truth", monthKey: "2026-09" },
    ],
  },
}, sharedUploadDb);
await postState(sharedUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: {
    id: "findmyshift-window-new",
    name: "Dandenong-FindMyShift-2026-08-07-to-2026-11-01.xlsx",
    sourceType: "ddh",
    sourceId: "dandenong-findmyshift-window",
    active: true,
    uploadedAt: "2026-08-08T00:00:00.000Z",
  },
  doctors: [{ key: "WINDOW DOCTOR", displayName: "Window Doctor", sourceType: "ddh" }],
  eventsByDoctor: {
    "WINDOW DOCTOR": [{ id: "window-new-overlap", source: "DDH", title: "New authoritative value", allDay: false, start: "2026-08-07T08:00:00+10:00", end: "2026-08-07T17:00:00+10:00", rawValue: "New authoritative value", monthKey: "2026-08" }],
  },
}, sharedUploadDb);
const rollingWindowEvents = [...sharedUploadDb.events.values()].filter((event) => event.doctor_key === "WINDOW DOCTOR");
assert.ok(rollingWindowEvents.some((event) => event.title === "Historical truth"), "older automated rows outside the new provider window should remain durable truth");
assert.equal(rollingWindowEvents.some((event) => event.title === "Old overlapping value"), false, "a newer automatic snapshot should remove overlapping rows from the older snapshot");
assert.equal(rollingWindowEvents.some((event) => event.title === "Old row absent from new truth"), false, "the declared provider window should remove stale old rows even when the new report has no replacement shift on that date");
assert.equal(rollingWindowEvents.filter((event) => event.start_date === "2026-08-07").length, 1, "rolling automatic snapshots should leave exactly one authoritative row per covered shift");
for (const [fileId, name, lastModified, start, end] of [
  ["adjacent-term-1", "MMC_Term1_2026.xlsx", 10, "2026-05-03", "2026-05-04"],
  ["adjacent-term-2", "MMC_Term2_2026.xlsx", 30, "2026-05-04", "2026-05-04"],
]) {
  await postState(sharedUploadStore, {
    action: "saveDerivedCalendarFile",
    email: "rhaydon@gmail.com",
    password: creatorPassword,
    file: { id: fileId, name, sourceType: "mmc", active: true, lastModified },
    doctors: [{ key: "ADJACENT DOCTOR", displayName: "Adjacent Doctor", sourceType: "mmc" }],
    eventsByDoctor: {
      "ADJACENT DOCTOR": [{
        id: `${fileId}-shift`,
        source: "MMC",
        title: fileId,
        allDay: false,
        start,
        end,
        rawValue: fileId,
        monthKey: start.slice(0, 7),
      }, ...(fileId === "adjacent-term-1" ? [{
        id: `${fileId}-early-shift`,
        source: "MMC",
        title: `${fileId}-early`,
        allDay: false,
        start: "2026-02-02",
        end: "2026-02-02",
        rawValue: `${fileId}-early`,
        monthKey: "2026-02",
      }] : [])],
    },
  }, sharedUploadDb);
}
assert.equal(sharedUploadDb.files.get("adjacent-term-1")?.active, 1, "adjacent terms should remain active when the earlier roster only ends on the next term boundary");
assert.equal(sharedUploadDb.files.get("adjacent-term-2")?.active, 1, "adjacent next-term roster should remain active");
await postState(sharedUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "ambiguous-left", name: "Ambiguous_Left.xlsx", sourceType: "mch", active: true, lastModified: 0 },
  doctors: [{ key: "SHARED USER", displayName: "Shared User", sourceType: "mch" }],
  eventsByDoctor: {
    "SHARED USER": [{ id: "ambiguous-left-shift", source: "MCH", title: "Ambiguous Left", allDay: true, start: "2026-05-08", end: "2026-05-08", rawValue: "Ambiguous Left", monthKey: "2026-05" }],
  },
}, sharedUploadDb);
await postState(sharedUploadStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "ambiguous-right", name: "Ambiguous_Right.xlsx", sourceType: "mch", active: true, lastModified: 0 },
  doctors: [{ key: "SHARED USER", displayName: "Shared User", sourceType: "mch" }],
  eventsByDoctor: {
    "SHARED USER": [{ id: "ambiguous-right-shift", source: "MCH", title: "Ambiguous Right", allDay: true, start: "2026-05-08", end: "2026-05-08", rawValue: "Ambiguous Right", monthKey: "2026-05" }],
  },
}, sharedUploadDb);
assert.ok(memoryD1AccountRecord(sharedUploadDb, "rhaydon@gmail.com").adminIssues.some((issue) => issue.message.includes("Could not determine the latest MCH roster")), "ambiguous supersession should create a creator admin issue");
const conferenceLeaveStore = new MemoryStore();
const conferenceLeaveDb = new MemoryD1();
const conferenceDoctor = { key: "CONFERENCE DOCTOR", displayName: "Conference Doctor", sourceType: "mmc" };
await seedRepository(conferenceLeaveStore, [
  repositoryFile("conference-mmc", { sourceType: "mmc", doctors: [conferenceDoctor] }),
  repositoryFile("conference-mch", { sourceType: "mch", doctors: [{ ...conferenceDoctor, sourceType: "mch" }] }),
]);
await postState(conferenceLeaveStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, conferenceLeaveDb);
await postState(conferenceLeaveStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "conference-mmc", name: "conference-mmc.xlsx", sourceType: "mmc", active: true },
  doctors: [{ ...conferenceDoctor, sourceType: "mmc" }],
  eventsByDoctor: {
    [conferenceDoctor.key]: [{
      id: "conference-mmc-leave",
      source: "MMC",
      title: "Conference Leave",
      allDay: true,
      start: "2026-05-04",
      end: "2026-05-11",
      rawValue: "Conference Leave",
      monthKey: "2026-05",
    }, {
      id: "conference-mmc-separate-leave",
      source: "MMC",
      title: "Conference Leave",
      allDay: true,
      start: "2026-05-18",
      end: "2026-05-25",
      rawValue: "Conference Leave",
      monthKey: "2026-05",
    }],
  },
}, conferenceLeaveDb);
await postState(conferenceLeaveStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "conference-mch", name: "conference-mch.xlsx", sourceType: "mch", active: true },
  doctors: [{ ...conferenceDoctor, sourceType: "mch" }],
  eventsByDoctor: {
    [conferenceDoctor.key]: [{
      id: "conference-mch-leave",
      source: "MCH",
      title: "CME Leave",
      allDay: true,
      start: "2026-05-04",
      end: "2026-05-11",
      rawValue: "CME/L",
      monthKey: "2026-05",
    }],
  },
}, conferenceLeaveDb);
const conferenceProfile = await postState(conferenceLeaveStore, {
  action: "loadDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: `${conferenceDoctor.key}::mmc::mch`,
  doctorKey: conferenceDoctor.key,
  displayName: conferenceDoctor.displayName,
  sourceTypes: ["mmc", "mch"],
}, conferenceLeaveDb);
const conferenceLeaves = conferenceProfile.snapshot.preview.events.filter((event) => event.title === "Conference Leave");
assert.equal(conferenceLeaves.length, 2, "overlapping conference/CME leave should merge, separate weeks should remain separate");
const overlappingConferenceLeave = conferenceLeaves.find((event) => event.start === "2026-05-04");
assert.ok(overlappingConferenceLeave);
assert.equal(overlappingConferenceLeave.end, "2026-05-11");
assert.deepEqual(overlappingConferenceLeave.sources.sort(), ["MCH", "MMC"]);
assert.equal(overlappingConferenceLeave.rawValue, "Conference Leave / CME/L");
const d1FeedResponse = await handleFeedGet({
  request: new Request(`http://fixture.test/api/feed?token=${d1DirectLogin.subscription.token}`),
  env: { ROSTER_DB: d1Store },
});
assert.equal(d1FeedResponse.ok, true);
const d1FeedText = await d1FeedResponse.text();
assert.ok(d1FeedText.includes("BEGIN:VCALENDAR"));
assert.ok(d1FeedText.includes("BEGIN:VEVENT"));
assert.ok(d1FeedText.includes("D1 Edited Shift"), "D1 subscription feed should apply D1 session overrides");
assert.ok(d1FeedText.includes("D1 Custom Event"), "D1 subscription feed should include D1 custom events");
await d1StateStore.delete(`subscription:token:${d1DirectLogin.subscription.token}`);
await d1StateStore.delete("snapshot:account:d1-user@example.com");
const d1OnlyFeedResponse = await handleFeedGet({
  request: new Request(`http://fixture.test/api/feed?token=${d1DirectLogin.subscription.token}&view=range`),
  env: { ROSTER_DB: d1Store },
});
assert.equal(d1OnlyFeedResponse.ok, true);
const d1OnlyFeedText = await d1OnlyFeedResponse.text();
assert.ok(d1OnlyFeedText.includes("D1 Custom Event"), "D1 feed should resolve account and session without KV token index or snapshot");
const d1NoKvLogin = await postState(null, {
  action: "login",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(d1NoKvLogin.snapshot?.preview?.derivedFromD1, true, "D1-only login should return a D1-derived snapshot");
const d1NoKvCalendar = await postState(null, {
  action: "loadCalendarEvents",
  email: "d1-user@example.com",
  password: "d1-password",
}, d1Store);
assert.equal(d1NoKvCalendar.snapshot?.preview?.derivedFromD1, true, "separate calendar load should work without KV when D1 has account and roster data");
const d1NoKvFeedResponse = await handleFeedGet({
  request: new Request(`http://fixture.test/api/feed?token=${d1DirectLogin.subscription.token}`),
  env: { ROSTER_DB: d1Store },
});
assert.equal(d1NoKvFeedResponse.ok, true, "subscription feed should resolve from D1 without KV");

const michaelStateStore = new MemoryStore();
michaelStateStore.d1 = new MemoryD1();
await postState(michaelStateStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
seedD1Repository(michaelStateStore.d1, [
  repositoryFile("michael-mmc", {
    name: "michael-mmc.xlsx",
    sourceType: "mmc",
    doctors: [{ key: "MICHAEL COMAN", displayName: "Michael COMAN", sourceType: "mmc" }],
  }),
  repositoryFile("michael-mch", {
    name: "michael-mch.xlsx",
    sourceType: "mch",
    doctors: [{ key: "DR MICHAEL COMAN", displayName: "Dr Michael Coman", sourceType: "mch" }],
  }),
]);
await seedUser(michaelStateStore, "michael@example.com", "michael-password", "Michael COMAN");
await postState(michaelStateStore, {
  action: "claimRosterName",
  email: "michael@example.com",
  password: "michael-password",
  claim: { sourceType: "mmc", key: "MICHAEL COMAN" },
});
const michaelDirectLogin = await postState(michaelStateStore, {
  action: "login",
  email: "michael@example.com",
  password: "michael-password",
});
const michaelEnrichedLogin = await postState(michaelStateStore, {
  action: "resolveAccountClaims",
  email: "michael@example.com",
  password: "michael-password",
});
assert.deepEqual(michaelEnrichedLogin.state.imports.map((item) => item.repoId).sort(), ["michael-mch", "michael-mmc"]);
assert.equal(michaelEnrichedLogin.claims.some((claim) => claim.sourceType === "mch" && claim.key === "DR MICHAEL COMAN"), true);
assert.equal(michaelEnrichedLogin.suggestedClaims.some((claim) => claim.sourceType === "mch" && claim.key === "DR MICHAEL COMAN"), false);
await postState(michaelStateStore, {
  action: "setAccountRosterClaims",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "michael@example.com",
  claims: [
    { sourceType: "mmc", key: "MICHAEL COMAN" },
    { sourceType: "mch", key: "DR MICHAEL COMAN" },
  ],
});
const michaelAdminLoad = await postState(michaelStateStore, {
  action: "adminLoadUser",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "michael@example.com",
});
assert.deepEqual(
  michaelAdminLoad.state.imports.map((item) => item.repoId).sort(),
  ["michael-mch", "michael-mmc"],
  "admin account loads should include the target account roster refs for the inline snapshot fast path",
);
const michaelPrimaryResolution = await postState(michaelStateStore, {
  action: "resolveDoctorAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctor: {
    key: "MICHAEL COMAN",
    displayName: "Michael COMAN",
    sourceTypes: ["mmc", "mch"],
    aliases: [
      { sourceType: "mmc", key: "MICHAEL COMAN", displayName: "Michael COMAN" },
      { sourceType: "mch", key: "DR MICHAEL COMAN", displayName: "Dr Michael Coman" },
    ],
  },
});
assert.equal(michaelPrimaryResolution.mode, "claimed-account");
assert.equal(michaelPrimaryResolution.email, "michael@example.com");
const michaelAliasResolution = await postState(michaelStateStore, {
  action: "resolveDoctorAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctor: {
    key: "DR MICHAEL COMAN",
    displayName: "Dr Michael Coman",
    sourceTypes: ["mch"],
  },
});
assert.equal(michaelAliasResolution.mode, "claimed-account");
assert.equal(michaelAliasResolution.email, "michael@example.com");
await postState(michaelStateStore, {
  action: "deleteAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "michael@example.com",
});
const michaelDeletedResolution = await postState(michaelStateStore, {
  action: "resolveDoctorAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctor: {
    key: "MICHAEL COMAN",
    displayName: "Michael COMAN",
    sourceTypes: ["mmc", "mch"],
    aliases: [
      { sourceType: "mmc", key: "MICHAEL COMAN", displayName: "Michael COMAN" },
      { sourceType: "mch", key: "DR MICHAEL COMAN", displayName: "Dr Michael Coman" },
    ],
  },
});
assert.equal(michaelDeletedResolution.mode, "doctor-profile");
assert.equal(michaelDeletedResolution.email, "");
seedMinimalD1DoctorEvent(michaelStateStore.d1, "michael-mmc", "MICHAEL COMAN", "mmc", "Michael COMAN");
seedMinimalD1DoctorEvent(michaelStateStore.d1, "michael-mch", "DR MICHAEL COMAN", "mch", "Dr Michael Coman");
const michaelDoctorProfile = await postState(michaelStateStore, {
  action: "loadDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: "MICHAEL COMAN::mch+mmc",
  doctorKey: "MICHAEL COMAN",
  displayName: "Michael COMAN",
  sourceTypes: ["mmc", "mch"],
  aliases: [
    { sourceType: "mmc", key: "MICHAEL COMAN", displayName: "Michael COMAN" },
    { sourceType: "mch", key: "DR MICHAEL COMAN", displayName: "Dr Michael Coman" },
  ],
});
assert.deepEqual(
  michaelDoctorProfile.snapshot.fileRefs.map((ref) => ref.id).sort(),
  ["michael-mch", "michael-mmc"],
  "doctor profile snapshots should include every roster file matched by alias keys",
);
const michaelDoctorProfileWithoutBrowserAliases = await postState(michaelStateStore, {
  action: "loadDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: "MICHAEL COMAN::mch+mmc",
  doctorKey: "MICHAEL COMAN",
  displayName: "Michael COMAN",
  sourceTypes: ["mmc", "mch"],
});
assert.deepEqual(
  [...new Set(michaelDoctorProfileWithoutBrowserAliases.snapshot.preview.events.map((event) => event.source))].sort(),
  ["mch", "mmc"],
  "doctor profile snapshots should expand canonical roster aliases even when the browser omits them",
);
await seedUser(michaelStateStore, "michael@example.com", "michael-password-2", "Michael COMAN");
const michaelRecreatedResolution = await postState(michaelStateStore, {
  action: "resolveDoctorAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctor: {
    key: "DR MICHAEL COMAN",
    displayName: "Dr Michael Coman",
    sourceTypes: ["mch"],
  },
});
assert.equal(michaelRecreatedResolution.mode, "claimed-account");
assert.equal(michaelRecreatedResolution.email, "michael@example.com");

const identityStore = new MemoryStore();
identityStore.d1 = new MemoryD1();
await postState(identityStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
seedD1Repository(identityStore.d1, [
  repositoryFile("identity-ddh", {
    sourceType: "ddh",
    doctors: [
      { key: "AARON BADWAL", displayName: "Aaron BADWAL", sourceType: "ddh" },
      { key: "ANDREA LIM", displayName: "Andrea LIM", sourceType: "ddh" },
      { key: "ABI THANIKASALAM", displayName: "Abi THANIKASALAM", sourceType: "ddh" },
      { key: "JOSEPH VU", displayName: "Joseph VU", sourceType: "ddh" },
    ],
  }),
  repositoryFile("identity-mch", {
    sourceType: "mch",
    doctors: [{ key: "DR ANDREA LIM", displayName: "Dr Andrea LIM", sourceType: "mch" }],
  }),
]);
const abiAutoClaim = await postState(identityStore, {
  action: "login",
  email: "abi@example.com",
  password: "abi-password",
  mode: "create",
  realName: "Abirama Thanikasalam",
});
assert.deepEqual(abiAutoClaim.claims.map((claim) => `${claim.sourceType}:${claim.key}`), ["ddh:ABI THANIKASALAM"]);
const josephEmailAutoClaim = await postState(identityStore, {
  action: "login",
  email: "joseph.vu@monashhealth.org",
  password: "joseph-password",
  mode: "create",
  realName: "jjj",
});
assert.deepEqual(
  josephEmailAutoClaim.claims.map((claim) => `${claim.sourceType}:${claim.key}`),
  ["ddh:JOSEPH VU"],
  "an exact firstname.surname email should link the matching clinician even when the entered name is unrelated",
);
await seedUser(identityStore, "unrelated@example.com", "unrelated-password", "Unrelated Person");
const duplicateJosephClaim = await postStateRaw(identityStore, {
  action: "claimRosterName",
  email: "unrelated@example.com",
  password: "unrelated-password",
  claim: { sourceType: "ddh", key: "JOSEPH VU" },
});
assert.equal(duplicateJosephClaim.response.status, 409);
assert.equal(duplicateJosephClaim.body.conflict, true, "a second account must not be able to claim an existing clinician identity");
const andreaLogin = await postState(identityStore, {
  action: "login",
  email: "andrea@example.com",
  password: "andrea-password",
  mode: "create",
  realName: "Andrea LIM",
});
const andreaEnrichedLogin = await postState(identityStore, {
  action: "resolveAccountClaims",
  email: "andrea@example.com",
  password: "andrea-password",
});
assert.deepEqual(andreaLogin.claims.map((claim) => `${claim.sourceType}:${claim.key}`).sort(), ["ddh:ANDREA LIM", "mch:DR ANDREA LIM"]);
assert.deepEqual(andreaEnrichedLogin.state.imports.map((item) => item.repoId).sort(), ["identity-ddh", "identity-mch"]);
assert.deepEqual(andreaEnrichedLogin.suggestedClaims, []);
const barryLogin = await postState(identityStore, {
  action: "login",
  email: "barry@example.com",
  password: "barry-password",
  mode: "create",
  realName: "Barry Cunningham",
});
assert.deepEqual(barryLogin.claims, []);
assert.deepEqual(barryLogin.suggestedClaims, []);
assert.deepEqual(barryLogin.state.imports, []);
await postState(identityStore, {
  action: "setAccountRosterClaims",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "barry@example.com",
  claims: [{ sourceType: "ddh", key: "AARON BADWAL" }],
});
const aaronAfterBadBarryClaim = await postState(identityStore, {
  action: "resolveDoctorAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  doctor: { sourceType: "ddh", key: "AARON BADWAL", displayName: "Aaron BADWAL" },
});
assert.equal(aaronAfterBadBarryClaim.mode, "claimed-account");
const usersAfterBadBarryClaim = await postState(identityStore, {
  action: "listUsers",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.deepEqual(usersAfterBadBarryClaim.users.find((user) => user.email === "barry@example.com")?.claims.map((claim) => `${claim.sourceType}:${claim.key}`) || [], ["ddh:AARON BADWAL"]);
const barryAfterBadClaim = await postState(identityStore, {
  action: "login",
  email: "barry@example.com",
  password: "barry-password",
});
const barryAfterBadClaimEnriched = await postState(identityStore, {
  action: "resolveAccountClaims",
  email: "barry@example.com",
  password: "barry-password",
});
assert.deepEqual(barryAfterBadClaim.claims.map((claim) => `${claim.sourceType}:${claim.key}`), ["ddh:AARON BADWAL"]);
assert.deepEqual(barryAfterBadClaimEnriched.state.imports.map((item) => item.repoId), ["identity-ddh"]);
await postState(identityStore, {
  action: "claimRosterName",
  email: "barry@example.com",
  password: "barry-password",
  claim: { sourceType: "ddh", key: "AARON BADWAL" },
});
assert.ok(
  memoryD1AccountRecord(identityStore.d1, "barry@example.com").adminIssues.some((issue) => issue.rawValue.includes("Manual roster claim review")),
  "mismatched manual claims should create a Creator review issue",
);
await postState(identityStore, {
  action: "setAccountRosterClaims",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "andrea@example.com",
  claims: [
    { sourceType: "ddh", key: "ANDREA LIM" },
    { sourceType: "mch", key: "DR ANDREA LIM" },
  ],
});
const andreaAssigned = await postState(identityStore, {
  action: "adminLoadUser",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "andrea@example.com",
});
assert.deepEqual(andreaAssigned.claims.map((claim) => `${claim.sourceType}:${claim.key}`).sort(), ["ddh:ANDREA LIM", "mch:DR ANDREA LIM"]);
assert.deepEqual(
  andreaAssigned.state.imports.map((item) => item.repoId).sort(),
  ["identity-ddh", "identity-mch"],
  "admin account loads should include roster refs needed for the inline snapshot fast path",
);
await postState(identityStore, {
  action: "removeRosterClaim",
  email: "andrea@example.com",
  password: "andrea-password",
  claim: { sourceType: "ddh", key: "ANDREA LIM" },
});
await postState(identityStore, {
  action: "reportRosterIdentityIssue",
  email: "andrea@example.com",
  password: "andrea-password",
  message: "Wrong roster name.",
});
const andreaProfileAfterReport = identityStore.d1.accountProfiles.get("andrea@example.com");
const andreaRecordAfterReport = {
  claims: [...identityStore.d1.accountClaims.values()].filter((claim) => claim.email === "andrea@example.com").map((claim) => ({
    sourceType: claim.source_type,
    key: claim.doctor_key,
  })),
  adminIssues: JSON.parse(andreaProfileAfterReport.admin_issues_json || "[]"),
};
assert.equal((andreaRecordAfterReport.claims || []).some((claim) => claim.sourceType === "ddh"), false);
assert.ok(andreaRecordAfterReport.adminIssues.length >= 1);

const manyDoctorsStore = new MemoryStore();
manyDoctorsStore.d1 = new MemoryD1();
await postState(manyDoctorsStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
seedD1Repository(manyDoctorsStore.d1, [
  repositoryFile("many-doctors", {
    doctors: Array.from({ length: 90 }, (_, index) => ({
      key: `DOCTOR ${index}`,
      displayName: `Doctor ${index}`,
      sourceType: "mmc",
    })),
  }),
]);
await seedUser(manyDoctorsStore, "doctor-1@example.com", "doctor-password", "Doctor 1");
await seedUser(manyDoctorsStore, "doctor-2@example.com", "doctor-password", "Doctor 2");
manyDoctorsStore.resetMetrics();
const manyDoctorsLogin = await postState(manyDoctorsStore, {
  action: "login",
  email: "new-doctor@example.com",
  password: "new-password",
  mode: "create",
  realName: "New Doctor",
});
const manyDoctorsEnrichment = await postState(manyDoctorsStore, {
  action: "resolveAccountClaims",
  email: "new-doctor@example.com",
  password: "new-password",
});
assert.equal(manyDoctorsEnrichment.availableDoctors.length, 90);
assert.ok(manyDoctorsStore.accountListCalls <= 2, "available doctor claimed status should avoid repeated account scans");

const profileImports = await postStateRaw(stateStore, {
  action: "loadDoctorProfileImports",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: "TITUS HACKMAN::mmc",
  doctorKey: "TITUS HACKMAN",
  displayName: "Titus HACKMAN",
  sourceTypes: ["mmc"],
});
assert.equal(profileImports.response.status, 410);

await postState(stateStore, {
  action: "saveDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: "TITUS HACKMAN::mmc",
  doctorKey: "TITUS HACKMAN",
  displayName: "Titus HACKMAN",
  sourceTypes: ["mmc"],
  state: { version: 1, imports: [], session: { hadPreview: true } },
  snapshot: {
    preview: {
      count: 1,
      date_range: "2026-07-27 to 2026-07-28",
      events: [{ id: "old-leave", source: "Casey", title: "Annual Leave", allDay: true, start: "2026-07-27", end: "2026-07-28", rawValue: "Annual Leave" }],
      review: [],
      issues: [],
      conflicts: [],
    },
    session: { hadPreview: true },
    doctorOptions: [{ key: "TITUS HACKMAN", displayName: "Titus HACKMAN", sourceTypes: ["mmc"] }],
    detectedSources: { mmc: ["fixture-roster"] },
    fileRefs: [{ repoId: "fixture-roster", id: "fixture-roster", sourceType: "mmc", name: "AdultMMCTerm2.2026.Ver1.pdf" }],
  },
});
const d1ProfileReload = await postState(stateStore, {
  action: "loadDoctorProfile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: "TITUS HACKMAN::mmc",
  doctorKey: "TITUS HACKMAN",
  displayName: "Titus HACKMAN",
  sourceTypes: ["mmc"],
});
assert.equal(d1ProfileReload.profile.profileId, "TITUS HACKMAN::mmc");

await seedUser(stateStore, "patrick@example.com", "patrick-password", "Patrick TAN");
await seedUser(stateStore, "senior@example.com", "senior-password", "Senior Registrar");
const initialUsers = await postState(stateStore, {
  action: "listUsers",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.equal(initialUsers.users.find((user) => user.email === "patrick@example.com")?.insightsEnabled, false);
const enabledInsights = await postState(stateStore, {
  action: "setUserInsightsEnabled",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "patrick@example.com",
  insightsEnabled: true,
});
assert.equal(enabledInsights.user.insightsEnabled, true);
const patrickLogin = await postState(stateStore, {
  action: "login",
  email: "patrick@example.com",
  password: "patrick-password",
});
assert.equal(patrickLogin.insightsEnabled, true);
const seniorLogin = await postState(stateStore, {
  action: "login",
  email: "senior@example.com",
  password: "senior-password",
});
assert.equal(seniorLogin.insightsEnabled, false);
stateStore.d1.accountClaims.set("patrick@example.com|mmc|PATRICK TAN", {
  email: "patrick@example.com",
  source_type: "mmc",
  doctor_key: "PATRICK TAN",
  display_name: "Patrick TAN",
  matched_at: "2026-05-01T00:00:00.000Z",
  updated_at: "2026-05-01T00:00:00.000Z",
});
const writerCodeIssue = {
  id: "writer-code",
  source: "MMC",
  seniority: "Senior Registrar",
  startDay: "2026-05-02",
  date: "2026-05-02",
  rawValue: "WRITER",
  code: "WRITER",
  status: "unknown",
  message: "MMC shift code not recognised.",
  resolutionType: "shift_code",
  fingerprint: "MMC::Senior Registrar::WRITER",
};
await postState(stateStore, {
  action: "saveDerivedCalendarFile",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  file: { id: "writer-code-file", name: "Writer Code.xlsx", sourceType: "mmc", size: 1, lastModified: 1 },
  doctors: [{ key: "PATRICK TAN", displayName: "Patrick TAN", sourceType: "mmc" }],
  eventsByDoctor: {
    "PATRICK TAN": [{
      id: "writer-code-event",
      source: "MMC",
      seniority: "Senior Registrar",
      title: "MMC: WRITER",
      allDay: true,
      start: "2026-05-02",
      end: "2026-05-03",
      rawValue: "WRITER",
    }],
  },
  issuesByDoctor: { "PATRICK TAN": [writerCodeIssue] },
  skipStatus: true,
});
assert.ok(memoryD1AccountRecord(stateStore.d1, "patrick@example.com").adminIssues.some((issue) => issue.code === "WRITER"), "D1 ingestion should promote unresolved shift-code diagnostics into Creator Errors");
const patrickWriterCalendar = await postState(stateStore, {
  action: "loadCalendarEvents",
  email: "patrick@example.com",
  password: "patrick-password",
  doctorKey: "PATRICK TAN",
});
assert.ok(patrickWriterCalendar.snapshot.preview.issues.some((issue) => issue.rawValue === "WRITER"), "D1-derived user calendar should still expose the warning panel issue");
await postState(stateStore, {
  action: "reportUserError",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "senior@example.com",
  errorId: writerCodeIssue.fingerprint,
  message: writerCodeIssue.message,
  issue: writerCodeIssue,
});
await postState(stateStore, {
  action: "saveLocalParserExtensionRule",
  email: "patrick@example.com",
  password: "patrick-password",
  fingerprint: writerCodeIssue.fingerprint,
  rawValue: "WRITER",
  rule: {
    source: "MMC",
    seniority: "Senior Registrar",
    code: "WRITER",
    kind: "ignore",
    ignore: true,
    includeAsShift: false,
  },
});
assert.equal(memoryD1AccountRecord(stateStore.d1, "patrick@example.com").adminIssues.some((issue) => issue.code === "WRITER"), false, "local ignored shift code should clear that user's warning evidence");
assert.equal(memoryD1AccountRecord(stateStore.d1, "senior@example.com").adminIssues.some((issue) => issue.code === "WRITER"), true, "local ignored shift code should not clear other users");
await postState(stateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  source: "MMC",
  rawValue: "WRITER",
  rule: {
    source: "MMC",
    seniority: "Senior Registrar",
    code: "WRITER",
    kind: "ignore",
    ignore: true,
    includeAsShift: false,
  },
});
assert.equal(memoryD1AccountRecord(stateStore.d1, "senior@example.com").adminIssues.some((issue) => issue.code === "WRITER"), false, "creator ignored shift code should clear matching warnings globally");
const n1Issue = {
  source: "MMC",
  seniority: "Senior Registrar",
  date: "2026-05-01",
  rawValue: "N1",
  message: "MMC shift code not recognised.",
  fingerprint: "MMC::Senior Registrar::N1",
};
await postState(stateStore, {
  action: "reportUserError",
  email: "patrick@example.com",
  password: "patrick-password",
  errorId: n1Issue.fingerprint,
  message: n1Issue.message,
  issue: n1Issue,
});
await postState(stateStore, {
  action: "reportUserError",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "senior@example.com",
  errorId: n1Issue.fingerprint,
  message: n1Issue.message,
  issue: n1Issue,
});
assert.equal(memoryD1AccountRecord(stateStore.d1, "patrick@example.com").adminIssues.length, 1);
assert.equal(memoryD1AccountRecord(stateStore.d1, "senior@example.com").adminIssues.length, 1);
await postState(stateStore, {
  action: "saveLocalParserExtensionRule",
  email: "patrick@example.com",
  password: "patrick-password",
  fingerprint: n1Issue.fingerprint,
  rawValue: "N1",
  rule: {
    source: "MMC",
    seniority: "Senior Registrar",
    code: "N1",
    kind: "shift",
    base: "SR IC",
    period: "NIGHT",
    suffix: "",
    allDay: false,
    startTime: "23:00",
    endTime: "09:00",
    includeAsShift: true,
  },
});
assert.equal(memoryD1AccountRecord(stateStore.d1, "patrick@example.com").adminIssues.length, 0, "local parser rule should clear matching user warning evidence");
assert.equal(memoryD1AccountRecord(stateStore.d1, "senior@example.com").adminIssues.length, 1, "local parser rule should not clear other users");
const creatorSuggestionView = await postState(stateStore, {
  action: "listUsers",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.equal(creatorSuggestionView.issueConfig.parserRuleSuggestions.length, 1, "user shift-code resolutions must be visible to the creator");
await postState(stateStore, {
  action: "decideParserRuleSuggestion",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  suggestionId: creatorSuggestionView.issueConfig.parserRuleSuggestions[0].id,
  decision: "reject",
});
const parserSave = await postState(stateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  fingerprint: n1Issue.fingerprint,
  source: "MMC",
  rawValue: "N1",
  rule: {
    source: "MMC",
    seniority: "Senior Registrar",
    code: "N1",
    kind: "shift",
    base: "SR IC",
    period: "NIGHT",
    suffix: "",
    allDay: false,
    startTime: "23:00",
    endTime: "09:00",
    includeAsShift: true,
  },
});
assert.ok(parserSave.parserExtensions.mmc.some((rule) => rule.seniority === "Senior Registrar" && rule.code === "N1"));
setParserExtensions(parserSave.parserExtensions);
const srN1Workbook = XLSX.utils.book_new();
const srN1Sheet = XLSX.utils.aoa_to_sheet([
  [],
  [],
  [],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "SENIOR REG"],
  ["", "", "", "Patrick TAN", "", "2300-0900 N1"],
]);
for (let index = 0; index < 7; index += 1) {
  srN1Sheet[XLSX.utils.encode_cell({ r: 3, c: 5 + index })] = { t: "d", v: new Date(`2026-05-${String(4 + index).padStart(2, "0")}T00:00:00`) };
}
XLSX.utils.book_append_sheet(srN1Workbook, XLSX.utils.aoa_to_sheet([[]]), "Whole thing");
XLSX.utils.book_append_sheet(srN1Workbook, srN1Sheet, "Week 1");
const srN1View = buildRosterView([{ id: "sr-n1", workbook: srN1Workbook, file: { name: "AdultTerm.xlsx", size: 1, lastModified: 1 } }], [], "PATRICK TAN");
assert.ok(srN1View.events.some((event) => event.rawValue === "2300-0900 N1" && event.title === "MMC: SR IC Night" && event.start.includes("23:00:00") && event.end.includes("09:00:00")), "Senior Registrar N1 explicit-time rules must render with the saved rule title");
assert.equal(srN1View.issues.some((issue) => issue.rawValue === "2300-0900 N1"), false);
assert.match(rosterSource, /add\(rules\.mmc, "MMC", "CS 0\.5", seniority, "CS", "", "", true/, "MMC CS 0.5 should resolve to an all-day CS event");
assert.match(rosterSource, /add\(rules\.mmc, "MMC", "N1", "SMS", "Night shift", "", "", false, "23:00", "08:30"/, "SMS N1 should resolve to a timed Night shift");
assert.equal(memoryD1AccountRecord(stateStore.d1, "patrick@example.com").adminIssues.length, 0, "global parser rule should keep direct-user warning evidence cleared");
assert.equal(memoryD1AccountRecord(stateStore.d1, "senior@example.com").adminIssues.length, 0, "global parser rule should clear matching switch-user warning evidence");
const staleReport = await postState(stateStore, {
  action: "reportUserError",
  email: "patrick@example.com",
  password: "patrick-password",
  errorId: "MMC::Senior Registrar::2200-0830 N1",
  message: n1Issue.message,
  issue: {
    ...n1Issue,
    rawValue: "2200-0830 N1",
    fingerprint: "MMC::Senior Registrar::2200-0830 N1",
  },
});
assert.equal(staleReport.ignored, true, "resolved global shift-code warnings must not be requeued from stale user previews");
assert.equal(memoryD1AccountRecord(stateStore.d1, "patrick@example.com").adminIssues.length, 0, "resolved global shift-code warning evidence must remain cleared");
const ssuBatchSave = await postState(stateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  source: "MMC",
  rawValue: "ASSJ",
  rules: ["HMO", "Intern"].map((seniority) => ({
    source: "MMC",
    seniority,
    code: "ASSJ",
    kind: "shift",
    base: "SSU",
    period: "AM",
    suffix: "",
    allDay: false,
    startTime: "07:30",
    endTime: "17:30",
    includeAsShift: true,
  })),
});
assert.ok(ssuBatchSave.parserExtensions.mmc.some((rule) => rule.seniority === "HMO" && rule.code === "ASSJ"), "batch parser save should add HMO ASSJ");
assert.ok(ssuBatchSave.parserExtensions.mmc.some((rule) => rule.seniority === "Intern" && rule.code === "ASSJ"), "batch parser save should add Intern ASSJ");
assert.equal(ssuBatchSave.parserExtensions.mmc.some((rule) => rule.seniority === "Senior Registrar" && rule.code === "ASSJ"), false, "batch parser save should not add unselected seniorities");
const pssjBatchSave = await postState(stateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  source: "MMC",
  rawValue: "PSSJ",
  rules: ["HMO", "Intern"].map((seniority) => ({
    source: "MMC",
    seniority,
    code: "PSSJ",
    kind: "shift",
    base: "SSU",
    period: "PM",
    suffix: "",
    allDay: false,
    startTime: "14:30",
    endTime: "00:00",
    includeAsShift: true,
  })),
});
setParserExtensions(pssjBatchSave.parserExtensions);
const hmoSsuWorkbook = XLSX.utils.book_new();
const hmoSsuSheet = XLSX.utils.aoa_to_sheet([
  [],
  [],
  [],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "HMO"],
  ["", "", "", "Patrick TAN", "", "0800-1730 ASSJ", "1430-0000 PSSJ"],
]);
for (let index = 0; index < 7; index += 1) {
  hmoSsuSheet[XLSX.utils.encode_cell({ r: 3, c: 5 + index })] = { t: "d", v: new Date(`2026-05-${String(4 + index).padStart(2, "0")}T00:00:00`) };
}
XLSX.utils.book_append_sheet(hmoSsuWorkbook, XLSX.utils.aoa_to_sheet([[]]), "Whole thing");
XLSX.utils.book_append_sheet(hmoSsuWorkbook, hmoSsuSheet, "Week 1");
const hmoSsuView = buildRosterView([{ id: "hmo-ssu", workbook: hmoSsuWorkbook, file: { name: "AdultTerm.xlsx", size: 1, lastModified: 1 } }], [], "PATRICK TAN");
assert.ok(hmoSsuView.events.some((event) => event.rawValue === "0800-1730 ASSJ" && event.title === "MMC: SSU AM" && event.start.includes("08:00:00") && event.end.includes("17:30:00")), "HMO ASSJ should resolve to SSU AM while preserving explicit roster time");
assert.ok(hmoSsuView.events.some((event) => event.rawValue === "1430-0000 PSSJ" && event.title === "MMC: SSU PM" && event.start.includes("14:30:00") && event.end.includes("00:00:00")), "HMO PSSJ should resolve to SSU PM");
assert.equal(hmoSsuView.issues.some((issue) => issue.rawValue.includes("SSJ")), false, "selected seniorities should no longer surface ASSJ/PSSJ as unresolved");
const deletedAssj = await postState(stateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  source: "MMC",
  rawValue: "ASSJ",
  replacementTargets: [
    { source: "MMC", seniority: "HMO", code: "ASSJ" },
    { source: "MMC", seniority: "Intern", code: "ASSJ" },
  ],
  rules: [{
    source: "MMC",
    seniority: "HMO",
    code: "ASSJ",
    kind: "shift",
    base: "SSU",
    period: "AM",
    suffix: "",
    allDay: false,
    startTime: "07:30",
    endTime: "17:30",
    includeAsShift: true,
  }],
});
assert.ok(deletedAssj.parserExtensions.mmc.some((rule) => rule.seniority === "HMO" && rule.code === "ASSJ"), "replacement save should keep selected matching seniority");
assert.equal(deletedAssj.parserExtensions.mmc.some((rule) => rule.seniority === "Intern" && rule.code === "ASSJ"), false, "replacement save should delete deselected matching seniority");
await postState(stateStore, {
  action: "deleteParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  rule: { source: "MMC", seniority: "HMO", code: "ASSJ" },
});
const hmoAssjReappears = await postState(stateStore, {
  action: "reportUserError",
  email: "patrick@example.com",
  password: "patrick-password",
  errorId: "MMC::HMO::0800-1730 ASSJ",
  message: "MMC shift code not recognised.",
  issue: {
    source: "MMC",
    seniority: "HMO",
    date: "2026-05-01",
    rawValue: "0800-1730 ASSJ",
    message: "MMC shift code not recognised.",
    fingerprint: "MMC::HMO::0800-1730 ASSJ",
  },
});
assert.equal(hmoAssjReappears.ignored, undefined, "deleted shift-code disambiguations should allow unresolved reports to reappear");
const knownDdhClinicalSupport = await postState(stateStore, {
  action: "reportUserError",
  email: "patrick@example.com",
  password: "patrick-password",
  errorId: "DDH::Unknown::Clinical Support",
  message: "DDH shift code not recognised; using explicit roster time.",
  issue: {
    source: "DDH",
    seniority: "Unknown",
    date: "2026-03-02",
    rawValue: "Clinical Support",
    message: "DDH shift code not recognised; using explicit roster time.",
    fingerprint: "DDH::Unknown::Clinical Support",
  },
});
assert.equal(knownDdhClinicalSupport.ignored, true, "known DDH Clinical Support mappings should not enter unresolved shift-code queues");
const knownDdhSsuSms = await postState(stateStore, {
  action: "reportUserError",
  email: "patrick@example.com",
  password: "patrick-password",
  errorId: "DDH::Unknown::SSU SMS",
  message: "DDH shift code not recognised; using explicit roster time.",
  issue: {
    source: "DDH",
    seniority: "Unknown",
    date: "2026-03-06",
    rawValue: "SSU SMS",
    message: "DDH shift code not recognised; using explicit roster time.",
    fingerprint: "DDH::Unknown::SSU SMS",
  },
});
assert.equal(knownDdhSsuSms.ignored, true, "known DDH SSU SMS mappings should not enter unresolved shift-code queues");
await seedUser(stateStore, "michael@example.com", "michael-password", "Michael Coman");
const michaelProfile = stateStore.d1.accountProfiles.get("michael@example.com");
michaelProfile.admin_issues_json = JSON.stringify([
  {
    id: "MCH::SMS::OCS",
    source: "MCH",
    seniority: "SMS",
    date: "2026-03-25",
    rawValue: "0800-1730 OCS",
    code: "OCS",
    timeLabel: "08:00-17:30",
    suggestedTitle: "MCH: OCS",
    fingerprint: "MCH::SMS::OCS",
    message: "MCH shift code not recognised; using explicit roster time.",
  },
  {
    id: "MCH::SMS::PHNW",
    source: "MCH",
    seniority: "SMS",
    date: "2026-03-09",
    rawValue: "0800-1730PHNW",
    code: "PHNW",
    timeLabel: "08:00-17:30",
    suggestedTitle: "MCH: PHNW",
    fingerprint: "MCH::SMS::PHNW",
    message: "MCH shift code not recognised; using explicit roster time.",
  },
  {
    id: "MCH::SMS::AM",
    source: "MCH",
    seniority: "SMS",
    date: "2026-02-10",
    rawValue: "0800-1730",
    code: "AM",
    timeLabel: "08:00-17:30",
    suggestedTitle: "MCH: AM",
    fingerprint: "MCH::SMS::AM",
    message: "MCH shift code not recognised; using explicit roster time.",
  },
  {
    id: "MMC::CMO::PM",
    source: "MMC",
    seniority: "CMO",
    date: "2026-02-09",
    rawValue: "1430-0000",
    code: "PM",
    timeLabel: "14:30-00:00",
    suggestedTitle: "MMC: PM",
    fingerprint: "MMC::CMO::PM",
    message: "MMC shift code not recognised; using explicit roster time.",
  },
]);
const cleanedUserList = await postState(stateStore, {
  action: "listUsers",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.equal(cleanedUserList.users.find((user) => user.email === "michael@example.com")?.adminIssues.length, 0, "creator user list should hide stale resolved shift-code issues");
assert.ok(memoryD1AccountRecord(stateStore.d1, "michael@example.com").adminIssues.length > 0, "creator user list should not persist a full stale-issue cleanup during load");
await postState(stateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  source: "DDH",
  rawValue: "Rover AM",
  rules: ["SMS", "CMO"].map((seniority) => ({
    source: "DDH",
    seniority,
    code: "ROVER AM",
    kind: "shift",
    base: "Rover",
    period: "AM",
    suffix: "",
    allDay: false,
    startTime: "08:00",
    endTime: "18:00",
    includeAsShift: true,
  })),
});
const staleUnknownRover = await postState(stateStore, {
  action: "reportUserError",
  email: "patrick@example.com",
  password: "patrick-password",
  errorId: "DDH::Unknown::Rover AM",
  message: "DDH shift label not recognised; using roster time.",
  issue: {
    source: "DDH",
    seniority: "Unknown",
    date: "2026-05-14",
    rawValue: "Rover AM",
    message: "DDH shift label not recognised; using roster time.",
    fingerprint: "DDH::Unknown::Rover AM",
  },
});
assert.equal(staleUnknownRover.ignored, true, "stale Unknown-seniority DDH warnings should resolve once a source/code rule exists");
const accParserSave = await postState(stateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  fingerprint: "MMC::Senior Registrar::ACC",
  source: "MMC",
  rawValue: "ACC",
  rule: {
    source: "MMC",
    seniority: "Senior Registrar",
    code: "ACC",
    kind: "shift",
    base: "Clinic",
    period: "AM",
    suffix: "Charge",
    allDay: false,
    startTime: "08:00",
    endTime: "17:30",
    includeAsShift: true,
  },
});
assert.ok(accParserSave.parserExtensions.mmc.some((rule) => rule.seniority === "Senior Registrar" && rule.code === "ACC"), "Senior Registrar charge/consultant-style codes must persist as explicit rules");
setParserExtensions(accParserSave.parserExtensions);
const srAccWorkbook = XLSX.utils.book_new();
const srAccSheet = XLSX.utils.aoa_to_sheet([
  [],
  [],
  [],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "SENIOR REG"],
  ["", "", "", "Patrick TAN", "", "ACC"],
]);
for (let index = 0; index < 7; index += 1) {
  srAccSheet[XLSX.utils.encode_cell({ r: 3, c: 5 + index })] = { t: "d", v: new Date(`2026-05-${String(4 + index).padStart(2, "0")}T00:00:00`) };
}
XLSX.utils.book_append_sheet(srAccWorkbook, XLSX.utils.aoa_to_sheet([[]]), "Whole thing");
XLSX.utils.book_append_sheet(srAccWorkbook, srAccSheet, "Week 1");
const srAccView = buildRosterView([{ id: "sr-acc", workbook: srAccWorkbook, file: { name: "AdultTerm.xlsx", size: 1, lastModified: 1 } }], [], "PATRICK TAN");
assert.ok(srAccView.events.some((event) => event.rawValue === "ACC" && event.title === "MMC: Clinic AM Charge"), "Senior Registrar explicit ACC rules must render in user calendars");
assert.equal(srAccView.issues.some((issue) => issue.rawValue === "ACC"), false);
setParserExtensions({});

const deletionStore = new MemoryStore();
deletionStore.d1 = new MemoryD1();
await postState(deletionStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
await seedRepository(deletionStore, [
  repositoryFile("keep-roster"),
  repositoryFile("missing-from-save", { name: "missing-from-save.xlsx" }),
  repositoryFile("remove-roster", { name: "remove-roster.xlsx" }),
]);
seedD1Repository(deletionStore.d1, [
  repositoryFile("keep-roster"),
  repositoryFile("missing-from-save", { name: "missing-from-save.xlsx" }),
  repositoryFile("remove-roster", { name: "remove-roster.xlsx" }),
]);

await postState(deletionStore, {
  action: "save",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  state: {
    version: 1,
    imports: [{ repoId: "keep-roster", id: "keep-roster", sourceType: "mmc", name: "keep-roster.xlsx" }],
    session: {},
  },
});
assert.equal(deletionStore.d1.files.has("missing-from-save"), true, "ordinary creator save must not delete omitted D1 roster files");
let deletionIndex = await deletionStore.get("repository:index", "json");
assert.ok(deletionIndex.files.some((file) => file.id === "missing-from-save"), "ordinary creator save must keep omitted files in the repository index");

deletionStore.d1.rawFiles.set("remove-roster", {
  file_id: "remove-roster",
  name: "remove-roster.xlsx",
  source_type: "mmc",
  size: 12,
  last_modified: 1,
  object_key: "rosters/remove-roster",
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  data_url: "",
  uploaded_at: "2026-01-01T00:00:00.000Z",
});
await deletionStore.r2.put("rosters/remove-roster", new Uint8Array([1, 2, 3]));
assert.ok(deletionStore.r2.objects.has("rosters/remove-roster"), "fixture should seed retained R2 source before removal");

await postState(deletionStore, {
  action: "save",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  state: {
    version: 1,
    imports: [{ repoId: "keep-roster", id: "keep-roster", sourceType: "mmc", name: "keep-roster.xlsx" }],
    session: {},
  },
  removedImportIds: ["remove-roster"],
});
assert.ok(await deletionStore.get("repository:file:remove-roster", "json"), "D1-only removal must not mutate legacy KV files");
assert.equal(deletionStore.d1.files.has("remove-roster"), false, "creator removal should delete the D1 roster file");
assert.equal(deletionStore.d1.rawFiles.has("remove-roster"), false, "creator removal should delete raw roster metadata");
assert.ok(!deletionStore.r2.objects.has("rosters/remove-roster"), "creator removal should delete retained R2 source bytes");
const removedRosterStatus = await postState(deletionStore, {
  action: "calendarStoreStatus",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
}, deletionStore.d1);
assert.equal(removedRosterStatus.files.some((file) => file.id === "remove-roster"), false, "calendar status must not list removed roster files");
deletionIndex = await deletionStore.get("repository:index", "json");
assert.equal(deletionIndex.files.some((file) => file.id === "remove-roster"), true, "D1-only removal must not mutate legacy KV index");
assert.ok(await deletionStore.get("repository:file:keep-roster", "json"));
assert.ok(await deletionStore.get("repository:file:missing-from-save", "json"));
await postState(deletionStore, {
  action: "replaceActiveRosterFiles",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  keepFileIds: ["keep-roster"],
  confirmation: "REBUILD",
}, deletionStore.d1);
assert.equal(deletionStore.d1.files.has("keep-roster"), true, "recovery should retain the requested current roster");
assert.equal(deletionStore.d1.files.has("missing-from-save"), false, "recovery should remove active rosters outside the current upload set");

const emptyReplacement = await postStateRaw(deletionStore, {
  action: "replaceActiveRosterFiles",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  keepFileIds: [],
  confirmation: "REBUILD",
}, deletionStore.d1);
assert.equal(emptyReplacement.response.status, 400, "rebuild-all must not accept an empty retained roster set");

const failedReplacementStore = new MemoryStore();
failedReplacementStore.d1 = new MemoryD1();
await postState(failedReplacementStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
seedD1Repository(failedReplacementStore.d1, [
  repositoryFile("last-known-good", { name: "last-known-good.xlsx" }),
]);
await postState(failedReplacementStore, {
  action: "save",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  state: {
    version: 1,
    imports: [{ repoId: "new-not-persisted", id: "new-not-persisted", sourceType: "mmc", name: "new-not-persisted.xlsx" }],
    session: {},
  },
});
assert.equal(
  failedReplacementStore.d1.files.has("last-known-good"),
  true,
  "creator save should keep the last-known-good D1 roster set when the replacement files are not yet persisted",
);

await postState(deletionStore, {
  action: "save",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  state: {
    version: 1,
    imports: [{ repoId: "keep-roster", id: "keep-roster", sourceType: "mmc", name: "keep-roster.xlsx" }],
    session: {},
  },
  removedImportIds: ["remove-roster"],
});
assert.ok(await deletionStore.get("repository:file:keep-roster", "json"));
assert.ok(await deletionStore.get("repository:file:missing-from-save", "json"));

await seedUser(deletionStore, "claimed-doctor@example.com", "claimed-password");
await postState(deletionStore, {
  action: "claimRosterName",
  email: "claimed-doctor@example.com",
  password: "claimed-password",
  claim: { sourceType: "mmc", key: "TITUS HACKMAN" },
});
const observerBeforeDelete = await postState(deletionStore, {
  action: "login",
  email: "observer@example.com",
  password: "observer-password",
  mode: "create",
  realName: "Observer Person",
});
const observerBeforeDeleteEnriched = await postState(deletionStore, {
  action: "resolveAccountClaims",
  email: "observer@example.com",
  password: "observer-password",
});
assert.equal(
  observerBeforeDeleteEnriched.availableDoctors.find((doctor) => doctor.key === "TITUS HACKMAN")?.claimedBy,
  "claimed-doctor@example.com",
  "repository doctor should be marked claimed before deleting the linked account",
);
await postState(deletionStore, {
  action: "deleteAccount",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "claimed-doctor@example.com",
});
assert.equal(await deletionStore.get("account:claimed-doctor@example.com", "json"), null, "deleteAccount must remove the claimed account record");
const observerAfterDelete = await postState(deletionStore, {
  action: "login",
  email: "observer@example.com",
  password: "observer-password",
});
assert.equal(
  observerAfterDelete.availableDoctors.find((doctor) => doctor.key === "TITUS HACKMAN")?.claimedBy || "",
  "",
  "repository doctor should become unclaimed after deleting the linked account",
);

await seedUser(deletionStore, "user@example.com", "user-password");
await postState(deletionStore, {
  action: "save",
  email: "user@example.com",
  password: "user-password",
  state: {
    version: 1,
    imports: [{ repoId: "keep-roster", id: "keep-roster", sourceType: "mmc", name: "keep-roster.xlsx" }],
    session: {},
  },
  removedImportIds: ["keep-roster"],
});
assert.ok(await deletionStore.get("repository:file:keep-roster", "json"), "standard users must not delete repository files");

const beforeLoadDeleteCount = deletionStore.deletedKeys.length;
const deletionCreatorImports = await postStateRaw(deletionStore, {
  action: "loadImports",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.equal(deletionStore.deletedKeys.length, beforeLoadDeleteCount, "loading imports must not delete repository records");
assert.equal(deletionCreatorImports.response.status, 410);

const deletionProfileImports = await postStateRaw(deletionStore, {
  action: "loadDoctorProfileImports",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: "TITUS HACKMAN::mmc",
  doctorKey: "TITUS HACKMAN",
  displayName: "Titus HACKMAN",
  sourceTypes: ["mmc"],
});
assert.equal(deletionProfileImports.response.status, 410);

const deletionUserImports = await postStateRaw(deletionStore, {
  action: "loadImports",
  email: "user@example.com",
  password: "user-password",
});
assert.equal(deletionUserImports.response.status, 410);

console.log("Fixture smoke test passed.");
