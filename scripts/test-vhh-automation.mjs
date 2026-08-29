import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { buildVhhDerivedRosterPayload, normaliseVhhRosterExtract, VHH_ROSTER_SOURCE_ID } from "../functions/_lib/vhh-roster.js";
import * as XLSX from "xlsx";
import { extractVhhRosterWorkbook } from "./vhh-roster-workbook.mjs";
import { automationSourceDefinition } from "../functions/_lib/automation-import.js";

assert.equal(automationSourceDefinition(VHH_ROSTER_SOURCE_ID)?.provider, "sharepoint", "VHH must use the raw SharePoint workbook ingress");
const appSource = await readFile(new URL("../public/static/app.js", import.meta.url), "utf8");
assert.match(appSource, /function normalizedDoctorSourceTypes[\s\S]*?item === "vhh"/, "VHH must remain a valid source when opening an individual roster calendar");

const rosterExtract = {
  sourceId: VHH_ROSTER_SOURCE_ID,
  providerModifiedAt: "2026-08-26T01:00:00Z",
  providerVersion: "vhh-etag-1",
  blocks: [{
    sheetName: "24.08-20.09", blockIndex: 1, headerRow: 4, teachingTimetableRow: 23,
    dates: [{ sourceColumn: "B", displayedDate: "24-Aug", date: "2026-08-24" }],
    rows: [
      { sourceRow: 7, sourceShiftLabel: "AM REG", shiftLabel: "AM REG", assignments: [{ date: "2026-08-24", displayedDate: "24-Aug", namesText: "Haydon, Richard (0800-1530)", sourceCell: "B7" }] },
      { sourceRow: 14, sourceShiftLabel: "PM SMS", shiftLabel: "PM SMS", assignments: [{ date: "2026-08-24", displayedDate: "24-Aug", namesText: "Smith, Alex\nJones, Sam", sourceCell: "B14" }] },
      { sourceRow: 15, sourceShiftLabel: "ON JMS", shiftLabel: "ON JMS", assignments: [{ date: "2026-08-24", displayedDate: "24-Aug", namesText: "Samantha Young", sourceCell: "B15" }] },
      { sourceRow: 23, sourceShiftLabel: "JMS Teaching Timetable", shiftLabel: "JMS Teaching Timetable", assignments: [{ date: "2026-08-24", displayedDate: "24-Aug", namesText: "Never, Import", sourceCell: "B23" }] },
    ],
  }],
};

const normalisedRoster = normaliseVhhRosterExtract(rosterExtract);
assert.ok(normalisedRoster, "valid VHH roster JSON should normalise");
assert.equal(normalisedRoster.blocks[0].rows.length, 3, "JMS Teaching Timetable must be excluded at the JSON boundary");
const derived = buildVhhDerivedRosterPayload({ extract: rosterExtract, contentHash: "vhh-test-content-hash", fileId: "vhh-test-file" });
assert.equal(derived.file.sourceType, "vhh");
assert.equal(derived.doctors.find((doctor) => doctor.key === "RICHARD HAYDON")?.displayName, "Richard Haydon", "VHH Last, First names must normalise to First Last");
assert.equal(derived.eventsByDoctor["RICHARD HAYDON"][0].start, "2026-08-24T08:00:00");
assert.equal(derived.eventsByDoctor["RICHARD HAYDON"][0].end, "2026-08-24T15:30:00");
assert.equal(derived.eventsByDoctor["ALEX SMITH"][0].start, "2026-08-24T14:30:00", "designation timings must be used when a cell has no override");
assert.equal(derived.eventsByDoctor["ALEX SMITH"][0].end, "2026-08-25T00:00:00", "midnight shifts must finish on the following day");
assert.equal(derived.eventsByDoctor["ALEX SMITH"][0].title, "VHH: PM SMS");
assert.equal(derived.eventsByDoctor["SAMANTHA YOUNG"][0].title, "VHH: Night JMS", "a strict First Last value in a recognised row must remain a clinician shift");
assert.equal(derived.doctors.some((doctor) => doctor.key === "IMPORT NEVER"), false, "timetable names must not become staff records");

const designationCases = [
  ["CST", "Clinical Support", "", "", true, "", "SMS"],
  ["T", "Clinical Support", "", "", true, "", "SMS"],
  ["AM SMS", "AM SMS", "08:00", "17:30", false, "VHH", "SMS"],
  ["AM REG", "AM Reg", "08:00", "17:30", false, "VHH", "Junior Registrar"],
  ["SSU HMO (8-4)", "SSU HMO", "08:00", "16:00", false, "VHH", "HMO"],
  ["AM JMS", "AM JMS", "08:00", "17:30", false, "VHH", "Junior Registrar"],
  ["SWING", "Swing SMS 1000", "10:00", "19:30", false, "VHH", "SMS"],
  ["SWING 1230PM", "Swing SMS 1230", "12:30", "22:00", false, "VHH", "SMS"],
  ["PM SMS", "PM SMS", "14:30", "00:00", false, "VHH", "SMS"],
  ["PM REG", "PM Reg", "14:30", "00:00", false, "VHH", "Junior Registrar"],
  ["PM JMS", "PM JMS", "14:30", "00:00", false, "VHH", "Junior Registrar"],
  ["ON REG", "Night Reg", "23:00", "08:30", false, "VHH", "Junior Registrar"],
  ["ON JMS", "Night JMS", "23:00", "08:30", false, "VHH", "Junior Registrar"],
];
const designationExtract = {
  sourceId: VHH_ROSTER_SOURCE_ID,
  blocks: [{
    sheetName: "Mappings", blockIndex: 1, visible: true,
    dates: [{ sourceColumn: "B", displayedDate: "24-Aug", date: "2026-08-24" }],
    rows: designationCases.map(([label], index) => ({
      sourceRow: index + 2,
      sourceShiftLabel: label,
      shiftLabel: label,
      assignments: [{ date: "2026-08-24", displayedDate: "24-Aug", namesText: `Person${String.fromCharCode(65 + index)}, Test`, sourceCell: `B${index + 2}` }],
    })),
  }],
};
const designationDerived = buildVhhDerivedRosterPayload({ extract: designationExtract, contentHash: "vhh-designations" });
designationCases.forEach(([label, title, startTime, endTime, allDay, location, seniority], index) => {
  const event = designationDerived.eventsByDoctor[`TEST PERSON${String.fromCharCode(65 + index)}`][0];
  assert.equal(event.title, `VHH: ${title}`, `${label} title`);
  assert.equal(event.allDay, allDay, `${label} all-day state`);
  assert.equal(event.location ? "VHH" : "", location, `${label} location`);
  assert.equal(event.seniority, seniority, `${label} seniority`);
  if (!allDay) {
    assert.equal(event.start.slice(11, 16), startTime, `${label} start`);
    assert.equal(event.end.slice(11, 16), endTime, `${label} end`);
  }
});
assert.throws(() => buildVhhDerivedRosterPayload({
  extract: {
    sourceId: VHH_ROSTER_SOURCE_ID,
    blocks: [{
      sheetName: "Unknown", blockIndex: 1, visible: true,
      dates: [{ sourceColumn: "B", displayedDate: "24-Aug", date: "2026-08-24" }],
      rows: [{ sourceRow: 2, sourceShiftLabel: "NEW SHIFT", shiftLabel: "NEW SHIFT", assignments: [{ date: "2026-08-24", namesText: "Doctor, Unknown", sourceCell: "B2" }] }],
    }],
  },
  contentHash: "vhh-unknown",
}), /Unsupported VHH shift designation: NEW SHIFT/);

const workbook = XLSX.utils.book_new();
const activeSheet = XLSX.utils.aoa_to_sheet([
  ["Shift Label", "Mon 24/08/2026", "Tue 25/08/2026"],
  ["AM REG", "Haydon, Richard (0800-1530)", ""],
  ["", "", "Smith, Alex"],
  ["AM JMS", "Hidden, Person\nJones, Sam (0830-1800)", ""],
  ["CST", "Public Holiday", "Support, Clinical (0800-1330)"],
  ["T", "Typo, Taylor", ""],
  ["SSU HMO (8-4)", "Officer, House", ""],
  ["MED STUDENT", "Student, Medical", ""],
  ["", "JMS Teaching Timetable", ""],
  ["", "Never, Import", ""],
  ["", "Ultrasound Teaching Sessions are available for booking via https://pocusprogram.com/ for registrars", ""],
]);
XLSX.utils.book_append_sheet(workbook, activeSheet, "24.08-20.09");
const hiddenSheet = XLSX.utils.aoa_to_sheet([["SHIFT LABEL", new Date("2026-08-24T00:00:00Z")], ["SSU HMO (8-4)", "Hidden, Person"]]);
XLSX.utils.book_append_sheet(workbook, hiddenSheet, "Hidden");
workbook.Workbook = { Sheets: [{ name: "24.08-20.09", Hidden: 0 }, { name: "Hidden", Hidden: 1 }] };
const workbookBytes = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
const workbookFile = new File([workbookBytes], "Active Medical Roster.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", lastModified: Date.parse("2026-08-26T01:00:00Z") });
const workbookExtract = await extractVhhRosterWorkbook(workbookFile, { providerModifiedAt: "2026-08-26T01:00:00Z", providerVersion: "vhh-etag-2" });
assert.equal(workbookExtract.fileName, "Active Medical Roster.xlsx");
assert.equal(workbookExtract.blocks.length, 2, "hidden worksheets must be retained only for term-wide seniority evidence");
assert.equal(workbookExtract.blocks[0].visible, true);
assert.equal(workbookExtract.blocks[1].visible, false);
assert.equal(workbookExtract.blocks[0].rows.length, 7, "the teaching timetable region must not be extracted");
assert.equal(workbookExtract.blocks[0].rows[1].shiftLabel, "AM REG", "continuation rows must inherit their shift label");
assert.equal(workbookExtract.blocks[0].rows[1].assignments[0].sourceCell, "C3");
const workbookDerived = buildVhhDerivedRosterPayload({ extract: workbookExtract, contentHash: "vhh-xlsx-test", fileId: "vhh-xlsx-file", fileSize: workbookFile.size, lastModified: workbookFile.lastModified });
assert.equal(workbookDerived.file.name, "Active Medical Roster.xlsx");
assert.equal(workbookDerived.file.size, workbookFile.size);
assert.equal(workbookDerived.file.lastModified, workbookFile.lastModified);
assert.equal(workbookDerived.doctors.find((doctor) => doctor.key === "PERSON HIDDEN")?.seniority, "HMO", "hidden SSU HMO evidence must classify a visible JMS clinician as HMO for the term");
assert.equal(workbookDerived.eventsByDoctor["PERSON HIDDEN"][0].seniority, "HMO");
assert.equal(workbookDerived.eventsByDoctor["SAM JONES"][0].start, "2026-08-24T08:30:00", "an individual timing must override the JMS default for that clinician only");
assert.equal(workbookDerived.eventsByDoctor["SAM JONES"][0].end, "2026-08-24T18:00:00");
assert.equal(workbookDerived.eventsByDoctor["PERSON HIDDEN"][0].start, "2026-08-24T08:00:00", "another clinician in the same cell must retain the designation default");
assert.equal(workbookDerived.eventsByDoctor["CLINICAL SUPPORT"][0].title, "VHH: Clinical Support");
assert.equal(workbookDerived.eventsByDoctor["CLINICAL SUPPORT"][0].location, "");
assert.equal(workbookDerived.eventsByDoctor["CLINICAL SUPPORT"][0].start, "2026-08-25T08:00:00");
assert.equal(workbookDerived.eventsByDoctor["CLINICAL SUPPORT"][0].end, "2026-08-25T13:30:00");
assert.equal(workbookDerived.eventsByDoctor["TAYLOR TYPO"][0].title, "VHH: Clinical Support", "T must use the CST mapping");
assert.equal(workbookDerived.eventsByDoctor["TAYLOR TYPO"][0].allDay, true);
assert.equal(workbookDerived.eventsByDoctor["TAYLOR TYPO"][0].location, "");
assert.equal(workbookDerived.eventsByDoctor["HOUSE OFFICER"][0].title, "VHH: SSU HMO");
assert.equal(workbookDerived.eventsByDoctor["HOUSE OFFICER"][0].start, "2026-08-24T08:00:00");
assert.equal(workbookDerived.eventsByDoctor["HOUSE OFFICER"][0].end, "2026-08-24T16:00:00");
assert.equal(workbookDerived.doctors.find((doctor) => doctor.key === "HOUSE OFFICER")?.seniority, "HMO");
assert.equal(workbookDerived.doctors.some((doctor) => doctor.key === "PUBLIC HOLIDAY"), false);
assert.equal(workbookDerived.doctors.some((doctor) => doctor.key === "MEDICAL STUDENT"), false);
assert.equal(workbookDerived.doctors.some((doctor) => doctor.key === "IMPORT NEVER"), false);

const incompleteTeachingWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(incompleteTeachingWorkbook, XLSX.utils.aoa_to_sheet([
  ["Shift Label", "Mon 24/08/2026"],
  ["AM JMS", "Doctor, Test"],
  ["", "JMS Teaching Timetable"],
]), "Incomplete");
const incompleteTeachingBytes = XLSX.write(incompleteTeachingWorkbook, { type: "array", bookType: "xlsx" });
await assert.rejects(
  extractVhhRosterWorkbook(new File([incompleteTeachingBytes], "Incomplete.xlsx")),
  /teaching timetable boundaries are incomplete/,
);

console.log("VHH automation fixtures passed.");
