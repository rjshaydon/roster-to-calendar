import assert from "node:assert/strict";
import { buildVhhDerivedRosterPayload, normaliseVhhRosterExtract, VHH_ROSTER_SOURCE_ID } from "../functions/_lib/vhh-roster.js";
import { normaliseContactListExtract, VHH_CONTACT_LIST_SOURCE_ID } from "../public/static/contact-allocations.js";
import * as XLSX from "xlsx";
import { extractVhhRosterWorkbook } from "./vhh-roster-workbook.mjs";
import { automationSourceDefinition } from "../functions/_lib/automation-import.js";

assert.equal(automationSourceDefinition(VHH_ROSTER_SOURCE_ID)?.provider, "sharepoint", "VHH must use the raw SharePoint workbook ingress");

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
      { sourceRow: 23, sourceShiftLabel: "JMS Teaching Timetable", shiftLabel: "JMS Teaching Timetable", assignments: [{ date: "2026-08-24", displayedDate: "24-Aug", namesText: "Never, Import", sourceCell: "B23" }] },
    ],
  }],
};

const normalisedRoster = normaliseVhhRosterExtract(rosterExtract);
assert.ok(normalisedRoster, "valid VHH roster JSON should normalise");
assert.equal(normalisedRoster.blocks[0].rows.length, 2, "JMS Teaching Timetable must be excluded at the JSON boundary");
const derived = buildVhhDerivedRosterPayload({ extract: rosterExtract, contentHash: "vhh-test-content-hash", fileId: "vhh-test-file" });
assert.equal(derived.file.sourceType, "vhh");
assert.equal(derived.doctors.find((doctor) => doctor.key === "RICHARD HAYDON")?.displayName, "Richard Haydon", "VHH Last, First names must normalise to First Last");
assert.equal(derived.eventsByDoctor["RICHARD HAYDON"][0].start, "2026-08-24T08:00:00");
assert.equal(derived.eventsByDoctor["RICHARD HAYDON"][0].end, "2026-08-24T15:30:00");
assert.equal(derived.eventsByDoctor["ALEX SMITH"][0].allDay, true, "a VHH cell with no stated hours must not invent timings");
assert.equal(derived.doctors.some((doctor) => doctor.key === "IMPORT NEVER"), false, "timetable names must not become staff records");

const workbook = XLSX.utils.book_new();
const activeSheet = XLSX.utils.aoa_to_sheet([
  ["SHIFT LABEL", new Date("2026-08-24T00:00:00Z"), new Date("2026-08-25T00:00:00Z")],
  ["AM REG", "Haydon, Richard (0800-1530)", ""],
  ["", "", "Smith, Alex"],
  ["JMS Teaching Timetable", "Never, Import", ""],
]);
XLSX.utils.book_append_sheet(workbook, activeSheet, "24.08-20.09");
const hiddenSheet = XLSX.utils.aoa_to_sheet([["SHIFT LABEL", new Date("2026-08-24T00:00:00Z")], ["AM REG", "Hidden, Person"]]);
XLSX.utils.book_append_sheet(workbook, hiddenSheet, "Hidden");
workbook.Workbook = { Sheets: [{ name: "24.08-20.09", Hidden: 0 }, { name: "Hidden", Hidden: 1 }] };
const workbookBytes = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
const workbookFile = new File([workbookBytes], "Active Medical Roster.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", lastModified: Date.parse("2026-08-26T01:00:00Z") });
const workbookExtract = await extractVhhRosterWorkbook(workbookFile, { providerModifiedAt: "2026-08-26T01:00:00Z", providerVersion: "vhh-etag-2" });
assert.equal(workbookExtract.fileName, "Active Medical Roster.xlsx");
assert.equal(workbookExtract.blocks.length, 1, "hidden worksheets must not be extracted");
assert.equal(workbookExtract.blocks[0].rows.length, 2, "the teaching timetable and later rows must not be extracted");
assert.equal(workbookExtract.blocks[0].rows[1].shiftLabel, "AM REG", "continuation rows must inherit their shift label");
assert.equal(workbookExtract.blocks[0].rows[1].assignments[0].sourceCell, "C3");
const workbookDerived = buildVhhDerivedRosterPayload({ extract: workbookExtract, contentHash: "vhh-xlsx-test", fileId: "vhh-xlsx-file", fileSize: workbookFile.size, lastModified: workbookFile.lastModified });
assert.equal(workbookDerived.file.name, "Active Medical Roster.xlsx");
assert.equal(workbookDerived.file.size, workbookFile.size);
assert.equal(workbookDerived.file.lastModified, workbookFile.lastModified);
assert.equal(workbookDerived.doctors.some((doctor) => doctor.key === "HIDDEN PERSON"), false);

const contactDirectory = normaliseContactListExtract({
  sourceId: VHH_CONTACT_LIST_SOURCE_ID,
  sourceDate: "2026-08-26",
  providerModifiedAt: "2026-08-26T01:00:00Z",
  cic: { phone: "90000", name: "CIC Doctor" },
  doctors: [
    { role: "AM SMS", phone: "90001", name: "First Doctor" },
    { role: "PM REG", phone: "90002", name: "" },
  ],
});
assert.ok(contactDirectory, "VHH CIC and doctor directory should normalise");
assert.deepEqual(contactDirectory.contacts, [], "VHH directory must not masquerade as shift-matched contacts before mapping is approved");
assert.equal(contactDirectory.directory.cic.phone, "90000");
assert.equal(contactDirectory.directory.doctors.length, 2);

console.log("VHH automation fixtures passed.");
