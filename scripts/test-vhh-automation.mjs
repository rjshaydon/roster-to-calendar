import assert from "node:assert/strict";
import { buildVhhDerivedRosterPayload, normaliseVhhRosterExtract, VHH_ROSTER_SOURCE_ID } from "../functions/_lib/vhh-roster.js";
import { normaliseContactListExtract, VHH_CONTACT_LIST_SOURCE_ID } from "../public/static/contact-allocations.js";

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
