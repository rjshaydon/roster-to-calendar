import assert from "node:assert/strict";
import {
  attachContactAllocations,
  contactExtractHasExpired,
  contactExtractStatus,
  normaliseContactListExtract,
} from "../public/static/contact-allocations.js";

const extract = normaliseContactListExtract({
  sourceId: "mmc-shift-allocations",
  sourceDate: "2026-08-24",
  contacts: [
    { area: "Adult Emergency", shift: "AM", role: "GREEN (CIC/AO) 25134", name: "", phone: "25134", isPopulated: true },
    { area: "Adult Emergency", shift: "PM", role: "GREEN (CIC/AO) 25134", name: "Pat", phone: "25134", isPopulated: true },
    { area: "Adult Emergency", shift: "PM", role: "CLINIC (SMS/SR) 25138", name: "Ananth", phone: "25138", isPopulated: true },
    { area: "Adult Emergency", shift: "PM", role: "AMBER (SMS/SR) 25168", name: "Jacqui", phone: "25168", isPopulated: true },
    { area: "Paediatric Emergency", shift: "PM", role: "Paeds Dr - CIC/AO", name: "Simon", phone: "25145", isPopulated: true },
  ],
});

assert.ok(extract, "valid contact list should normalise");
assert.equal(extract.contacts[0].isPopulated, false, "a phone without a named doctor must not be live");
assert.equal(contactExtractStatus(extract, { date: "2026-08-24", now: new Date("2026-08-24T01:00:00Z") }), "available");
assert.equal(contactExtractStatus(extract, { date: "2026-08-25", now: new Date("2026-08-24T01:00:00Z") }), "not-current");
assert.equal(contactExtractHasExpired("2026-08-24", new Date("2026-08-24T23:59:00Z")), false, "retain through the Night shift");
assert.equal(contactExtractHasExpired("2026-08-24", new Date("2026-08-25T00:00:00Z")), true, "expire at 10:00 Melbourne the following day");

const assignments = [
  assignment("MMC", "PM", "Green", "Pat FINN", "SMS"),
  assignment("MMC", "PM", "Clinic", "Ananth SUNDARALINGAM", "SMS"),
  assignment("MMC", "PM", "Amber", "Jacqueline MOREL", "SMS"),
  assignment("MCH", "PM", "Paeds", "Simon EXAMPLE", "Senior Registrar"),
];
const matches = attachContactAllocations(assignments, extract.contacts);
assert.equal(matches.matchedCount, 4, "stream, period and name should produce four safe matches");
assert.equal(matches.assignments[0].contactAllocation.phone, "25134");
assert.equal(matches.assignments[1].contactAllocation.phone, "25138");
assert.equal(matches.assignments[2].contactAllocation.matchMethod, "alias", "Jacqui should match Jacqueline only through a controlled alias");
assert.equal(matches.assignments[3].contactAllocation.phone, "25145", "Paediatric allocation should stay within MCH");

const ambiguous = attachContactAllocations([
  assignment("MMC", "PM", "Green", "Pat FINN", "SMS"),
  assignment("MMC", "PM", "Green", "Patrick EXAMPLE", "SMS"),
], [extract.contacts[1]]);
assert.equal(ambiguous.matchedCount, 0, "an ambiguous short name must not receive a phone allocation");

console.log("Contact allocation matching fixtures passed.");

function assignment(source, period, team, displayName, seniority) {
  return {
    source,
    period,
    team,
    suggestedTitle: `${team} ${period}`,
    person: { doctorKey: displayName.toLowerCase().replaceAll(" ", "-"), displayName, sourceType: source.toLowerCase(), seniority },
    event: { title: `${source}: ${team} ${period}` },
  };
}
