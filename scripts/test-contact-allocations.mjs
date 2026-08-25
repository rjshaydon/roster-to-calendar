import assert from "node:assert/strict";
import {
  attachContactAllocations,
  contactsAfterShiftChange,
  contactExtractHasExpired,
  contactExtractStatus,
  contactOperationalDate,
  normaliseContactListExtract,
  shouldCarryPreviousNightContacts,
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
assert.equal(contactOperationalDate(new Date("2026-08-25T21:29:00Z")), "2026-08-25", "the operational date remains Tuesday before 07:30 Melbourne Wednesday");
assert.equal(contactOperationalDate(new Date("2026-08-25T21:30:00Z")), "2026-08-26", "the operational date rolls over at 07:30 Melbourne");
assert.equal(contactExtractHasExpired("2026-08-24", new Date("2026-08-24T23:59:00Z")), false, "retain Night phone JSON until 10:00 Melbourne the following day during testing");
assert.equal(contactExtractHasExpired("2026-08-24", new Date("2026-08-25T00:00:00Z")), true, "expire Night phone JSON at 10:00 Melbourne the following day during testing");

const shiftChangeContacts = [
  contact("Orange Dr 1", "Alex", "49901"),
  { ...contact("Orange Dr 7", "Pat", "49905"), shift: "PM" },
  { ...contact("Orange DR 1", "Chris", "49726"), shift: "Night" },
];
assert.deepEqual(contactsAfterShiftChange(shiftChangeContacts, { date: "2026-08-25", now: new Date("2026-08-24T21:29:00Z") }).map((entry) => entry.shift), [],
  "the morning contacts must remain hidden before the 07:30 handover");
assert.deepEqual(contactsAfterShiftChange(shiftChangeContacts, { date: "2026-08-25", now: new Date("2026-08-24T21:30:00Z") }).map((entry) => entry.shift), ["AM", "Night"],
  "the previous Night contacts should remain at the 07:30 handover");
assert.deepEqual(contactsAfterShiftChange(shiftChangeContacts, { date: "2026-08-25", now: new Date("2026-08-24T23:59:00Z") }).map((entry) => entry.shift), ["AM", "Night"],
  "the previous Night contacts should remain until 10:00 during testing");
assert.deepEqual(contactsAfterShiftChange(shiftChangeContacts, { date: "2026-08-25", now: new Date("2026-08-25T00:00:00Z") }).map((entry) => entry.shift), ["AM"],
  "Night contacts should clear at 10:00 during testing");
assert.equal(shouldCarryPreviousNightContacts("2026-08-24", "2026-08-25", new Date("2026-08-24T23:59:00Z")), true,
  "the previous extract should remain usable for Night contacts until 10:00 during testing");
assert.equal(shouldCarryPreviousNightContacts("2026-08-24", "2026-08-25", new Date("2026-08-25T00:00:00Z")), false,
  "the previous extract should stop being usable at 10:00 during testing");
assert.deepEqual(contactsAfterShiftChange(shiftChangeContacts, { date: "2026-08-25", now: new Date("2026-08-25T00:30:00Z") }).map((entry) => entry.shift), ["AM"],
  "at 10:30 Melbourne, future PM and Night contact entries must remain hidden");
assert.deepEqual(contactsAfterShiftChange(shiftChangeContacts, { date: "2026-08-25", now: new Date("2026-08-25T05:00:00Z") }).map((entry) => entry.shift), ["AM", "PM"],
  "the PM contacts should become available after 15:00 Melbourne");
assert.deepEqual(contactsAfterShiftChange(shiftChangeContacts, { date: "2026-08-25", now: new Date("2026-08-25T13:00:00Z") }).map((entry) => entry.shift), ["AM", "PM", "Night"],
  "the Night contacts should become available after 23:00 Melbourne");

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

const dailyAllocation = attachContactAllocations([
  assignment("MMC", "AM", "Amber", "Tara KAMATH", "SMS"),
  assignment("MMC", "AM", "Float", "Tara JOHANSSON", "SMS"),
  assignment("MMC", "AM", "SSU", "Qingyang CHEN", "SMS"),
  assignment("MMC", "AM", "Junior Registrar", "Sophie HE", "Junior Registrar"),
  assignment("MMC", "AM", "Junior Registrar", "Yee Ann SOO", "Junior Registrar"),
  assignment("MMC", "AM", "Clinic", "Stephen GILDFIND", "SMS"),
  assignment("MMC", "AM", "HMO", "Arnav MEHTA", "HMO"),
], [
  contact("AMBER (SMS/SR) 25168", "Tara K", "25168"),
  contact("CLINIC (SMS/SR) 25138", "Tara", "25138"),
  contact("SSU (SMS/SR) 25143", "Qing", "25143"),
  contact("SEPSIS DR MUST CARRY SEPSIS #25192", "Sophie", "25192"),
  contact("Dr", "Ann", "25179"),
  contact("RESUS (SMS/SR) 25140", "Steve G", "25140"),
  contact("Dr", "Ama", "25721"),
]);
assert.equal(dailyAllocation.matchedCount, 6, "safe live names should match despite a changed roster stream");
assert.equal(allocationFor(dailyAllocation, "Tara KAMATH").phone, "25168");
assert.equal(allocationFor(dailyAllocation, "Tara KAMATH").streamLabel, "Amber");
assert.equal(allocationFor(dailyAllocation, "Tara JOHANSSON").phone, "25138");
assert.equal(allocationFor(dailyAllocation, "Tara JOHANSSON").streamLabel, "Clinic");
assert.equal(dailyAllocation.assignments.find((entry) => entry.person.displayName === "Tara JOHANSSON").team, "Float", "a confirmed contact allocation must not override the roster stream");
assert.equal(allocationFor(dailyAllocation, "Qingyang CHEN").matchMethod, "first-name-prefix");
assert.equal(allocationFor(dailyAllocation, "Sophie HE").streamLabel, "Sepsis");
assert.equal(allocationFor(dailyAllocation, "Yee Ann SOO").matchMethod, "internal-given-name");
assert.equal(allocationFor(dailyAllocation, "Stephen GILDFIND").matchMethod, "alias-surname-initial");
assert.equal(dailyAllocation.assignments.find((entry) => entry.person.displayName === "Arnav MEHTA").contactAllocation, undefined, "Ama must never be guessed as Arnav");
assert.equal(dailyAllocation.unmatched.length, 1);
assert.equal(dailyAllocation.unmatched[0].name, "Ama");
assert.equal(dailyAllocation.unmatched[0].reviewReason, "No safe name match");

const rosterAuthoritative = attachContactAllocations([
  assignment("MMC", "AM", "Float", "Tara JOHANSSON", "SMS"),
  assignment("MMC", "AM", "Clinic", "Stephen GILDFIND", "SMS"),
], [contact("CLINIC (SMS/SR) 25138", "Tara", "25138")]);
assert.equal(allocationFor(rosterAuthoritative, "Tara JOHANSSON").phone, "25138");
assert.equal(rosterAuthoritative.assignments.find((entry) => entry.person.displayName === "Tara JOHANSSON").team, "Float",
  "a contact role only supplies a phone number; the roster remains the source of the stream");
assert.equal(rosterAuthoritative.assignments.find((entry) => entry.person.displayName === "Stephen GILDFIND").contactDisplacedBy, undefined,
  "a contact match must not displace another rostered clinician");

const duplicateQing = attachContactAllocations([
  assignment("MMC", "AM", "SSU", "Qingyang CHEN", "SMS"),
  assignment("MMC", "AM", "SSU", "Qing LI", "SMS"),
], [contact("SSU (SMS/SR) 25143", "Qing", "25143")]);
assert.equal(duplicateQing.matchedCount, 0, "a short name remains unresolved where two roster candidates are equally safe");

const crossPeriod = attachContactAllocations([
  assignment("MMC", "PM", "SSU", "Qingyang CHEN", "SMS"),
], [contact("SSU (SMS/SR) 25143", "Qing", "25143")]);
assert.equal(crossPeriod.matchedCount, 0, "a contact allocation cannot cross periods");

const mmcNightAliases = attachContactAllocations([
  assignment("MMC", "Night", "SSU", "Melanie ENGEL", "Transitional/Intermediate Registrar"),
  assignment("MMC", "Night", "SSU", "Michael FERGUSON", "HMO"),
  assignment("MMC", "Night", "Main team", "Yiran WANG", "Intern"),
], [
  { ...contact("SSU (SMS/SR) 25143", "Mel", "25143"), shift: "Night" },
  { ...contact("SSU Dr 25144", "Mickey", "25144"), shift: "Night" },
  { ...contact("Dr", "Ian", "25721"), shift: "Night" },
]);
assert.equal(mmcNightAliases.matchedCount, 3, "approved MMC first-name aliases should attach the Night phones");
assert.deepEqual(mmcNightAliases.assignments.map((entry) => entry.contactAllocation?.phone), ["25143", "25144", "25721"]);

const temporaryContact = contact("Dr", "Ama", "25721");
const temporaryAssignments = [
  assignment("MMC", "AM", "Fast Track", "Arnav MEHTA", "HMO"),
  assignment("MMC", "AM", "Fast Track", "Other DOCTOR", "HMO"),
];
const unresolvedTemporary = attachContactAllocations(temporaryAssignments, [temporaryContact]);
const temporaryKey = unresolvedTemporary.unmatched[0].contactKey;
const resolvedTemporary = attachContactAllocations(temporaryAssignments, [temporaryContact], [{ id: "resolution-1", contactKey: temporaryKey, doctorKey: temporaryAssignments[0].person.doctorKey, revision: 1, active: true }]);
assert.equal(allocationFor(resolvedTemporary, "Arnav MEHTA").phone, "25721", "a temporary correction should attach only the daily phone allocation");
assert.equal(allocationFor(resolvedTemporary, "Arnav MEHTA").matchMethod, "manual");
assert.equal(resolvedTemporary.unmatched.length, 0, "a resolved allocation should leave the review list");

const ddhExcluded = normaliseContactListExtract({
  sourceId: "ddh-daily-contact-sheet", sourceDate: "2026-08-25", contacts: [
    { area: "Dandenong Emergency", shift: "AM", role: "Geriatrician in ED consultant", name: "A", phone: "49901", isPopulated: true },
    { area: "Dandenong Emergency", shift: "AM", role: "CART NP/NPC", name: "B", phone: "49902", isPopulated: true },
    { area: "Dandenong Emergency", shift: "AM", role: "Miprep HMO", name: "C", phone: "49903", isPopulated: true },
    { area: "Dandenong Emergency", shift: "PM", role: "Miprep HMO", name: "D", phone: "49904", isPopulated: true },
  ],
});
assert.deepEqual(ddhExcluded.contacts.map((entry) => entry.name), ["D"], "only the three requested DDH AM contact-review rows should be excluded");

const ddhExtract = normaliseContactListExtract({
  sourceId: "ddh-daily-contact-sheet",
  sourceDate: "2026-08-25",
  providerModifiedAt: "2026-08-25T03:00:00Z",
  contacts: [
    { area: "Dandenong Emergency", shift: "AM", role: "Orange Dr IC", name: "Alex", phone: "49900", isPopulated: true },
    { area: "Dandenong Emergency", shift: "PM", role: "Silver Dr 1", name: "Pat", phone: "49903", isPopulated: true },
    { area: "Dandenong Emergency", shift: "Night", role: "FT Clinician ND IC", name: "Chris", phone: "49742", isPopulated: true },
  ],
});
assert.ok(ddhExtract, "a DDH clinicians extract should normalise independently from MMC");
const ddhMatches = attachContactAllocations([
  assignment("DDH", "AM", "Orange", "Alex EXAMPLE", "SMS"),
  assignment("DDH", "PM", "Silver", "Patrick EXAMPLE", "SMS"),
  assignment("DDH", "Night", "Fast Track", "Chris EXAMPLE", "HMO"),
], ddhExtract.contacts);
assert.equal(ddhMatches.matchedCount, 3, "DDH Orange, Silver and Fast Track contacts should match their roster streams");
assert.deepEqual(ddhMatches.assignments.map((entry) => entry.contactAllocation?.phone), ["49900", "49903", "49742"]);

const ddhAlternatePhone = attachContactAllocations([
  assignment("DDH", "AM", "Fast Track", "Alex LIN", "HMO"),
], [{ area: "Dandenong Emergency", shift: "AM", role: "FT Clinician 3", name: "Alex LIN", phone: "49981", isPopulated: true }]);
assert.equal(allocationFor(ddhAlternatePhone, "Alex LIN").phone, "49981",
  "a five-digit extension from a DDH continuation row should attach to the rostered clinician");

const fixedServiceAllocations = attachContactAllocations([
  assignment("MMC", "AM", "Sepsis", "Sophie SEPSIS", "HMO"),
  assignment("MMC", "AM", "Geriatrician", "Georgia GERIATRICIAN", "SMS"),
  assignment("MMC", "AM", "CART clinician", "Casey CART", "NP"),
  assignment("DDH", "AM", "ED Care-Co", "Eddie CARECO", "HMO"),
  assignment("DDH", "AM", "GAP", "Grace GAP", "SMS"),
  assignment("DDH", "AM", "CS onsite", "Sally SUPPORT", "SMS"),
], [
  { area: "Adult Emergency", shift: "AM", role: "SEPSIS DR MUST CARRY SEPSIS #", name: "", phone: "25192", isPopulated: false },
  { area: "Adult Emergency", shift: "AM", role: "GERIATRICIAN", name: "", phone: "25768", isPopulated: false },
  { area: "Adult Emergency", shift: "AM", role: "CART Clinician", name: "", phone: "0417 489 358", isPopulated: false },
  { area: "Dandenong Emergency", shift: "AM", role: "ED Care-Co", name: "", phone: "49754", isPopulated: false },
  { area: "Dandenong Emergency", shift: "AM", role: "GAP (Geriatric AH)", name: "", phone: "0488 762 079", isPopulated: false },
  { area: "Dandenong Emergency", shift: "AM", role: "Clinical Support on site", name: "", phone: "03 95549098", isPopulated: false },
]);
assert.equal(fixedServiceAllocations.matchedCount, 6, "approved empty-name service rows should attach by unique roster service");
assert.equal(allocationFor(fixedServiceAllocations, "Sophie SEPSIS").phone, "25192");
assert.equal(allocationFor(fixedServiceAllocations, "Georgia GERIATRICIAN").phone, "25768");
assert.equal(allocationFor(fixedServiceAllocations, "Casey CART").phone, "0417 489 358");
assert.equal(allocationFor(fixedServiceAllocations, "Eddie CARECO").phone, "49754");
assert.equal(allocationFor(fixedServiceAllocations, "Grace GAP").phone, "0488 762 079");
assert.equal(allocationFor(fixedServiceAllocations, "Sally SUPPORT").phone, "03 95549098");

const fixedServicePriority = attachContactAllocations([
  assignment("MMC", "AM", "Sepsis", "Sophie SEPSIS", "HMO"),
], [
  { area: "Adult Emergency", shift: "AM", role: "Dr", name: "Sophie", phone: "25179", isPopulated: true },
  { area: "Adult Emergency", shift: "AM", role: "SEPSIS DR MUST CARRY SEPSIS #", name: "", phone: "25192", isPopulated: false },
]);
assert.equal(allocationFor(fixedServicePriority, "Sophie SEPSIS").phone, "25192", "the dedicated service phone should take priority over a generic named row");

const ordinaryEmptyPhone = attachContactAllocations([
  assignment("MMC", "AM", "Green", "Greta GREEN", "SMS"),
], [{ area: "Adult Emergency", shift: "AM", role: "GREEN (CIC/AO)", name: "", phone: "25134", isPopulated: false }]);
assert.equal(ordinaryEmptyPhone.matchedCount, 0, "ordinary fixed phones must remain ignored when their rows have no name");

const unrosteredFixedService = attachContactAllocations([], [
  { area: "Adult Emergency", shift: "AM", role: "GERIATRICIAN", name: "", phone: "25768", isPopulated: false },
  { area: "Adult Emergency", shift: "AM", role: "CART Clinician", name: "Ben G", phone: "0417 489 358", isPopulated: true },
  { area: "Dandenong Emergency", shift: "AM", role: "ED Care-Co", name: "", phone: "49754", isPopulated: false },
]);
assert.equal(unrosteredFixedService.unmatched.length, 0, "an unrostered fixed service phone should not create a review item");
assert.deepEqual(unrosteredFixedService.serviceContacts.map((entry) => entry.role), ["GERIATRICIAN", "CART Clinician", "ED Care-Co"],
  "explicit services should remain available for dedicated cards when they are not doctor-roster allocations");

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

function contact(role, name, phone) {
  return { area: "Adult Emergency", shift: "AM", role, name, phone, isPopulated: true };
}

function allocationFor(matches, displayName) {
  const allocation = matches.assignments.find((entry) => entry.person.displayName === displayName)?.contactAllocation;
  assert.ok(allocation, `${displayName} should have a contact allocation`);
  return allocation;
}
