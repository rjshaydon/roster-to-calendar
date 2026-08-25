import assert from "node:assert/strict";
import XLSX from "xlsx";

import { extractDdhClinicianContactsFromWorkbook } from "../functions/_lib/contact-list-workbook.js";

const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
  ["ED Nursing", "This sheet must not be parsed"],
]), "ED Nursing");

const clinicianRows = Array.from({ length: 70 }, () => Array(14).fill(""));
clinicianRows[0][5] = "ED Clinician Phones";
clinicianRows[2][1] = "AM";
clinicianRows[2][6] = "PM";
clinicianRows[2][11] = "ND";
clinicianRows[4][0] = "Doctor In Charge / AVAO";
clinicianRows[4][1] = "Di";
clinicianRows[4][2] = "DF";
clinicianRows[4][3] = "49900";
clinicianRows[5][0] = "Orange Dr IC";
clinicianRows[5][1] = "Shilpa 0406370706";
clinicianRows[5][2] = "ST";
clinicianRows[5][3] = "49970";
clinicianRows[21][5] = "Silver Dr IC";
clinicianRows[21][6] = "Pat";
clinicianRows[21][7] = "PF";
clinicianRows[21][8] = "49903";
clinicianRows[40][10] = "FT Clinician ND IC";
clinicianRows[40][11] = "Alex";
clinicianRows[40][12] = "AX";
clinicianRows[40][13] = "49742";
XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(clinicianRows), "ED Clinicians");

const bytes = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
const extract = await extractDdhClinicianContactsFromWorkbook(bytes, {
  providerModifiedAt: "2026-08-24T23:30:00Z",
});

assert.equal(extract.sourceId, "ddh-daily-contact-sheet");
assert.equal(extract.sourceDate, "2026-08-25", "the source date should use Melbourne time");
assert.deepEqual(extract.contacts.map(({ shift, role, name, phone, isPopulated }) => [shift, role, name, phone, isPopulated]), [
  ["AM", "Doctor In Charge / AVAO", "Di", "49900", true],
  ["AM", "Orange Dr IC", "Shilpa", "49970", true],
  ["PM", "Silver Dr IC", "Pat", "49903", true],
  ["Night", "FT Clinician ND IC", "Alex", "49742", true],
]);

console.log("DDH contact workbook extraction fixtures passed.");
