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
clinicianRows[6][1] = "Mina NESSIM 49948";
clinicianRows[7][1] = "Clare 0422067042";
clinicianRows[9][1] = "Albert EXAMPLE 49981";
clinicianRows[11][0] = "Silver Dr 8";
clinicianRows[11][3] = "49985 (OFF FLOOR)";
clinicianRows[13][1] = "Sehrish EXAMPLE 49771";
clinicianRows[15][0] = "FT Clinician 3";
clinicianRows[15][3] = "49937";
clinicianRows[16][1] = "Alex LIN";
clinicianRows[16][2] = "49981";
clinicianRows[18][1] = "Morgan TEST";
clinicianRows[18][2] = "49888";
clinicianRows[21][5] = "Silver Dr IC";
clinicianRows[21][6] = "Pat";
clinicianRows[21][7] = "PF";
clinicianRows[21][8] = "49903";
clinicianRows[23][6] = "PM Silver EXTRA 49887";
clinicianRows[26][5] = "Orange Dr 8";
clinicianRows[26][8] = "49905";
clinicianRows[29][6] = "PM Orange EXTRA 49889";
clinicianRows[32][5] = "FT Clinician 3";
clinicianRows[32][8] = "49937";
clinicianRows[36][6] = "PM Fast EXTRA";
clinicianRows[36][7] = "49891";
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
  ["AM", "Orange Dr IC", "Mina NESSIM", "49948", true],
  ["AM", "Orange Dr IC", "Albert EXAMPLE", "49981", true],
  ["AM", "Silver Dr 8", "", "49985", false],
  ["AM", "Silver Dr 8", "Sehrish EXAMPLE", "49771", true],
  ["AM", "FT Clinician 3", "", "49937", false],
  ["AM", "FT Clinician 3", "Alex LIN", "49981", true],
  ["AM", "FT Clinician 3", "Morgan TEST", "49888", true],
  ["PM", "Silver Dr IC", "Pat", "49903", true],
  ["PM", "Silver Dr IC", "PM Silver EXTRA", "49887", true],
  ["PM", "Orange Dr 8", "", "49905", false],
  ["PM", "Orange Dr 8", "PM Orange EXTRA", "49889", true],
  ["PM", "FT Clinician 3", "", "49937", false],
  ["PM", "FT Clinician 3", "PM Fast EXTRA", "49891", true],
  ["Night", "FT Clinician ND IC", "Alex", "49742", true],
]);

console.log("DDH contact workbook extraction fixtures passed.");
