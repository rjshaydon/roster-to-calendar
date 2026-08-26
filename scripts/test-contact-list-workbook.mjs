import assert from "node:assert/strict";
import XLSX from "xlsx";

import { extractDdhClinicianContactsFromWorkbook, extractMmcDoctorContactsFromWorkbook } from "../functions/_lib/contact-list-workbook.js";

const mmcWorkbook = XLSX.utils.book_new();
const mmcRows = Array.from({ length: 42 }, () => Array(9).fill(""));
mmcRows[1][3] = "25th August 2026";
mmcRows[5][0] = "CART Clinician";
mmcRows[5][1] = "Casey";
mmcRows[5][2] = "0417 489 358";
mmcRows[6][3] = "SEPSIS DR";
mmcRows[6][4] = "Sophie";
mmcRows[6][5] = "25192";
mmcRows[7][6] = "ADULT SMS ON CALL";
mmcRows[7][7] = "Switch";
mmcRows[7][8] = "Call Switch - 92";
XLSX.utils.book_append_sheet(mmcWorkbook, XLSX.utils.aoa_to_sheet(mmcRows), "SHIFT ALLOCATIONS");
const mmcExtract = await extractMmcDoctorContactsFromWorkbook(XLSX.write(mmcWorkbook, { type: "array", bookType: "xlsx" }));
assert.deepEqual(mmcExtract.contacts.map(({ shift, role, name, phone }) => [shift, role, name, phone]), [
  ["AM", "CART Clinician", "Casey", "0417 489 358"],
  ["PM", "SEPSIS DR", "Sophie", "25192"],
  ["Night", "ADULT SMS ON CALL", "Switch", ""],
], "MMC should retain 5-10 digit telephone numbers and reject shorter instructions");

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
clinicianRows[19][0] = "Clinical Support on site";
clinicianRows[19][1] = "Shawn SUPPORT";
clinicianRows[19][3] = "03 95549098";
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
clinicianRows[44][10] = "Dr IC";
clinicianRows[44][11] = "Tinnie";
clinicianRows[44][12] = "TC";
clinicianRows[44][13] = "49900";
clinicianRows[45][10] = "Orange Dr 1";
clinicianRows[45][11] = "Jonathan (49931)";
clinicianRows[45][12] = "JW";
clinicianRows[45][13] = "49726";
clinicianRows[46][10] = "Orange Dr 2";
clinicianRows[46][11] = "Jaz";
clinicianRows[46][12] = "49971";
clinicianRows[46][13] = "49741";
clinicianRows[47][10] = "Orange Dr 3";
clinicianRows[47][11] = "Titus";
clinicianRows[47][12] = "TH";
clinicianRows[47][13] = "12345";
clinicianRows[48][10] = "Orange Dr 4";
clinicianRows[48][11] = "Titus";
clinicianRows[48][12] = "54321";
clinicianRows[48][13] = "12345";
clinicianRows[49][10] = "Orange Dr 5";
clinicianRows[49][11] = "Titus (ph24680)";
clinicianRows[49][12] = "12345";
clinicianRows[49][13] = "12345";
clinicianRows[50][10] = "Orange Dr 6";
clinicianRows[50][11] = "Titus 1";
clinicianRows[50][12] = "TH";
clinicianRows[50][13] = "12345";
XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(clinicianRows), "ED Clinicians");

const bytes = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
const extract = await extractDdhClinicianContactsFromWorkbook(bytes, {
  providerModifiedAt: "2026-08-24T23:30:00Z",
});

assert.equal(extract.sourceId, "ddh-daily-contact-sheet");
assert.equal(extract.sourceDate, "2026-08-25", "the source date should use Melbourne time");
const earlyMorningExtract = await extractDdhClinicianContactsFromWorkbook(bytes, {
  providerModifiedAt: "2026-08-25T20:00:00Z",
});
assert.equal(earlyMorningExtract.sourceDate, "2026-08-25",
  "a DDH update before 07:30 Wednesday should remain attached to Tuesday's Night shift");
assert.deepEqual(extract.contacts.filter(({ shift, role }) => shift === "Night" && /^(?:Dr IC|Orange Dr [12])$/.test(role))
  .map(({ role, name, phone }) => [role, name, phone]), [
  ["Dr IC", "Tinnie", "49900"],
  ["Orange Dr 1", "Jonathan", "49931"],
  ["Orange Dr 2", "Jaz", "49971"],
], "DDH Night contacts should prefer a phone with the name, then EMR, then the normal Number column");
assert.deepEqual(extract.contacts.filter(({ shift, role }) => shift === "Night" && /^Orange Dr [3-6]$/.test(role))
  .map(({ role, phone }) => [role, phone]), [
  ["Orange Dr 3", "12345"],
  ["Orange Dr 4", "54321"],
  ["Orange Dr 5", "24680"],
  ["Orange Dr 6", "12345"],
], "DDH Night phone precedence should cover plain names, EMR phones, embedded phones, and non-phone name suffixes");
assert.deepEqual(extract.contacts.map(({ shift, role, name, phone, isPopulated }) => [shift, role, name, phone, isPopulated]), [
  ["AM", "Doctor In Charge / AVAO", "Di", "49900", true],
  ["AM", "Orange Dr IC", "Shilpa", "49970", true],
  ["AM", "Orange Dr IC", "Mina NESSIM", "49948", true],
  ["AM", "Orange Dr IC", "Clare", "0422067042", true],
  ["AM", "Orange Dr IC", "Albert EXAMPLE", "49981", true],
  ["AM", "Silver Dr 8", "", "49985", false],
  ["AM", "Silver Dr 8", "Sehrish EXAMPLE", "49771", true],
  ["AM", "FT Clinician 3", "", "49937", false],
  ["AM", "FT Clinician 3", "Alex LIN", "49981", true],
  ["AM", "FT Clinician 3", "Morgan TEST", "49888", true],
  ["AM", "Clinical Support on site", "Shawn SUPPORT", "03 95549098", true],
  ["PM", "Silver Dr IC", "Pat", "49903", true],
  ["PM", "Silver Dr IC", "PM Silver EXTRA", "49887", true],
  ["PM", "Orange Dr 8", "", "49905", false],
  ["PM", "Orange Dr 8", "PM Orange EXTRA", "49889", true],
  ["PM", "FT Clinician 3", "", "49937", false],
  ["PM", "FT Clinician 3", "PM Fast EXTRA", "49891", true],
  ["Night", "FT Clinician ND IC", "Alex", "49742", true],
  ["Night", "Dr IC", "Tinnie", "49900", true],
  ["Night", "Orange Dr 1", "Jonathan", "49931", true],
  ["Night", "Orange Dr 2", "Jaz", "49971", true],
  ["Night", "Orange Dr 3", "Titus", "12345", true],
  ["Night", "Orange Dr 4", "Titus", "54321", true],
  ["Night", "Orange Dr 5", "Titus", "24680", true],
  ["Night", "Orange Dr 6", "Titus 1", "12345", true],
]);

console.log("DDH contact workbook extraction fixtures passed.");
