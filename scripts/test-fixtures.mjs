import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import XLSX from "xlsx";

import { buildRosterView, doctorOptions, parseUploadForm, previewSummary } from "../public/static/roster.js";

const mmcWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/AdultTerm1.2026.xlsx", import.meta.url)), {
  cellDates: true,
});
const ddhWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/Dandenong_Emergency_Doctors_Roster_02-02-2026_to_03-05-2026.xlsx", import.meta.url)), {
  cellDates: true,
});

const doctors = doctorOptions(mmcWorkbook, ddhWorkbook);
assert.ok(doctors.length > 100);
const richard = doctors.find((doctor) => doctor.displayName === "Richard HAYDON");
assert.ok(richard);
assert.deepEqual(richard.sourceTypes, ["mmc", "ddh"]);
assert.ok(doctors.find((doctor) => doctor.displayName === "Brianna Dawn Murphy"));
assert.ok(doctors.find((doctor) => doctor.displayName === "Patrick Tan"));
assert.equal(doctors.find((doctor) => doctor.displayName === "Aarushi Pathania"), undefined);

const view = buildRosterView(mmcWorkbook, ddhWorkbook, richard.key);
const summary = previewSummary(view.events);

assert.equal(view.events.length, 37);
assert.equal(summary.date_range, "2026-02-09 to 2026-05-02");
assert.ok(view.reviewItems.length >= view.events.length);
assert.ok(view.events.some((event) => event.title === "Annual Leave"));
assert.ok(view.events.some((event) => event.title === "DDH: Orange PM"));
assert.ok(view.events.some((event) => event.title === "DDH: Sick Leave"));

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
assert.ok(pdfView.issues.some((issue) => issue.rawValue === "0800-1730"));

console.log("Fixture smoke test passed.");
