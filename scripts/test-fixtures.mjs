import assert from "node:assert/strict";
import { fileURLToPath } from "node:url";
import XLSX from "xlsx";

import { buildRosterView, doctorOptions, previewSummary } from "../functions/_lib/roster.js";

const mmcWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/AdultTerm1.2026.xlsx", import.meta.url)), {
  cellDates: true,
});
const ddhWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/Dandenong_Emergency_Doctors_Roster_02-02-2026_to_03-05-2026.xlsx", import.meta.url)), {
  cellDates: true,
});

const doctors = doctorOptions(mmcWorkbook, ddhWorkbook);
assert.equal(doctors.length, 1);
assert.equal(doctors[0].displayName, "Richard HAYDON");

const view = buildRosterView(mmcWorkbook, ddhWorkbook, doctors[0].key);
const summary = previewSummary(view.events);

assert.equal(view.events.length, 37);
assert.equal(summary.date_range, "2026-02-09 to 2026-05-02");
assert.ok(view.reviewItems.length >= view.events.length);
assert.ok(view.events.some((event) => event.title === "MMC: Annual Leave"));
assert.ok(view.events.some((event) => event.title === "DDH: Orange PM"));
assert.ok(view.events.some((event) => event.title === "DDH: Sick Leave"));

console.log("Fixture smoke test passed.");
