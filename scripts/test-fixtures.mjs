import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import XLSX from "xlsx";

import { onRequestPost as handleStatePost } from "../functions/api/state.js";
import { buildRosterView, doctorOptions, parseUploadForm, parserRuleDefaults, previewSummary } from "../public/static/roster.js";

const mmcWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/AdultTerm1.2026.xlsx", import.meta.url)), {
  cellDates: true,
});
const ddhWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/Dandenong_Emergency_Doctors_Roster_02-02-2026_to_03-05-2026.xlsx", import.meta.url)), {
  cellDates: true,
});

const doctors = doctorOptions(mmcWorkbook, ddhWorkbook);
const defaultRules = parserRuleDefaults();
const mmcRules = defaultRules.mmc || [];
const hasMmcRule = (seniority, code) => mmcRules.some((rule) => rule.seniority === seniority && rule.code === code);
assert.ok(hasMmcRule("SMS", "AGC"));
assert.ok(hasMmcRule("CMO", "AGC"));
assert.equal(hasMmcRule("Senior Registrar", "AGC"), false);
assert.equal(hasMmcRule("HMO", "AGC"), false);
assert.equal(hasMmcRule("SMS", "ACR"), false);
assert.equal(hasMmcRule("SMS", "ARR"), false);
assert.equal(hasMmcRule("SMS", "ASSR"), false);
assert.ok(hasMmcRule("Senior Registrar", "SWA"));
assert.ok(hasMmcRule("Transitional/Intermediate Registrar", "SWP"));
assert.ok(hasMmcRule("Junior Registrar", "AHJ"));
assert.ok(hasMmcRule("HMO", "PHJ"));
assert.ok(hasMmcRule("Intern", "NSSJ"));
const nssjRule = mmcRules.find((rule) => rule.seniority === "HMO" && rule.code === "NSSJ");
assert.equal(nssjRule.startTime, "23:00");
assert.equal(nssjRule.endTime, "08:30");
assert.ok(doctors.length > 100);
const richard = doctors.find((doctor) => doctor.displayName === "Richard HAYDON");
assert.ok(richard);
assert.deepEqual(richard.sourceTypes, ["mmc", "ddh"]);
assert.ok(doctors.find((doctor) => doctor.displayName === "Brianna Dawn Murphy"));
assert.ok(doctors.find((doctor) => doctor.displayName === "Patrick Tan"));
assert.equal(doctors.find((doctor) => doctor.displayName === "Aarushi Pathania"), undefined);
assert.equal(doctors.find((doctor) => doctor.displayName === "HMO MUST BE"), undefined);

const markDouglas = doctors.find((doctor) => doctor.displayName === "Mark Douglas");
assert.ok(markDouglas);
const markView = buildRosterView(mmcWorkbook, [], markDouglas.key);
assert.ok(markView.events.some((event) => event.title === "MMC: AM"));
assert.ok(markView.events.some((event) => event.title === "MMC: PM"));

const deslinAraullo = doctors.find((doctor) => doctor.displayName === "Deslin Araullo");
assert.ok(deslinAraullo);
const deslinView = buildRosterView(mmcWorkbook, [], deslinAraullo.key);
assert.ok(deslinView.events.some((event) => event.title === "MMC: Hub PM"));
assert.ok(deslinView.events.some((event) => event.title === "MMC: Swing AM"));

const view = buildRosterView(mmcWorkbook, ddhWorkbook, richard.key);
const summary = previewSummary(view.events);

assert.equal(view.events.length, 37);
assert.equal(summary.date_range, "2026-02-09 to 2026-05-02");
assert.ok(view.reviewItems.length >= view.events.length);
assert.ok(view.events.some((event) => event.title === "Annual Leave"));
assert.ok(view.events.some((event) => event.title === "DDH: Orange PM"));
assert.ok(view.events.some((event) => event.title === "DDH: Sick Leave"));

const ddhFullWorkbook = XLSX.utils.book_new();
const ddhFullSheet = XLSX.utils.aoa_to_sheet([
  ["", "Mon. Feb. 02, 2026", "Tue. Feb. 03, 2026", "Wed. Feb. 04, 2026", "Thu. Feb. 05, 2026", "Fri. Feb. 06, 2026", "Sat. Feb. 07, 2026", "Sun. Feb. 08, 2026"],
  ["Richard Haydon", "", "", "", "", "", "", ""],
  ["SENIOR MEDICAL STAFF", "", "", "", "", "", "", ""],
  ["Jim Barton", "AVAO AM", "", "Orange PM (on-call)", "AVAO PM", "Clinical Support", "", ""],
  ["", "07:30-17:00", "", "15:00-00:00", "14:30-00:00", "", "", ""],
  ["Caroline Bolt", "Orange PM (on-call)", "", "AVAO AM", "", "Orange AM IC", "", ""],
  ["", "15:00-00:00", "", "07:30-17:00", "", "08:00-18:00", "", ""],
  ["Di Flood", "CS AM", "SSU SMS", "Clinical Support", "", "HITH PM", "", ""],
  ["", "", "07:30-17:30", "", "", "", "", ""],
]);
XLSX.utils.book_append_sheet(ddhFullWorkbook, ddhFullSheet, "Sheet1");
const ddhFullDoctors = doctorOptions([], ddhFullWorkbook);
assert.ok(ddhFullDoctors.find((doctor) => doctor.displayName === "Jim Barton"));
assert.ok(ddhFullDoctors.find((doctor) => doctor.displayName === "Caroline Bolt"));
assert.ok(ddhFullDoctors.find((doctor) => doctor.displayName === "Di Flood"));
assert.equal(ddhFullDoctors.find((doctor) => doctor.displayName === "SENIOR MEDICAL STAFF"), undefined);

const jim = ddhFullDoctors.find((doctor) => doctor.displayName === "Jim Barton");
const jimView = buildRosterView([], ddhFullWorkbook, jim.key);
assert.ok(jimView.events.some((event) => event.title === "DDH: AVAO AM"));
assert.ok(jimView.events.some((event) => event.title === "DDH: Orange PM"));
assert.ok(jimView.events.some((event) => event.title === "DDH: AVAO PM"));
assert.ok(jimView.events.some((event) => event.title === "DDH: CS"));

const diFlood = ddhFullDoctors.find((doctor) => doctor.displayName === "Di Flood");
const diFloodView = buildRosterView([], ddhFullWorkbook, diFlood.key);
assert.ok(diFloodView.events.some((event) => event.title === "DDH: CS AM"));
assert.ok(diFloodView.events.some((event) => event.title === "DDH: SSU" && event.start.includes("07:30:00")));
assert.ok(diFloodView.events.some((event) => event.title === "DDH: HITH PM"));

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
assert.ok(pdfView.events.some((event) => event.rawValue === "0800-1730" && event.title === "MMC: AM"));

class MemoryStore {
  constructor() {
    this.records = new Map();
  }

  async get(key, type) {
    const value = this.records.get(key);
    if (value === undefined) return null;
    return type === "json" ? JSON.parse(value) : value;
  }

  async put(key, value) {
    this.records.set(key, String(value));
  }

  async delete(key) {
    this.records.delete(key);
  }

  async list(options = {}) {
    const prefix = options.prefix || "";
    return {
      keys: [...this.records.keys()].filter((name) => name.startsWith(prefix)).map((name) => ({ name })),
    };
  }
}

async function postState(store, payload) {
  const response = await handleStatePost({
    request: new Request("http://fixture.test/api/state", {
      method: "POST",
      body: JSON.stringify(payload),
      headers: { "content-type": "application/json" },
    }),
    env: { ROSTER_STORE: store },
  });
  const body = await response.json();
  assert.equal(response.ok, true, body.error || "state request failed");
  return body;
}

const stateStore = new MemoryStore();
const creatorPassword = "fixture-password";
await postState(stateStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
await stateStore.put("repository:index", JSON.stringify({
  files: [{
    repoId: "fixture-roster",
    id: "fixture-roster",
    name: "AdultMMCTerm2.2026.Ver1.pdf",
    sourceType: "mmc",
    active: true,
    doctors: [{
      key: "TITUS HACKMAN",
      displayName: "Titus HACKMAN",
      sourceType: "unknown",
    }],
  }],
}));
await stateStore.put("repository:file:fixture-roster", JSON.stringify({
  repoId: "fixture-roster",
  id: "fixture-roster",
  name: "AdultMMCTerm2.2026.Ver1.pdf",
  size: 12,
  lastModified: 1,
  sourceType: "mmc",
  dataUrl: "data:application/pdf;base64,cm9zdGVy",
}));

const creatorImports = await postState(stateStore, {
  action: "loadImports",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.equal(creatorImports.imports.length, 1);
assert.equal(creatorImports.imports[0].repoId, "fixture-roster");

const profileImports = await postState(stateStore, {
  action: "loadDoctorProfileImports",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: "TITUS HACKMAN::mmc",
  doctorKey: "TITUS HACKMAN",
  displayName: "Titus HACKMAN",
  sourceTypes: ["mmc"],
});
assert.equal(profileImports.imports.length, 1);
assert.equal(profileImports.imports[0].repoId, "fixture-roster");

await stateStore.delete("repository:index");
await stateStore.put("repository:file:fixture-roster", JSON.stringify({
  repoId: "fixture-roster",
  name: "AdultMMCTerm2.2026.Ver1.pdf",
  size: 12,
  lastModified: 1,
  sourceType: "mmc",
  doctors: [{
    key: "TITUS HACKMAN",
    displayName: "Titus HACKMAN",
    sourceType: "unknown",
  }],
  dataUrl: "data:application/pdf;base64,cm9zdGVy",
}));
const creatorImportsWithoutIndex = await postState(stateStore, {
  action: "loadImports",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.equal(creatorImportsWithoutIndex.imports.length, 1);
assert.equal(creatorImportsWithoutIndex.imports[0].repoId, "fixture-roster");
const profileImportsWithoutIndex = await postState(stateStore, {
  action: "loadDoctorProfileImports",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: "TITUS HACKMAN::mmc",
  doctorKey: "TITUS HACKMAN",
  displayName: "Titus HACKMAN",
  sourceTypes: ["mmc"],
});
assert.equal(profileImportsWithoutIndex.imports.length, 1);
assert.equal(profileImportsWithoutIndex.imports[0].repoId, "fixture-roster");

console.log("Fixture smoke test passed.");
