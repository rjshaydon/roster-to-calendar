import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import XLSX from "xlsx";

import { onRequestPost as handleStatePost } from "../functions/api/state.js";
import { buildRosterView, doctorOptions, parseUploadForm, parserRuleDefaults, previewSummary, setParserExtensions } from "../public/static/roster.js";

const mmcWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/AdultTerm1.2026.xlsx", import.meta.url)), {
  cellDates: true,
});
const ddhWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/Dandenong_Emergency_Doctors_Roster_02-02-2026_to_03-05-2026.xlsx", import.meta.url)), {
  cellDates: true,
});
const caseyWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/Casey_Term_2_2026_DRAFT.xlsm", import.meta.url)), {
  cellDates: true,
});
const mchWorkbook = XLSX.readFile(fileURLToPath(new URL("../fixtures/Paeds_Term_2_2026.xlsx", import.meta.url)), {
  cellDates: true,
});
const caseyBytes = await readFile(fileURLToPath(new URL("../fixtures/Casey_Term_2_2026_DRAFT.xlsm", import.meta.url)));
const mchBytes = await readFile(fileURLToPath(new URL("../fixtures/Paeds_Term_2_2026.xlsx", import.meta.url)));
const caseyFormData = new FormData();
caseyFormData.append("rosterFiles", new File([caseyBytes], "Casey_Term_2_2026_DRAFT.xlsm", { type: "application/vnd.ms-excel.sheet.macroEnabled.12" }));
const parsedCaseyUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: caseyFormData }));
assert.equal(parsedCaseyUpload.sources.casey.length, 1);
const mchFormData = new FormData();
mchFormData.append("rosterFiles", new File([mchBytes], "Paeds_Term_2_2026.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
const parsedMchUpload = await parseUploadForm(new Request("http://fixture.test/api/analyze", { method: "POST", body: mchFormData }));
assert.equal(parsedMchUpload.sources.mch.length, 1);

const doctors = doctorOptions(mmcWorkbook, ddhWorkbook, caseyWorkbook);
const defaultRules = parserRuleDefaults();
const mmcRules = defaultRules.mmc || [];
const hasMmcRule = (seniority, code) => mmcRules.some((rule) => rule.seniority === seniority && rule.code === code);
assert.ok(hasMmcRule("SMS", "AGC"));
assert.ok(hasMmcRule("CMO", "AGC"));
assert.ok(hasMmcRule("SMS", "CS"));
assert.ok(hasMmcRule("CMO", "CSO"));
assert.equal(hasMmcRule("Senior Registrar", "AGC"), false);
assert.equal(hasMmcRule("HMO", "AGC"), false);
assert.equal(hasMmcRule("Senior Registrar", "CS"), false);
assert.equal(hasMmcRule("HMO", "CSO"), false);
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
for (const sourceRules of [defaultRules.ddh, defaultRules.casey]) {
  assert.equal(sourceRules.some((rule) => rule.seniority === "Senior Registrar" && rule.base === "CS"), false);
  assert.equal(sourceRules.some((rule) => rule.seniority === "HMO" && rule.base === "CS"), false);
  assert.ok(sourceRules.some((rule) => rule.seniority === "SMS" && rule.base === "CS"));
  assert.ok(sourceRules.some((rule) => rule.seniority === "CMO" && rule.base === "CS"));
}
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

const caseyDoctors = doctorOptions([], [], caseyWorkbook);
assert.ok(caseyDoctors.length > 130);
assert.ok(caseyDoctors.find((doctor) => doctor.displayName === "Andrew Dyall"));
assert.ok(caseyDoctors.find((doctor) => doctor.displayName === "Dennis Chung"));
assert.ok(caseyDoctors.find((doctor) => doctor.displayName === "Rizwana Sadaf"));
assert.ok(caseyDoctors.find((doctor) => doctor.displayName === "Victor Ki Chung Li"));
assert.equal(caseyDoctors.find((doctor) => doctor.displayName === "Rostered staff"), undefined);

const patrickTan = doctorOptions(mmcWorkbook, [], caseyWorkbook).find((doctor) => doctor.displayName === "Patrick Tan");
assert.ok(patrickTan);
assert.deepEqual(patrickTan.sourceTypes, ["mmc", "casey"]);
const patrickTanView = buildRosterView(mmcWorkbook, [], patrickTan.key, undefined, {}, {}, [], caseyWorkbook);
const patrickCaseyEvents = patrickTanView.events.filter((event) => event.source === "Casey");
assert.ok(patrickCaseyEvents.length > 40);
assert.equal(patrickCaseyEvents.some((event) => event.start.startsWith("2025")), false);
assert.ok(patrickCaseyEvents.some((event) => event.title === "Casey: MIC PM"));
assert.ok(patrickCaseyEvents.some((event) => event.title === "Casey: AM" && event.rawValue === "Orient 0800-1730" && event.start.includes("08:00:00") && event.end.includes("17:30:00")));

const andrewDyallCasey = caseyDoctors.find((doctor) => doctor.displayName === "Andrew Dyall");
const andrewCaseyView = buildRosterView([], [], andrewDyallCasey.key, undefined, {}, {}, [], caseyWorkbook);
assert.ok(andrewCaseyView.events.some((event) => event.title === "Casey: TL AM"));
assert.ok(andrewCaseyView.events.some((event) => event.title === "Casey: UFD PM"));
assert.ok(andrewCaseyView.events.some((event) => event.title === "Casey: MIC AM"));
assert.ok(andrewCaseyView.events.some((event) => event.title === "Casey: PAEDS PM"));
assert.ok(andrewCaseyView.events.some((event) => event.title === "Casey: CS" && event.start.includes("08:00:00") && event.end.includes("17:30:00")));

const bashirCasey = caseyDoctors.find((doctor) => doctor.displayName === "Bashir Gondal");
const bashirCaseyView = buildRosterView([], [], bashirCasey.key, undefined, {}, {}, [], caseyWorkbook);
assert.ok(bashirCaseyView.events.some((event) => event.title === "Casey: SSU AM"));

const dennisCasey = caseyDoctors.find((doctor) => doctor.displayName === "Dennis Chung");
const dennisCaseyView = buildRosterView([], [], dennisCasey.key, undefined, {}, {}, [], caseyWorkbook);
assert.ok(dennisCaseyView.events.some((event) => event.title === "Casey: Night" && event.start.includes("23:00:00") && event.end.startsWith("2026-05-06")));
assert.equal(dennisCaseyView.events.filter((event) => event.title === "Casey: Night" && event.start.startsWith("2026-05-05") && event.end.startsWith("2026-05-06")).length, 1);

const jasonAwCasey = caseyDoctors.find((doctor) => doctor.displayName === "Jason Aw");
const jasonAwCaseyView = buildRosterView([], [], jasonAwCasey.key, undefined, {}, {}, [], caseyWorkbook);
assert.ok(jasonAwCaseyView.events.some((event) => event.title === "Annual Leave"));

const mustafaCasey = caseyDoctors.find((doctor) => doctor.displayName === "Mustafa Al-Asaad");
const mustafaCaseyView = buildRosterView([], [], mustafaCasey.key, undefined, {}, {}, [], caseyWorkbook);
assert.ok(mustafaCaseyView.events.some((event) => event.title === "Conference Leave"));

const mchDoctors = doctorOptions([], [], [], mchWorkbook);
assert.ok(mchDoctors.length >= 60);
assert.ok(mchDoctors.find((doctor) => doctor.displayName === "Dr Adam West"));
assert.ok(mchDoctors.find((doctor) => doctor.displayName === "Mark Lim"));
assert.ok(mchDoctors.find((doctor) => doctor.displayName === "Firas Hamdan"));
assert.ok(mchDoctors.find((doctor) => doctor.displayName === "Peter Ahn"));
assert.equal(mchDoctors.find((doctor) => doctor.displayName === "ONCALL 0000-0800"), undefined);
assert.equal(mchDoctors.find((doctor) => doctor.displayName === "requested off"), undefined);

const adamWestMch = mchDoctors.find((doctor) => doctor.displayName === "Dr Adam West");
const adamWestMchView = buildRosterView([], [], adamWestMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(adamWestMchView.events.some((event) => event.title === "MCH: CS" && event.rawValue === "0800-1730 CS"));
assert.ok(adamWestMchView.events.some((event) => event.title === "MCH: PM" && event.rawValue === "1430-0000" && event.end.startsWith("2026-05-09")));

const bobSeithMch = mchDoctors.find((doctor) => doctor.displayName === "Dr Bob Seith");
const bobSeithMchView = buildRosterView([], [], bobSeithMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(bobSeithMchView.events.some((event) => event.title === "MCH: DEMT" && event.rawValue === "0800-1730 DEMT"));
assert.ok(bobSeithMchView.events.some((event) => event.title === "MCH: CS" && event.rawValue === "0800-1730CS"));

const andrewHardyMch = mchDoctors.find((doctor) => doctor.displayName === "Dr Andrew Hardy");
const andrewHardyMchView = buildRosterView([], [], andrewHardyMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(andrewHardyMchView.events.some((event) => event.title === "MCH: OCS" && event.rawValue === "0800-1730 OCS"));
assert.ok(andrewHardyMchView.events.some((event) => event.title === "MCH: Exam Leave" && event.rawValue === "ME/L" && event.allDay));
assert.ok(andrewHardyMchView.events.some((event) => event.title === "Conference Leave" && event.rawValue === "CME/L" && event.allDay));
assert.ok(andrewHardyMchView.events.some((event) => event.title === "Conference Leave" && event.rawValue === "CME/L" && event.start === "2026-06-08" && event.end === "2026-06-15"));

const adamWestMchWeek6 = adamWestMchView.events.filter((event) => event.rawValue === "PHNW 0800-1730");
assert.ok(adamWestMchWeek6.some((event) => event.title === "MCH: PHNW"));

const markLimMch = mchDoctors.find((doctor) => doctor.displayName === "Mark Lim");
const markLimMchView = buildRosterView([], [], markLimMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(markLimMchView.events.some((event) => event.title === "MCH: Night" && event.rawValue === "2300-0830" && event.end.startsWith("2026-05-09")));
assert.equal(markLimMchView.events.filter((event) => event.title === "MCH: Night" && event.rawValue === "2300-0830" && event.start.startsWith("2026-05-08") && event.end.startsWith("2026-05-09")).length, 1);
assert.ok(markLimMchView.events.some((event) => event.title === "Conference Leave" && event.rawValue === "C/L" && event.allDay));

const firasMch = mchDoctors.find((doctor) => doctor.displayName === "Firas Hamdan");
const firasMchView = buildRosterView([], [], firasMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(firasMchView.events.some((event) => event.title === "Annual Leave" && event.rawValue === "AL 0.5" && event.allDay));

const marianPanlilioMch = mchDoctors.find((doctor) => doctor.displayName === "Marian Panlilio");
const marianPanlilioMchView = buildRosterView([], [], marianPanlilioMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.ok(marianPanlilioMchView.events.some((event) => event.title === "MCH: Sick Leave PM" && event.rawValue.trim() === "S/L PM" && event.allDay));

const houshmandMch = mchDoctors.find((doctor) => doctor.displayName === "Houshmand Refaei");
const houshmandMchView = buildRosterView([], [], houshmandMch.key, undefined, {}, {}, [], [], mchWorkbook);
assert.equal(houshmandMchView.events.some((event) => String(event.rawValue || "").includes("EDO")), false);

const overlappingConferenceWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(overlappingConferenceWorkbook, XLSX.utils.aoa_to_sheet([
  ["TERM 2, 2026", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
  ["", "20-Jul", "21-Jul", "22-Jul", "23-Jul", "24-Jul", "25-Jul", "26-Jul"],
  ["Dr Michael Coman", "C/L", "C/L", "C/L", "C/L", "C/L", "C/L", "C/L"],
  ["Dr Daily Leave", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave", "Annual Leave"],
]), "Week 1");
const michaelComan = doctorOptions([], [], overlappingConferenceWorkbook, mchWorkbook).find((doctor) => doctor.displayName === "Dr Michael Coman");
assert.ok(michaelComan);
const michaelComanView = buildRosterView([], [], michaelComan.key, undefined, {}, {}, [], overlappingConferenceWorkbook, mchWorkbook);
const michaelConferenceEvents = michaelComanView.events.filter((event) => event.title === "Conference Leave" && event.start === "2026-07-20");
assert.equal(michaelConferenceEvents.length, 1);
assert.equal(michaelConferenceEvents[0].end, "2026-07-27");
assert.equal(michaelConferenceEvents[0].rawValue, "C/L / CME/L");
assert.ok(michaelComanView.reviewItems.some((item) => item.id === michaelConferenceEvents[0].id));

const dailyLeave = doctorOptions([], [], overlappingConferenceWorkbook).find((doctor) => doctor.displayName === "Dr Daily Leave");
assert.ok(dailyLeave);
const dailyLeaveView = buildRosterView([], [], dailyLeave.key, undefined, {}, {}, [], overlappingConferenceWorkbook);
const dailyAnnualLeave = dailyLeaveView.events.filter((event) => event.title === "Annual Leave");
assert.equal(dailyAnnualLeave.length, 1);
assert.equal(dailyAnnualLeave[0].start, "2026-07-20");
assert.equal(dailyAnnualLeave[0].end, "2026-07-27");
assert.ok(dailyLeaveView.reviewItems.some((item) => item.id === dailyAnnualLeave[0].id));

const view = buildRosterView(mmcWorkbook, ddhWorkbook, richard.key);
const summary = previewSummary(view.events);

const aftabMmc = doctors.find((doctor) => doctor.displayName === "Aftab SAMDANI");
assert.ok(aftabMmc);
const aftabMmcView = buildRosterView(mmcWorkbook, [], aftabMmc.key);
assert.ok(aftabMmcView.events.some((event) => event.title === "Conference Leave" && event.rawValue.toUpperCase() === "CME LEAVE" && event.allDay));

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
    this.deletedKeys = [];
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
    this.deletedKeys.push(key);
    this.records.delete(key);
  }

  async list(options = {}) {
    const prefix = options.prefix || "";
    return {
      keys: [...this.records.keys()].filter((name) => name.startsWith(prefix)).map((name) => ({ name })),
    };
  }
}

function repositoryFile(id, overrides = {}) {
  return {
    repoId: id,
    id,
    name: `${id}.xlsx`,
    sourceType: "mmc",
    active: true,
    size: 12,
    lastModified: 1,
    doctors: [{
      key: "TITUS HACKMAN",
      displayName: "Titus HACKMAN",
      sourceType: "mmc",
    }],
    ...overrides,
  };
}

async function seedRepository(store, files) {
  await store.put("repository:index", JSON.stringify({ version: 1, files }));
  for (const file of files) {
    await store.put(`repository:file:${file.id}`, JSON.stringify({
      ...file,
      dataUrl: `data:application/octet-stream;base64,${Buffer.from(file.id).toString("base64")}`,
    }));
  }
}

async function seedUser(store, email, password, realName = "Titus Hackman") {
  await postState(store, {
    action: "login",
    email,
    password,
    mode: "create",
    realName,
  });
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
await seedRepository(stateStore, [repositoryFile("fixture-roster", {
  name: "AdultMMCTerm2.2026.Ver1.pdf",
  sourceType: "mmc",
})]);

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

await seedUser(stateStore, "patrick@example.com", "patrick-password", "Patrick Tan");
await seedUser(stateStore, "senior@example.com", "senior-password", "Senior Registrar");
const initialUsers = await postState(stateStore, {
  action: "listUsers",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.equal(initialUsers.users.find((user) => user.email === "patrick@example.com")?.insightsEnabled, false);
const enabledInsights = await postState(stateStore, {
  action: "setUserInsightsEnabled",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "patrick@example.com",
  insightsEnabled: true,
});
assert.equal(enabledInsights.user.insightsEnabled, true);
const patrickLogin = await postState(stateStore, {
  action: "login",
  email: "patrick@example.com",
  password: "patrick-password",
});
assert.equal(patrickLogin.insightsEnabled, true);
const seniorLogin = await postState(stateStore, {
  action: "login",
  email: "senior@example.com",
  password: "senior-password",
});
assert.equal(seniorLogin.insightsEnabled, false);
const n1Issue = {
  source: "MMC",
  seniority: "Senior Registrar",
  date: "2026-05-01",
  rawValue: "N1",
  message: "MMC shift code not recognised.",
  fingerprint: "MMC::Senior Registrar::N1",
};
await postState(stateStore, {
  action: "reportUserError",
  email: "patrick@example.com",
  password: "patrick-password",
  errorId: n1Issue.fingerprint,
  message: n1Issue.message,
  issue: n1Issue,
});
await postState(stateStore, {
  action: "reportUserError",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  targetEmail: "senior@example.com",
  errorId: n1Issue.fingerprint,
  message: n1Issue.message,
  issue: n1Issue,
});
assert.equal((await stateStore.get("account:patrick@example.com", "json")).adminIssues.length, 1);
assert.equal((await stateStore.get("account:senior@example.com", "json")).adminIssues.length, 1);
const parserSave = await postState(stateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  fingerprint: n1Issue.fingerprint,
  source: "MMC",
  rawValue: "N1",
  rule: {
    source: "MMC",
    seniority: "Senior Registrar",
    code: "N1",
    kind: "shift",
    base: "SR IC",
    period: "NIGHT",
    suffix: "",
    allDay: false,
    startTime: "23:00",
    endTime: "09:00",
    includeAsShift: true,
  },
});
assert.ok(parserSave.parserExtensions.mmc.some((rule) => rule.seniority === "Senior Registrar" && rule.code === "N1"));
setParserExtensions(parserSave.parserExtensions);
const srN1Workbook = XLSX.utils.book_new();
const srN1Sheet = XLSX.utils.aoa_to_sheet([
  [],
  [],
  [],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "SENIOR REG"],
  ["", "", "", "Patrick Tan", "", "2300-0900 N1"],
]);
for (let index = 0; index < 7; index += 1) {
  srN1Sheet[XLSX.utils.encode_cell({ r: 3, c: 5 + index })] = { t: "d", v: new Date(`2026-05-${String(4 + index).padStart(2, "0")}T00:00:00`) };
}
XLSX.utils.book_append_sheet(srN1Workbook, XLSX.utils.aoa_to_sheet([[]]), "Whole thing");
XLSX.utils.book_append_sheet(srN1Workbook, srN1Sheet, "Week 1");
const srN1View = buildRosterView([{ id: "sr-n1", workbook: srN1Workbook, file: { name: "AdultTerm.xlsx", size: 1, lastModified: 1 } }], [], "PATRICK TAN");
assert.ok(srN1View.events.some((event) => event.rawValue === "2300-0900 N1" && event.title === "MMC: SR IC Night" && event.start.includes("23:00:00") && event.end.includes("09:00:00")), "Senior Registrar N1 explicit-time rules must render with the saved rule title");
assert.equal(srN1View.issues.some((issue) => issue.rawValue === "2300-0900 N1"), false);
assert.equal((await stateStore.get("account:patrick@example.com", "json")).adminIssues.length, 0, "global parser rule must clear direct-user warnings");
assert.equal((await stateStore.get("account:senior@example.com", "json")).adminIssues.length, 0, "global parser rule must clear switch-user warnings");
const staleReport = await postState(stateStore, {
  action: "reportUserError",
  email: "patrick@example.com",
  password: "patrick-password",
  errorId: "MMC::Senior Registrar::2200-0830 N1",
  message: n1Issue.message,
  issue: {
    ...n1Issue,
    rawValue: "2200-0830 N1",
    fingerprint: "MMC::Senior Registrar::2200-0830 N1",
  },
});
assert.equal(staleReport.ignored, true, "resolved global shift-code warnings must not be requeued from stale user previews");
assert.equal((await stateStore.get("account:patrick@example.com", "json")).adminIssues.length, 0, "resolved global shift-code warnings must not return after user login");
const accParserSave = await postState(stateStore, {
  action: "saveParserExtensionRule",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  fingerprint: "MMC::Senior Registrar::ACC",
  source: "MMC",
  rawValue: "ACC",
  rule: {
    source: "MMC",
    seniority: "Senior Registrar",
    code: "ACC",
    kind: "shift",
    base: "Clinic",
    period: "AM",
    suffix: "Charge",
    allDay: false,
    startTime: "08:00",
    endTime: "17:30",
    includeAsShift: true,
  },
});
assert.ok(accParserSave.parserExtensions.mmc.some((rule) => rule.seniority === "Senior Registrar" && rule.code === "ACC"), "Senior Registrar charge/consultant-style codes must persist as explicit rules");
setParserExtensions(accParserSave.parserExtensions);
const srAccWorkbook = XLSX.utils.book_new();
const srAccSheet = XLSX.utils.aoa_to_sheet([
  [],
  [],
  [],
  ["", "", "", "", "", "", "", "", "", "", "", ""],
  ["", "", "SENIOR REG"],
  ["", "", "", "Patrick Tan", "", "ACC"],
]);
for (let index = 0; index < 7; index += 1) {
  srAccSheet[XLSX.utils.encode_cell({ r: 3, c: 5 + index })] = { t: "d", v: new Date(`2026-05-${String(4 + index).padStart(2, "0")}T00:00:00`) };
}
XLSX.utils.book_append_sheet(srAccWorkbook, XLSX.utils.aoa_to_sheet([[]]), "Whole thing");
XLSX.utils.book_append_sheet(srAccWorkbook, srAccSheet, "Week 1");
const srAccView = buildRosterView([{ id: "sr-acc", workbook: srAccWorkbook, file: { name: "AdultTerm.xlsx", size: 1, lastModified: 1 } }], [], "PATRICK TAN");
assert.ok(srAccView.events.some((event) => event.rawValue === "ACC" && event.title === "MMC: Clinic AM Charge"), "Senior Registrar explicit ACC rules must render in user calendars");
assert.equal(srAccView.issues.some((issue) => issue.rawValue === "ACC"), false);
setParserExtensions({});

const deletionStore = new MemoryStore();
await postState(deletionStore, {
  action: "login",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
await seedRepository(deletionStore, [
  repositoryFile("keep-roster"),
  repositoryFile("missing-from-save", { name: "missing-from-save.xlsx" }),
  repositoryFile("remove-roster", { name: "remove-roster.xlsx" }),
]);

await postState(deletionStore, {
  action: "save",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  state: {
    version: 1,
    imports: [{ repoId: "keep-roster", id: "keep-roster", sourceType: "mmc", name: "keep-roster.xlsx" }],
    session: {},
  },
});
assert.ok(await deletionStore.get("repository:file:missing-from-save", "json"), "ordinary creator save must not delete omitted repository files");
let deletionIndex = await deletionStore.get("repository:index", "json");
assert.ok(deletionIndex.files.some((file) => file.id === "missing-from-save"), "ordinary creator save must keep omitted files in the repository index");

await postState(deletionStore, {
  action: "save",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  state: {
    version: 1,
    imports: [{ repoId: "keep-roster", id: "keep-roster", sourceType: "mmc", name: "keep-roster.xlsx" }],
    session: {},
  },
  removedImportIds: ["remove-roster"],
});
assert.equal(await deletionStore.get("repository:file:remove-roster", "json"), null);
deletionIndex = await deletionStore.get("repository:index", "json");
assert.equal(deletionIndex.files.some((file) => file.id === "remove-roster"), false);
assert.ok(await deletionStore.get("repository:file:keep-roster", "json"));
assert.ok(await deletionStore.get("repository:file:missing-from-save", "json"));

await postState(deletionStore, {
  action: "save",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  state: {
    version: 1,
    imports: [{ repoId: "keep-roster", id: "keep-roster", sourceType: "mmc", name: "keep-roster.xlsx" }],
    session: {},
  },
  removedImportIds: ["remove-roster"],
});
assert.ok(await deletionStore.get("repository:file:keep-roster", "json"));
assert.ok(await deletionStore.get("repository:file:missing-from-save", "json"));

await seedUser(deletionStore, "user@example.com", "user-password");
await postState(deletionStore, {
  action: "save",
  email: "user@example.com",
  password: "user-password",
  state: {
    version: 1,
    imports: [{ repoId: "keep-roster", id: "keep-roster", sourceType: "mmc", name: "keep-roster.xlsx" }],
    session: {},
  },
  removedImportIds: ["keep-roster"],
});
assert.ok(await deletionStore.get("repository:file:keep-roster", "json"), "standard users must not delete repository files");

const beforeLoadDeleteCount = deletionStore.deletedKeys.length;
const deletionCreatorImports = await postState(deletionStore, {
  action: "loadImports",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
});
assert.equal(deletionStore.deletedKeys.length, beforeLoadDeleteCount, "loading imports must not delete repository records");
assert.equal(deletionCreatorImports.imports.some((item) => item.repoId === "remove-roster"), false);
assert.ok(deletionCreatorImports.imports.some((item) => item.repoId === "keep-roster"));
assert.ok(deletionCreatorImports.imports.some((item) => item.repoId === "missing-from-save"));

const deletionProfileImports = await postState(deletionStore, {
  action: "loadDoctorProfileImports",
  email: "rhaydon@gmail.com",
  password: creatorPassword,
  profileId: "TITUS HACKMAN::mmc",
  doctorKey: "TITUS HACKMAN",
  displayName: "Titus HACKMAN",
  sourceTypes: ["mmc"],
});
assert.equal(deletionProfileImports.imports.some((item) => item.repoId === "remove-roster"), false);

const deletionUserImports = await postState(deletionStore, {
  action: "loadImports",
  email: "user@example.com",
  password: "user-password",
});
assert.equal(deletionUserImports.imports.some((item) => item.repoId === "remove-roster"), false);

console.log("Fixture smoke test passed.");
