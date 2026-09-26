import assert from "node:assert/strict";
import { readFile, readdir } from "node:fs/promises";
import { DatabaseSync } from "node:sqlite";
import XLSX from "xlsx";

import { onRequestPost as ingestContactWorkbook } from "../functions/api/automation/contact-workbook-extract.js";

class LocalD1 {
  constructor(sqlite) { this.sqlite = sqlite; this.rowsWritten = 0; }
  prepare(sql) {
    const owner = this;
    return {
      args: [],
      bind(...args) { this.args = args; return this; },
      async run() {
        const result = owner.sqlite.prepare(sql).run(...this.args);
        owner.rowsWritten += Number(result.changes || 0);
        return { success: true, meta: { changes: Number(result.changes || 0), rows_read: 0, rows_written: Number(result.changes || 0) } };
      },
      async all() { return { success: true, results: owner.sqlite.prepare(sql).all(...this.args), meta: { rows_read: 0, rows_written: 0 } }; },
      async first() { return owner.sqlite.prepare(sql).get(...this.args) || null; },
    };
  }
  async batch(statements) { return Promise.all(statements.map((statement) => statement.run())); }
}

class LocalR2 {
  constructor() { this.objects = new Map(); this.puts = 0; }
  async put(key, value) {
    this.puts += 1;
    const bytes = value instanceof Uint8Array ? value : new Uint8Array(value);
    this.objects.set(key, bytes);
  }
  async get(key) {
    const bytes = this.objects.get(key);
    if (!bytes) return null;
    return {
      arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
      text: async () => new TextDecoder().decode(bytes),
    };
  }
  async delete(key) { this.objects.delete(key); }
}

const workbookBytes = mmcWorkbookBytes();
const forbiddenDb = new Proxy({}, { get() { throw new Error("D1 must not be touched"); } });
const baseHeaders = {
  authorization: "Bearer contact-token",
  "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "x-contact-source-id": "mmc-shift-allocations",
  "x-contact-file-name": "SHIFT ALLOCATIONS.xlsx",
  "x-provider-modified-at": "2026-08-25T01:00:00Z",
  "x-provider-version": "etag-1",
};

assert.equal((await call(workbookBytes, {}, { ...baseHeaders, authorization: "Bearer wrong" })).status, 401,
  "invalid credentials must be rejected before D1");
assert.equal((await call(workbookBytes, {}, { ...baseHeaders, "x-contact-source-id": "unknown" })).status, 400,
  "unknown sources must be rejected before D1");
assert.equal((await call(workbookBytes, {}, { ...baseHeaders, "x-contact-file-name": "Other.xlsx" })).status, 400,
  "an unexpected filename must be rejected before D1");
assert.equal((await call(workbookBytes, {}, { ...baseHeaders, "content-type": "image/png" })).status, 415,
  "non-workbook content must be rejected before D1");
assert.equal((await call(new Uint8Array([1]), {
  CONTACT_AUTOMATION_WRITES_ENABLED: "true",
  CONTACT_AUTOMATION_SOURCE_ALLOWLIST: "mmc-shift-allocations",
}, { ...baseHeaders, "content-length": String(5 * 1024 * 1024 + 1) })).status, 413,
  "an oversized declared body must be rejected before D1");
assert.equal((await call(new TextEncoder().encode("must not be parsed"), {
  ROSTER_DB: forbiddenDb,
  CONTACT_AUTOMATION_WRITES_ENABLED: "false",
  CONTACT_AUTOMATION_SOURCE_ALLOWLIST: "",
}, baseHeaders)).status, 503, "a disabled source must stop before D1");

const sqlite = new DatabaseSync(":memory:");
for (const name of (await readdir(new URL("../migrations", import.meta.url))).filter((name) => name.endsWith(".sql")).sort()) {
  sqlite.exec(await readFile(new URL(`../migrations/${name}`, import.meta.url), "utf8"));
}
const db = new LocalD1(sqlite);
const r2 = new LocalR2();
const enabledEnv = {
  ROSTER_DB: db,
  ROSTER_FILES: r2,
  CONTACT_AUTOMATION_WRITES_ENABLED: "true",
  CONTACT_AUTOMATION_SOURCE_ALLOWLIST: "mmc-shift-allocations",
  FACILITY_SHARED_CONTACTS_BUILD_ENABLED: "false",
};
const stored = await call(workbookBytes, enabledEnv, baseHeaders);
assert.equal(stored.status, 200);
assert.equal((await stored.json()).status, "stored");
assert.equal(sqlite.prepare("SELECT COUNT(*) AS count FROM contact_list_files").get().count, 1);
assert.equal(r2.puts, 1, "a new extract should store one small JSON source object");
const writesAfterStored = db.rowsWritten;
const putsAfterStored = r2.puts;

const unchanged = await call(workbookBytes, enabledEnv, {
  ...baseHeaders,
  "x-provider-modified-at": "2026-08-25T01:05:00Z",
  "x-provider-version": "etag-2",
});
assert.equal(unchanged.status, 200);
assert.equal((await unchanged.json()).status, "unchanged");
assert.equal(db.rowsWritten, writesAfterStored, "metadata-only workbook replay must write zero D1 rows");
assert.equal(r2.puts, putsAfterStored, "metadata-only workbook replay must write zero R2 objects");
assert.equal(sqlite.prepare("SELECT COUNT(*) AS count FROM contact_list_files").get().count, 1);

const powerAutomateEnvelope = new TextEncoder().encode(JSON.stringify({
  "$content-type": "application/octet-stream",
  "$content": toBase64(workbookBytes),
}));
const envelopeReplay = await call(powerAutomateEnvelope, enabledEnv, {
  ...baseHeaders,
  "content-type": "application/json",
  "x-provider-modified-at": "2026-08-25T01:10:00Z",
  "x-provider-version": "etag-3",
});
assert.equal(envelopeReplay.status, 200, "Power Automate's connector envelope must be accepted");
assert.equal((await envelopeReplay.json()).status, "unchanged");
assert.equal(db.rowsWritten, writesAfterStored, "an encoded unchanged replay must write zero D1 rows");
assert.equal(r2.puts, putsAfterStored, "an encoded unchanged replay must write zero R2 objects");

const textReplay = await call(new TextEncoder().encode(toBase64(workbookBytes)), enabledEnv, {
  ...baseHeaders,
  "content-type": "text/plain; charset=utf-8",
  "x-provider-modified-at": "2026-08-25T01:15:00Z",
  "x-provider-version": "etag-4",
});
assert.equal(textReplay.status, 200, "Power Automate's base64 text representation must be accepted");
assert.equal((await textReplay.json()).status, "unchanged");
assert.equal(db.rowsWritten, writesAfterStored);
assert.equal(r2.puts, putsAfterStored);

const originalConsoleError = console.error;
console.error = () => {};
const invalidEnvelope = await call(new TextEncoder().encode(JSON.stringify({ value: "not-a-workbook" })), enabledEnv, {
  ...baseHeaders,
  "content-type": "application/json",
});
assert.equal(invalidEnvelope.status, 422, "an invalid connector envelope must fail without storing data");
const malformed = await call(new TextEncoder().encode("not an xlsx"), enabledEnv, baseHeaders);
console.error = originalConsoleError;
assert.equal(malformed.status, 422, "a malformed workbook must fail without storing data");
assert.equal(db.rowsWritten, writesAfterStored);
assert.equal(r2.puts, putsAfterStored);

console.log("Contact workbook ingress safeguards passed.");

function call(body, envOverrides = {}, headers = baseHeaders) {
  return ingestContactWorkbook({
    request: new Request("https://example.test/api/automation/contact-workbook-extract", { method: "POST", headers, body }),
    env: {
      ROSTER_AUTOMATION_TOKEN: "contact-token",
      DDH_CONTACT_AUTOMATION_TOKEN: "ddh-token",
      ...envOverrides,
    },
  });
}

function mmcWorkbookBytes() {
  const workbook = XLSX.utils.book_new();
  const rows = Array.from({ length: 42 }, () => Array(9).fill(""));
  rows[1][3] = "25th August 2026";
  rows[5][0] = "CART Clinician";
  rows[5][1] = "Casey";
  rows[5][2] = "0417 489 358";
  rows[30][0] = "Paeds Dr - CIC/AO";
  rows[30][1] = "Simon";
  rows[30][2] = "25145";
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), "SHIFT ALLOCATIONS");
  return XLSX.write(workbook, { type: "array", bookType: "xlsx" });
}

function toBase64(bytes) {
  let binary = "";
  for (const byte of new Uint8Array(bytes)) binary += String.fromCharCode(byte);
  return btoa(binary);
}
