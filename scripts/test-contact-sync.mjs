import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

import { onRequestPost as rejectBinaryContactUpload } from "../functions/api/automation/contact-list-binary.js";
import { onRequestPost as rejectLegacyContactUpload } from "../functions/api/automation/contact-list.js";
import { automationSourceDate, onRequestPost as acceptContactExtract } from "../functions/api/automation/contact-list-extract.js";

assert.equal(automationSourceDate("Tuesday 25th AUGUST 2026"), "2026-08-25");
assert.equal(automationSourceDate("2026-08-25"), "2026-08-25");

const vhhCredential = await acceptContactExtract({
  request: new Request("https://example.test/api/automation/contact-list-extract", {
    method: "POST",
    headers: { authorization: "Bearer vhh-only-token", "content-type": "application/json" },
    body: JSON.stringify({ sourceId: "mmc-shift-allocations", sourceDate: "2026-08-25", contacts: [] }),
  }),
  env: { ROSTER_AUTOMATION_TOKEN: "roster-token", DDH_CONTACT_AUTOMATION_TOKEN: "ddh-token", VHH_AUTOMATION_TOKEN: "vhh-only-token" },
});
assert.equal(vhhCredential.status, 401, "the VHH roster credential must not authorize any contact-list submission");

const authorized = await rejectBinaryContactUpload({
  request: new Request("https://example.test/api/automation/contact-list-binary", {
    method: "POST",
    headers: { authorization: "Bearer test-token" },
    body: "not-a-workbook",
  }),
  env: { ROSTER_AUTOMATION_TOKEN: "test-token" },
});
assert.equal(authorized.status, 410, "authenticated full-workbook MMC uploads must be disabled");
assert.equal((await authorized.json()).endpoint, "/api/automation/contact-list-extract");

const unauthorized = await rejectBinaryContactUpload({
  request: new Request("https://example.test/api/automation/contact-list-binary", { method: "POST", body: "x" }),
  env: { ROSTER_AUTOMATION_TOKEN: "test-token" },
});
assert.equal(unauthorized.status, 401, "the retired endpoint must not disclose automation details without a valid token");

const legacyWorkbook = await rejectLegacyContactUpload({
  request: new Request("https://example.test/api/automation/contact-list", {
    method: "POST", headers: { authorization: "Bearer test-token" }, body: "not-a-workbook",
  }),
  env: { ROSTER_AUTOMATION_TOKEN: "test-token" },
});
assert.equal(legacyWorkbook.status, 410, "the original full-workbook route must also be disabled");
assert.equal((await legacyWorkbook.json()).endpoint, "/api/automation/contact-list-extract");

const contactExtractSource = await readFile(new URL("../functions/api/automation/contact-list-extract.js", import.meta.url), "utf8");
const stateSource = await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8");
const appSource = await readFile(new URL("../public/static/app.js", import.meta.url), "utf8");

assert.doesNotMatch(contactExtractSource, /WHERE source_id = \? AND provider_version = \?/,
  "a repeated provider version must not hide changed JSON contacts");
assert.match(contactExtractSource, /const contentHash = await sha256Hex\(bytes\)[\s\S]*matchingHash/,
  "MMC JSON ingestion should deduplicate only identical extracts");
assert.match(contactExtractSource, /automationSourceDate[\s\S]*st\|nd\|rd\|th/,
  "the JSON boundary should normalize the date label emitted by the existing MMC Office Script");
assert.match(contactExtractSource, /pruneStoredContactExtracts[\s\S]*contactExtractHasExpired/,
  "the previous operational day's JSON should be retained until its testing-window expiry");
assert.doesNotMatch(contactExtractSource, /VHH_AUTOMATION_TOKEN|VHH_CONTACT_LIST_SOURCE_ID/,
  "VHH roster credentials and contact payloads must remain outside contact ingestion until VHH contacts are approved");
assert.match(stateSource, /action === "queryFacilityOverviewContactList"[\s\S]*loadLiveContactListForOnShift/,
  "the UI should have a lightweight contact-only refresh action");
assert.match(stateSource, /content_type[\s\S]*reason: "legacy-workbook"/,
  "stored full workbooks must produce an explicit legacy-feed status");
assert.match(stateSource, /LIMIT 8[\s\S]*extract\.sourceDate === date/,
  "On shift should select the retained JSON extract for the requested operational date");
assert.match(stateSource, /code === "DDH"[\s\S]*DDH_CONTACT_LIST_SOURCE_ID[\s\S]*code === "MMC" \|\| code === "MCH"[\s\S]*MMC_CONTACT_LIST_SOURCE_ID/,
  "MMC, MCH and DDH must retain their established contact-feed mappings");
assert.doesNotMatch(
  stateSource.match(/async function loadLiveContactListForOnShift[\s\S]*?function sanitizeFindmyshiftHistoricalRange/)?.[0] || "",
  /code === "VHH"/,
  "VHH must not expose contact details in On shift until its mapping is explicitly approved",
);
assert.match(appSource, /FACILITY_OVERVIEW_CONTACT_REFRESH_MS = 10_000[\s\S]*refreshFacilityOverviewContactList[\s\S]*queryFacilityOverviewContactList/,
  "an open On shift view should poll the small JSON contact feed");
assert.match(appSource, /legacy-workbook[\s\S]*full Excel workbook instead of the doctors-only JSON extract/,
  "legacy MMC uploads should be visible rather than silently hidden");

console.log("Contact sync safeguards passed.");
