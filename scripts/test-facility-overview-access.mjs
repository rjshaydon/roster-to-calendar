import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { resolveFacilityOverviewAccess } from "../functions/api/state.js";

const claim = (sourceType, key = "TEST DOCTOR") => ({ sourceType, key, displayName: "Test Doctor" });
const shift = (sourceType, start, seniority = "HMO", key = "TEST DOCTOR") => ({
  id: `${sourceType}-${start}-${seniority}`,
  sourceType,
  doctorKey: key,
  start: `${start}T08:00:00+10:00`,
  end: `${start}T17:00:00+10:00`,
  title: "Clinical shift",
  rawValue: "AM",
  seniority,
});
const account = (claims) => ({ email: "doctor@example.com", role: "user", facilityOverviewEnabled: true, claims });

const movedSite = await resolveFacilityOverviewAccess(null, account([claim("ddh"), claim("mmc")]), {
  today: "2026-08-25",
  events: [shift("ddh", "2026-05-20"), shift("mmc", "2026-08-25")],
});
assert.equal(movedSite.mode, "site");
assert.equal(movedSite.facilityKey, "MMC", "current work must supersede an old-site claim at term changeover");
assert.equal(movedSite.workingToday, true);

const sms = await resolveFacilityOverviewAccess(null, account([claim("mmc")]), {
  today: "2026-08-25",
  events: [shift("mmc", "2026-08-25", "SMS")],
});
assert.equal(sms.mode, "all", "current-term SMS must retain the All EDs selector");
assert.equal(sms.isSms, true);
assert.equal(sms.workingToday, true);

const dayOff = await resolveFacilityOverviewAccess(null, account([claim("ddh")]), {
  today: "2026-08-25",
  events: [shift("ddh", "2026-08-27", "Junior Registrar")],
});
assert.equal(dayOff.mode, "site");
assert.equal(dayOff.facilityKey, "DDH");
assert.equal(dayOff.workingToday, false, "a day off must not trigger the On shift landing view");

const ambiguous = await resolveFacilityOverviewAccess(null, account([claim("ddh"), claim("mmc")]), {
  today: "2026-08-25",
  events: [shift("ddh", "2026-08-25"), shift("mmc", "2026-08-25")],
});
assert.equal(ambiguous.mode, "denied", "ambiguous non-SMS site evidence must fail closed");

const stateSource = await readFile(new URL("../functions/api/state.js", import.meta.url), "utf8");
const d1Source = await readFile(new URL("../functions/_lib/d1-calendar.js", import.meta.url), "utf8");
for (const action of ["Metadata", "ByStream", "OnShift", "Staff", "WorkingTogether"]) {
  const block = stateSource.match(new RegExp(`action === "queryFacilityOverview${action}"[\\s\\S]*?(?=\\n    if \\(action ===|\\n    const account =|$)`))?.[0] || "";
  assert.match(block, /facilityOverviewAccess\(\)/, `${action} must resolve server-side site access`);
  assert.match(block, /facilityOverviewAccessDeniedResponse|constrainFacilityOverviewSourceTypes/, `${action} must enforce or constrain the authenticated site`);
}
const contactResolutionBlock = stateSource.match(/action === "setContactAllocationResolution"[\s\S]*?(?=\n    if \(action ===|$)/)?.[0] || "";
assert.match(contactResolutionBlock, /facilityOverviewAccess\(\)[\s\S]*requestedFacility !== access\.facilityKey/, "temporary contact corrections must retain non-SMS site enforcement");
assert.match(contactResolutionBlock, /queryFacilityOverviewOnShift[\s\S]*facilityOverviewEventPeriod/, "a contact correction must target a rostered clinician in the same period");
assert.match(contactResolutionBlock, /attachContactAllocations[\s\S]*safe automatic match/, "safe automatic allocations must remain non-editable");
assert.match(stateSource, /queryContactAllocationResolutions[\s\S]*loadLiveContactListForOnShift/, "On shift should return temporary resolutions in its existing request");
assert.match(d1Source.match(/async function calendarSchemaIsCurrent[\s\S]*?async function ensureColumn/)?.[0] || "", /contact_allocation_resolutions[\s\S]*contact_allocation_resolution_history[\s\S]*has_contact_resolutions[\s\S]*has_contact_resolution_history/, "the schema fast path must not skip temporary contact-resolution tables");

const appSource = await readFile(new URL("../public/static/app.js", import.meta.url), "utf8");
assert.match(stateSource, /facilityOverviewAccess: prepared\.facilityOverviewAccess/, "fast login must carry access context in its existing response");
assert.match(appSource, /launchClinicalOnShiftWorkspace[\s\S]*workingToday[\s\S]*facilityOverviewState\.tab = "on-shift"/, "a working clinician must land on On shift without calendar hydration");
assert.match(appSource, /facilityOverviewIsSiteScoped\(\)[\s\S]*facility-overview-fixed-facility/, "non-SMS site scope must render as a fixed label");
assert.match(appSource, /data-facility-overview-contact-resolution[\s\S]*data-facility-overview-contact-resolution-target[\s\S]*setContactAllocationResolution/, "On shift review rows must offer an editable temporary-assignment selector");
assert.match(appSource, /is-manual[\s\S]*<sup aria-hidden="true">\*<\/sup>/, "a manually assigned number should use only a small adjacent asterisk");
assert.doesNotMatch(appSource.match(/function renderFacilityOverviewContactAllocation[\s\S]*?function renderFacilityOverviewContactListStatus/)?.[0] || "", />Manual</, "manual phone allocations should not display a Manual text label");
assert.match(appSource, /selectedContact && !selectedIsUnresolved[\s\S]*renderFacilityOverviewContactResolutionMenu/, "clicking a resolved manual number must still render its reassignment selector");
assert.match(appSource, /Reset assignment[\s\S]*<s>\$\{escapeHtml\(resetLabel\)\}<\/s>/, "resetting a manual association should strike through its current name and number");

console.log("Facility overview access tests passed.");
