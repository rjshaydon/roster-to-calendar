import assert from "node:assert/strict";
import { readFile, readdir } from "node:fs/promises";
import { DatabaseSync } from "node:sqlite";
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

class TracedD1 {
  constructor(sqlite) { this.sqlite = sqlite; this.sql = []; this.rowsWritten = 0; }
  prepare(sql) {
    const owner = this;
    owner.sql.push(sql.replace(/\s+/g, " ").trim());
    return {
      args: [],
      bind(...args) { this.args = args; return this; },
      async first() { return owner.sqlite.prepare(sql).get(...this.args) || null; },
      async all() { return { success: true, results: owner.sqlite.prepare(sql).all(...this.args) }; },
      async run() { const result = owner.sqlite.prepare(sql).run(...this.args); owner.rowsWritten += Number(result.changes || 0); return { success: true, meta: { changes: Number(result.changes || 0) } }; },
    };
  }
}

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

const sqlite = new DatabaseSync(":memory:");
for (const name of (await readdir(new URL("../migrations", import.meta.url))).filter((name) => name.endsWith(".sql")).sort()) {
  sqlite.exec(await readFile(new URL(`../migrations/${name}`, import.meta.url), "utf8"));
}
const db = new TracedD1(sqlite);
sqlite.prepare("INSERT INTO roster_files (id, name, source_type, active) VALUES ('access-file', 'Access.xlsx', 'mmc', 1)").run();
sqlite.prepare(`INSERT INTO facility_term_staff_contributions
  (source_type, term_start, doctor_key, file_id, display_name, seniority, first_applicable_date, last_applicable_date)
  VALUES ('mmc', '2026-08-03', 'TEST DOCTOR', 'access-file', 'Test Doctor', 'HMO', '2026-08-03', '2026-11-02')`).run();
sqlite.prepare("INSERT INTO roster_daily_presence (date, source_type, doctor_key, display_name, event_id) VALUES ('2026-08-25', 'mmc', 'TEST DOCTOR', 'Test Doctor', 'access-event')").run();
const materializedOptions = { today: "2026-08-25", now: new Date("2026-08-25T01:01:00Z"), materializedOnly: true };
const materializedFirst = await resolveFacilityOverviewAccess(db, account([claim("mmc")]), materializedOptions);
assert.equal(materializedFirst.mode, "site");
assert.equal(materializedFirst.facilityKey, "MMC");
assert.equal(materializedFirst.workingToday, true);
assert.equal(materializedFirst.cache, "refreshed");
assert.ok(new Date(materializedFirst.expiresAt).getTime() - materializedOptions.now.getTime() <= 15 * 60 * 1000, "access decisions must expire within 15 minutes");
assert.equal(db.sql.some((sql) => /\broster_events\b/i.test(sql)), false, "materialized access must never inspect roster_events");
db.sql = [];
db.rowsWritten = 0;
const materializedRepeat = await resolveFacilityOverviewAccess(db, account([claim("mmc")]), materializedOptions);
assert.equal(materializedRepeat.cache, "hit", "repeat access should reuse the compact decision");
assert.equal(db.rowsWritten, 0, "repeat access should write nothing");
assert.equal(db.sql.length, 1, "repeat access should require one compact indexed read");
assert.match(db.sql[0], /facility_access_sessions/);
const accessPlan = sqlite.prepare(`EXPLAIN QUERY PLAN
  SELECT access_json FROM facility_access_sessions
  WHERE subject_email = ? AND access_date = ? AND subject_revision = ? AND expires_at > ?`).all(
    "doctor@example.com", "2026-08-25", "revision", "2026-08-25T01:01:00Z",
  ).map((row) => String(row.detail || ""));
assert.ok(accessPlan.some((detail) => /INDEX.*subject_email/i.test(detail)), `access cache lookup must use its subject key: ${accessPlan.join("; ")}`);
const changedClaims = await resolveFacilityOverviewAccess(db, account([claim("ddh", "OTHER DOCTOR")]), materializedOptions);
assert.equal(changedClaims.preparing, true, "a claim revision without compact evidence must fail closed");
db.sql = [];
const repeatedMissing = await resolveFacilityOverviewAccess(db, account([claim("ddh", "OTHER DOCTOR")]), materializedOptions);
assert.equal(repeatedMissing.preparing, true);
assert.equal(repeatedMissing.cache, "hit", "missing compact evidence must use the bounded negative cache");
assert.equal(db.sql.length, 1, "a repeated missing decision must not probe every compact source table");
const nextDate = await resolveFacilityOverviewAccess(db, account([claim("mmc")]), { ...materializedOptions, today: "2026-08-26", now: new Date("2026-08-26T01:01:00Z") });
assert.equal(nextDate.cache, "refreshed", "a Melbourne roster-date change must not reuse yesterday's access decision");
assert.equal(nextDate.workingToday, false);
sqlite.prepare(`INSERT INTO facility_sms_memberships
  (source_type, doctor_key, display_name, first_seen_date, last_seen_date, created_at, updated_at)
  VALUES ('ddh', 'PERMANENT SMS', 'Permanent SMS', '2026-01-01', '2026-05-01', '2026-01-01T00:00:00Z', '2026-05-01T00:00:00Z')`).run();
const permanentSms = await resolveFacilityOverviewAccess(db, account([claim("ddh", "PERMANENT SMS")]), materializedOptions);
assert.equal(permanentSms.mode, "all", "continuing SMS membership must survive a term with no roster contribution");
db.sql = [];
const revoked = await resolveFacilityOverviewAccess(db, { ...account([claim("mmc")]), facilityOverviewEnabled: false }, materializedOptions);
assert.equal(revoked.mode, "denied", "revoked At a glance permission must override a cached decision immediately");
assert.equal(db.sql.length, 0, "explicit revocation must fail before access-cache reads");

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
assert.match(stateSource, /facilityOverviewSubject[\s\S]*targetEmail[\s\S]*facilityOverviewEnabled/, "Creator-entered accounts must be authorised using the entered user's At a glance entitlement");
assert.match(stateSource, /facilityOverviewSubject[\s\S]*materializedOnly: facilityAccessMaterialized/, "every At a glance handler must use the guarded compact access path when enabled");
assert.match(stateSource, /access\?\.preparing === true[\s\S]*status: 503[\s\S]*Retry-After/, "missing compact access evidence must return a bounded preparing response");
assert.match(stateSource.match(/if \(action === "adminLoadUser"\)[\s\S]*?if \(action === "loadAccountContext"\)/)?.[0] || "", /responseMode === "fast"[\s\S]*prepareFastLoginEnvelope[\s\S]*loadFastAccountSnapshotPayload/, "Creator account switching must not build a full account snapshot in its fast response");
assert.match(stateSource.match(/if \(action === "queryDoctorProfileFacilityOverviewAccess"\)[\s\S]*?if \(action === "saveDoctorProfile"\)/)?.[0] || "", /resolveDoctorAccount[\s\S]*facilityOverviewAccountEmail[\s\S]*facilityOverviewAccess/, "doctor profiles must resolve their linked user's At a glance scope separately from calendar loading");
assert.match(appSource, /launchClinicalOnShiftWorkspace[\s\S]*workingToday[\s\S]*facilityOverviewState\.tab = "on-shift"/, "a working clinician must land on On shift without calendar hydration");
assert.match(appSource.match(/function canUseFacilityOverview\(\)[\s\S]*?function sanitizeFacilityOverviewAccess/)?.[0] || "", /doctor-profile[\s\S]*currentFacilityOverviewEnabled[\s\S]*currentFacilityOverviewAccess\.mode !== "denied"/, "doctor-profile views must not inherit Creator At a glance access");
assert.match(appSource, /resetFacilityOverviewAccessForEnteredUser\(\)[\s\S]*enterUserAccount[\s\S]*resetFacilityOverviewAccessForEnteredUser\(\)[\s\S]*enterDoctorProfileView/, "account and profile switches must clear Creator access before the entered user's access is loaded");
assert.match(appSource, /targetEmail: facilityOverviewTargetEmail\(\)/, "At a glance queries must send the entered account as their access subject");
assert.match(appSource, /validateDoctorProfileCalendarInBackground[\s\S]*allowInlineBuild: false[\s\S]*loadDoctorProfileFacilityOverviewAccess/, "profile switching must avoid inline snapshot builds and load At a glance access independently");
assert.match(appSource, /facilityOverviewIsSiteScoped\(\)[\s\S]*facility-overview-fixed-facility/, "non-SMS site scope must render as a fixed label");
assert.match(appSource, /data-facility-overview-contact-resolution[\s\S]*data-facility-overview-contact-resolution-target[\s\S]*setContactAllocationResolution/, "On shift review rows must offer an editable temporary-assignment selector");
assert.match(appSource, /is-manual[\s\S]*<sup aria-hidden="true">\*<\/sup>/, "a manually assigned number should use only a small adjacent asterisk");
assert.doesNotMatch(appSource.match(/function renderFacilityOverviewContactAllocation[\s\S]*?function renderFacilityOverviewContactListStatus/)?.[0] || "", />Manual</, "manual phone allocations should not display a Manual text label");
assert.match(appSource, /selectedContact && !selectedIsUnresolved[\s\S]*renderFacilityOverviewContactResolutionMenu/, "clicking a resolved manual number must still render its reassignment selector");
assert.doesNotMatch(appSource, /Reset assignment/, "the cancellation row should not use a separate Reset assignment label");
assert.match(appSource, /const resetOption = existing[\s\S]*data-facility-overview-contact-resolution-clear[\s\S]*<strong><s>\$\{escapeHtml\(existing\.displayName[\s\S]*return `<div class="facility-overview-contact-resolution-menu"[\s\S]*\$\{resetOption\}\$\{candidates/, "the struck-through current person should be the first standard option in the selector");
assert.match(appSource, /contactReviewOpen: false[\s\S]*addEventListener\("toggle"[\s\S]*contactReviewOpen = review\.open === true/, "the review disclosure should preserve its state only while On shift remains active");
assert.match(appSource.match(/function closeFacilityOverview[\s\S]*?function renderFacilityOverview/)?.[0] || "", /collapseFacilityOverviewContactReview\(\)/, "returning to the calendar should collapse contact allocations needing review");
assert.match(appSource.match(/async function loadFacilityOverviewOnShift[\s\S]*?function renderFacilityOverviewOnShiftResults/)?.[0] || "", /collapseFacilityOverviewContactReview\(\)/, "entering or reloading On shift should begin with contact review collapsed");
assert.match(appSource, /window\.addEventListener\("pagehide", \(\) => collapseFacilityOverviewContactReview\(\)\)/, "closing or navigating away from the app should reset contact review disclosure state");
assert.match(appSource, /serviceContacts:[\s\S]*ED Care-Co[\s\S]*GAP \/ Geriatric AH[\s\S]*Geriatrician[\s\S]*CART clinician/, "the requested service contacts should be carried into their period renderers");
assert.match(appSource, /renderFacilityOverviewGroupedServiceCard\("ED Care-Co \/ GAP"[\s\S]*renderFacilityOverviewGroupedServiceCard\("Geriatrician \/ CART"/, "each site's two explicit services should share one grouped stream card");
assert.match(appSource.match(/function renderFacilityOverviewGroupedServiceCard[\s\S]*?function renderFacilityOverviewUnstreamedCard/)?.[0] || "", /group\.label[\s\S]*facility-overview-contact-number[\s\S]*hideContactAllocation: true/, "a grouped service row should place its phone beside the service label and any name on the following line");
assert.match(appSource.match(/async function loadCloudCalendarEvents[\s\S]*?function cloudCalendarEventRange/)?.[0] || "", /response\.status === 503[\s\S]*allowInlineBuild: false/, "a resource-limit calendar response should retry once without synchronous rebuilding");
const workingEventSource = stateSource.match(/function isFacilityOverviewWorkingEvent[\s\S]*?function facilityOverviewEventPeriod/)?.[0] || "";
assert.match(workingEventSource, /includeClinicalSupport !== true && isClinicalSupportRosterEvent\(event\)/, "all recognised CS variants must be hidden by the server when Include CS is off");
assert.doesNotMatch(workingEventSource, /facilityKey[^\n]+VHH[^\n]+allDay/, "all-day VHH Clinical Support must not be discarded before the Include CS rule is applied");
assert.match(appSource.match(/function facilityOverviewIsClinicalSupportAssignment[\s\S]*?function facilityOverviewIsOnsiteClinicalSupportAssignment/)?.[0] || "", /clinicalSupportRosterMode\(assignment\)/, "the browser must use the shared Clinical Support classifier");
const onShiftPeriodSource = appSource.match(/function renderFacilityOverviewOnShiftPeriod[\s\S]*?function facilityOverviewIsDdhPeriod/)?.[0] || "";
assert.match(onShiftPeriodSource, /clinicalSupportCard[\s\S]*return `\$\{clinicalSupportCard\}\$\{renderFacilityOverview/, "Clinical Support must render before ordinary stream cards for every hospital");
assert.match(appSource, /clinicalSupportMode === "onsite"[\s\S]*\(On-site\)[\s\S]*clinicalSupportMode === "office"[\s\S]*\(Office\)/, "On Shift must label on-site and office Clinical Support clinicians explicitly");
assert.match(stateSource, /function sanitizeDetectedSources[\s\S]*vhh: Array\.isArray\(input\.vhh\)[\s\S]*function detectedSourcesForSnapshot[\s\S]*sourceType === "vhh"/, "server snapshots must preserve VHH as a detected source");
assert.match(appSource.match(/function sanitizeWorkspaceSnapshot[\s\S]*?function sanitizeInsightCache/)?.[0] || "", /vhh: Array\.isArray\(value\.detectedSources\?\.vhh\)/, "browser snapshot caching must preserve VHH as a detected source");

console.log("Facility overview access tests passed.");
