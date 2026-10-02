import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import vm from 'node:vm';
import { attachContactAllocations, contactRosterAssignments, partitionDdhNightReview, ddhNightReviewWindow } from '../public/static/contact-allocations.js';
import { loadPublishedPreviousDdhNight } from '../functions/_lib/facility-overview-cache.js';
import { loadPublishedFacilityContacts, publishFacilityContactExtract } from '../functions/_lib/facility-contact-cache.js';
import { issueFacilityContactAccessToken, verifyFacilityContactAccessToken } from '../functions/_lib/facility-contact-access.js';

const at = local => new Date(`${local}+10:00`);
const date = '2026-10-02', previousDate = '2026-10-01';
const row = (name, day = date, sourceType = 'ddh') => ({ sourceType, doctorKey: name.toUpperCase(), displayName: name, seniority: 'HMO', event: { title: 'DDH: Orange Night', start: `${day}T23:00:00+10:00`, end: `${day === date ? '2026-10-03' : date}T09:00:00+10:00`, seniority: 'HMO' } });
const contact = (name, phone = '49901') => ({ area: 'Dandenong Emergency', shift: 'Night', role: 'Orange Dr', name, phone, isPopulated: true, contactKey: name });
const currentRows = [row('Tonight SMITH'), row('Vanessa NEW')];
const previousRows = [row('Tony OLD', previousDate), row('Vanessa OLD', previousDate)];
const contacts = [contact('Tony'), contact('Vanessa', '49902'), contact('Unknown', '49903')];
const assignments = contactRosterAssignments(currentRows);
const context = { available: true, nightDate: date, previousNightDate: previousDate, rows: previousRows };
let now = at(`${date}T23:01:00`);
const matches = attachContactAllocations(assignments, contacts, [], { now });
let partition = partitionDdhNightReview(matches, assignments, { date, now, previousNightRoster: context });
assert.deepEqual(partition.previousNight.map(c => c.name), ['Tony']);
assert.deepEqual(partition.unresolved.map(c => c.name), ['Unknown']);
assert.equal(matches.assignments.find(a => a.person.displayName === 'Vanessa NEW').contactAllocation.sourceName, 'Vanessa', 'names working both nights remain on tonight');
assert.equal(partitionDdhNightReview(matches, assignments, { date, now }).unresolved.length, 2, 'missing history must preserve ordinary review');
assert.equal(partitionDdhNightReview(matches, assignments, { date, now, previousNightRoster: { ...context, nightDate: previousDate } }).previousNight.length, 0);
const ambiguousCurrent = contactRosterAssignments([row('Tony SMITH'), row('Tony JONES')]);
const ambiguousMatches = attachContactAllocations(ambiguousCurrent, [contacts[0]], [], { now });
assert.equal(partitionDdhNightReview(ambiguousMatches, ambiguousCurrent, { date, now, previousNightRoster: context }).previousNight.length, 0, 'plausible current names cannot be hidden as leftovers');
assert.equal(partitionDdhNightReview(matches, assignments, { date, now, previousNightRoster: { ...context, rows: [row('Tony OLD', previousDate), row('Tony OTHER', previousDate)] } }).previousNight.length, 0, 'prior match must also be unique');
for (const override of [{ rawValue: 'Night sick leave' }, { status: 'unknown' }, { allDay: true }, { title: 'DDH: HITH Night' }, { end: '2026-10-01T09:00:00+10:00' }]) {
 const invalid = row('Tony OLD', previousDate); invalid.event = { ...invalid.event, ...override };
 assert.equal(partitionDdhNightReview(matches, assignments, { date, now, previousNightRoster: { ...context, rows: [invalid] } }).previousNight.length, 0, 'non-working and invalid prior events cannot hide a contact');
}
for (const local of [`${date}T23:00:00`, '2026-10-03T00:01:00', '2026-10-03T07:29:00']) {
 assert.equal(ddhNightReviewWindow(date, at(local)).nightDate, date);
 assert.equal(partitionDdhNightReview(matches, assignments, { date, now: at(local), previousNightRoster: context }).previousNight.length, 1);
}
assert.equal(ddhNightReviewWindow('2026-10-03', at('2026-10-03T08:59:00')).hideNight, false, 'outgoing shift persists until 09:00');
for (const local of [`${date}T09:00:00`, `${date}T22:59:00`]) {
 assert.equal(ddhNightReviewWindow(date, at(local)).hideNight, true);
 assert.equal(partitionDdhNightReview(matches, assignments, { date, now: at(local), previousNightRoster: context }).unresolved.length, 0);
}
// Melbourne daylight saving starts on 4 October; midnight still belongs to 3 October.
assert.equal(ddhNightReviewWindow('2026-10-03', new Date('2026-10-04T04:00:00+11:00')).nightDate, '2026-10-03');

class MemoryR2 {
 constructor() { this.objects = new Map(); this.gets = 0; this.puts = 0; }
 async get(key) { this.gets++; const value = this.objects.get(key); if (value === undefined) return null; const bytes = new TextEncoder().encode(value); return { text: async () => value, arrayBuffer: async () => bytes.buffer, etag: 'fixture' }; }
 async put(key, value) { this.puts++; this.objects.set(key, typeof value === 'string' ? value : new TextDecoder().decode(value)); return { etag: 'fixture' }; }
}
const r2 = new MemoryR2();
r2.objects.set('facility-overview/v1/ddh/manifest.json', JSON.stringify({ days: { [previousDate]: { key: 'previous-day' } } }));
r2.objects.set('previous-day', JSON.stringify({ date: previousDate, rows: [...previousRows, row('Wrong site', previousDate, 'mmc'), row('Wrong day')] }));
const loaded = await loadPublishedPreviousDdhNight(r2, date, now);
assert.equal(r2.gets, 2); assert.equal(r2.puts, 0); assert.equal(loaded.rows.length, 2);
assert.equal(await loadPublishedPreviousDdhNight(r2, date, at(`${date}T22:59:00`)), null);
assert.equal(r2.gets, 2, 'no additional roster reads before the night window');
await publishFacilityContactExtract(r2, { sourceId: 'ddh-daily-contact-sheet', sourceDate: date, contacts: contacts.map(c => ({ ...c })) });
const beforeHandover = await loadPublishedFacilityContacts(r2, { date, facilityKeys: ['DDH'], now: at(`${date}T22:59:00`) });
const afterHandover = await loadPublishedFacilityContacts(r2, { date, facilityKeys: ['DDH'], now });
assert.notEqual(beforeHandover.revision, afterHandover.revision, '23:00 must update contact visibility even if the sheet is unchanged');
assert.equal(beforeHandover.contacts.length, 0); assert.equal(afterHandover.contacts.length, 3);

// Actual token refresh path: scoped authentication, zero D1, one previous-roster
// fetch at the night transition, then existing three-R2-read refreshes only.
const stateSource = await readFile(new URL('../functions/api/state.js', import.meta.url), 'utf8');
const start = stateSource.indexOf('async function refreshPublishedFacilityContactsWithToken(');
const end = stateSource.indexOf('\nasync function ', start + 1);
const secret = 'ddh-night-test-secret-with-at-least-32-chars';
const token = await issueFacilityContactAccessToken(secret, { facilityKey: 'DDH', now: now.getTime() });
const route = vm.createContext({ Response, verifyFacilityContactAccessToken: (s,t) => verifyFacilityContactAccessToken(s,t,{ now: now.getTime() }), sanitizeSourceTypes: values => values.map(v => v.toLowerCase()), facilityContactReaderSources: () => ['ddh'], loadPublishedFacilityContacts: (store,options) => loadPublishedFacilityContacts(store,{ ...options,now }), loadPublishedPreviousDdhNight: (store,d) => loadPublishedPreviousDdhNight(store,d,now) });
vm.runInContext(stateSource.slice(start,end), route);
const env = { FACILITY_SHARED_CONTACTS_ENABLED: 'true', FACILITY_CONTACT_ACCESS_SECRET: secret, ROSTER_FILES: r2, ROSTER_DB: new Proxy({}, { get() { throw new Error('Unexpected D1 access'); } }) };
const body = { date, facilityKey: 'DDH', contactAccessToken: token, contactRevision: beforeHandover.revision };
const first = await route.refreshPublishedFacilityContactsWithToken({ env }, body);
const firstData = await first.json(); assert.equal(first.status, 200); assert.equal(firstData.previousNightRoster.available, true); assert.equal(firstData.contactList.contacts.length, 3);
const readBefore = r2.gets, writeBefore = r2.puts;
for (let i=0;i<100;i++) {
 const response = await route.refreshPublishedFacilityContactsWithToken({ env }, { ...body, contactRevision: afterHandover.revision, previousNightRosterDate: date });
 const data = await response.json(); assert.equal(data.unchanged, true); assert.equal(data.previousNightRoster, undefined);
}
assert.equal(r2.gets-readBefore, 300); assert.equal(r2.puts, writeBefore);
const denied = await route.refreshPublishedFacilityContactsWithToken({ env }, { ...body, facilityKey: 'MMC' });
assert.equal(denied.status,403);

// Actual review renderer: Creator-only nested list, collapsed by default;
// ordinary users see only genuinely unresolved entries, with no old-name editor.
const app = await readFile(new URL('../public/static/app.js', import.meta.url), 'utf8');
let creator = true;
const state = { date, previousNightRoster: context, contactList: { status: 'available', sourceDate: date, contacts }, contactResolutionMenu: 'Tony' };
const ui = vm.createContext({ facilityOverviewState: state, isViewingCreatorAccount: () => creator, escapeHtml: value => String(value || '').replaceAll('&','&amp;').replaceAll('<','&lt;'), formatFacilityOverviewContactTime: () => '', ddhNightReviewWindow: d => ddhNightReviewWindow(d,now), partitionDdhNightReview: (m,a,options) => partitionDdhNightReview(m,a,{ ...options,now }), renderFacilityOverviewContactReviewRow: c => `<span>${c.name}</span>`, renderFacilityOverviewContactResolutionMenu: () => '<p>Editor</p>' });
const uiStart = app.indexOf('function renderFacilityOverviewContactListStatus(');
const uiEnd = app.indexOf('\nfunction ',uiStart+1);
vm.runInContext(app.slice(uiStart,uiEnd),ui);
const reviewStart = app.indexOf('function renderFacilityOverviewContactReviewRow(');
vm.runInContext(app.slice(reviewStart,app.indexOf('\nfunction ',reviewStart+1)),ui);
const creatorHtml = ui.renderFacilityOverviewContactListStatus(matches,assignments);
assert.match(creatorHtml,/data-facility-overview-previous-night-review/);
assert.match(creatorHtml,/Tony/); assert.match(creatorHtml,/Unknown/);
assert.doesNotMatch(creatorHtml,/data-facility-overview-previous-night-review open/);
state.contactPreviousNightReviewOpen = true;
assert.match(ui.renderFacilityOverviewContactListStatus(matches,assignments),/data-facility-overview-previous-night-review open/);
creator = false;
const ordinaryHtml = ui.renderFacilityOverviewContactListStatus(matches,assignments);
assert.doesNotMatch(ordinaryHtml,/Tony|previous-night-review|Editor/); assert.match(ordinaryHtml,/Unknown/);
console.log('DDH night review passed: 23:00–09:00 and midnight/DST boundaries; unique prior matches; current ambiguity retained; Creator-only collapsed detail; ordinary-user hiding; handover revision; zero D1/writes and no extra roster reads during repeated refresh.');
