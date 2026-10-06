import { facilityAccessKeys, facilityAccessAllows } from '../public/static/facility-access-policy.js';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import vm from 'node:vm';
import { DatabaseSync } from 'node:sqlite';
import { attachContactAllocations, contactAllocationCandidates, contactRosterAssignments, mergeContactResolutionRefresh, assignmentMatchesContactContext, validateContactResolutionSelection, normaliseContactListExtract } from '../public/static/contact-allocations.js';
import { queryContactAllocationResolutions, saveContactAllocationResolution } from '../functions/_lib/d1-calendar.js';
import { loadPublishedFacilityContacts, publishFacilityContactExtract as publishContactExtract, publishFacilityContactResolutions } from '../functions/_lib/facility-contact-cache.js';

const now = new Date('2026-10-02T01:00:00Z'); // 11:00 Melbourne, before daylight saving.
// Publication expiry uses the wall clock too; keep historical fixtures live
// regardless of the date on which this regression suite is run.
async function publishFacilityContactExtract(...args) {
  const RealDate = globalThis.Date;
  globalThis.Date = class extends RealDate { constructor(...values) { super(...(values.length ? values : [now.getTime()])); } };
  try { return await publishContactExtract(...args); }
  finally { globalThis.Date = RealDate; }
}

const area = { MMC: 'Adult Emergency', MCH: 'Paediatric Emergency', DDH: 'Dandenong Emergency', VHH: 'Victorian Heart Hospital Emergency' };
const staff = (name, source = 'DDH', team = 'Orange', seniority = 'HMO', period = 'AM', start = '08:00', end = '17:30') => ({
  source, period, team, person: { doctorKey: name.toUpperCase(), displayName: name, sourceType: source.toLowerCase(), seniority },
  event: { source, title: `${source}: ${team} ${period}`, seniority, start: `2026-10-02T${start}:00`, end: `2026-10-02T${end}:00` },
});
const sheet = (name, source = 'DDH', phone = '49900', role = 'Orange Dr 1', shift = 'AM') => ({
  area: area[source], shift: source === 'VHH' ? 'Current' : shift, role, name, phone, isPopulated: Boolean(name), sourceDate: '2026-10-02', contactKey: `${source}|${shift}|${role}|${name}|${phone}`,
});
const match = (roster, contacts, resolutions = []) => attachContactAllocations(roster, contacts, resolutions, { now });
const matchedNames = result => result.assignments.filter(a => a.contactAllocation).map(a => a.person.displayName);

// Held-out synthetic cases verify mechanics, not clinical accuracy calibration.
const positiveCases = [
  ['Thisun (Tea)', 'Tea GUNASEN', 'explicit-alternate-name'],
  ['Thisun “Tea”', 'Tea GUNASEN', 'explicit-alternate-name'],
  ['Thisun (Tea) Gunasen', 'Tea GUNASEN', 'explicit-alternate-name'],
  ['Tea Gunasen (Thisun)', 'Thisun GUNASEN', 'explicit-alternate-name'],
  ['Jacquline Morel', 'Jacqueline MOREL', 'spelling-or-alternate-full-name'],
  ['Jacquline', 'Jacqueline MOREL', 'spelling'],
  ['Jacqeuline', 'Jacqueline MOREL', 'spelling'],
  ['Katherine Smthson', 'Katherine SMITHSON', 'spelling-or-alternate-full-name'],
  ['Ann Soo', 'Yee Ann SOO', 'spelling-or-alternate-full-name'],
  ['Qing C', 'Qingyang CHEN', 'surname-initial'],
  ['J Smithson', 'Julia SMITHSON', 'spelling-or-alternate-full-name'],
  ['José Núñez', 'Jose NUNEZ', 'exact'],
  ['Anne-Marie O’Connor', 'Anne Marie OCONNOR', 'exact'],
  ['王 小明', '小明 王', 'reordered-name'],
  ['Craig Jirayut', 'Craig', 'roster-given-name-only'],
  ['Craig Jirayut', 'Craig PROMPEN', 'approved-identity-alias'],
  ['Ollie', 'Oliver DEANS', 'alias'],
  ['Meg', 'Megha PHILIP', 'alias'],
  ['Ben', 'Benjamin BRENNAN DOYLE', 'alias'],
  ['Rosie', 'Rosemary SASSE', 'alias'],
  ['Youshna', 'Youstina NAN', 'spelling'],
];
for (const [contactName, rosterName, method] of positiveCases) {
  const result = match([staff(rosterName)], [sheet(contactName)]);
  assert.equal(result.matchedCount, 1, `${contactName} should match ${rosterName}`);
  assert.equal(result.assignments[0].contactAllocation.matchMethod, method);
  assert.equal(result.assignments[0].contactAllocation.uncertain, !['exact', 'reordered-name'].includes(method));
}
for (const source of ['MMC', 'MCH', 'DDH', 'VHH']) {
  const result = match([staff('Jacqueline MOREL', source)], [sheet('Jacquline', source)]);
  assert.equal(result.matchedCount, 1, `same matcher should work at ${source}`);
  assert.equal(result.assignments[0].contactAllocation.uncertain, true);
}
for (const [contactName, rosterName] of [['Alex SMITH', 'Alex JONES'], ['Ama', 'Arnav MEHTA'], ['Teo', 'Tea GUNASEN'], ['Jhon', 'John SMITH'], ['Unknown', 'Tea GUNASEN'], ['Thisun (Tea) OTHER', 'Tea GUNASEN'], ['Thisun OTHER (Tea)', 'Tea GUNASEN']]) {
  assert.equal(match([staff(rosterName)], [sheet(contactName)]).matchedCount, 0, `must not guess ${contactName} as ${rosterName}`);
}
assert.equal(match([staff('Tea GUNASEN'), staff('Tea OTHER')], [sheet('Thisun (Tea)')]).matchedCount, 0);
assert.equal(match([staff('Pat FINN'), staff('Patrick OTHER')], [sheet('Pat')]).matchedCount, 0);
assert.equal(match([staff('Craig'), { ...staff('Craig'), person: { ...staff('Craig').person, doctorKey: 'SECOND CRAIG' } }], [sheet('Craig Jirayut')]).matchedCount, 0, 'a fuller sheet name cannot resolve two given-name-only roster identities');
assert.equal(match([staff('Craig'), staff('Craig JIRAYUT')], [sheet('Craig Jirayut')]).matchedCount, 0, 'do not prefer a full-name candidate without enough separation from an incomplete roster identity');
assert.equal(match([staff('Ollie JONES'), staff('Oliver DEANS')], [sheet('Ollie')]).matchedCount, 0, 'nickname expansion must not override an equally plausible given name');
assert.equal(match([staff('Megha PHILIP'), staff('Megan JONES')], [sheet('Meg')]).matchedCount, 0, 'Meg must remain ambiguous when another plausible expansion is rostered');
assert.equal(match([staff('Benjamin DOYLE'), staff('Benedict JONES')], [sheet('Ben')]).matchedCount, 0);
assert.equal(match([staff('Rosemary SASSE'), staff('Rose JONES')], [sheet('Rosie')]).matchedCount, 0);
assert.equal(match([staff('Youstina NAN'), staff('Youshena JONES')], [sheet('Youshna')]).matchedCount, 0, 'two plausible anchored spellings remain ambiguous');
for (const [contact, name] of [['Youshna JONES', 'Youstina NAN'], ['Youstne', 'Youstina NAN'], ['Yostna', 'Youstina NAN'], ['Youstina Youshna', 'Youstina YOUSTINA']]) {
  assert.equal(match([staff(name)], [sheet(contact)]).matchedCount, 0, 'shorter/poorly anchored spellings and conflicting surnames must not use the wider given-name rule');
}
const headerRows = ['AM', 'PM', 'Night'].map(period => sheet('NAME', 'MCH', period === 'Night' ? '' : 'PHONE', 'ROLE', period));
assert.equal(match([], headerRows).unmatched.length, 0, 'worksheet headers are not contact allocations');
const normalizedHeaders = normaliseContactListExtract({ sourceId: 'mmc-shift-allocations', sourceDate: '2026-10-02', contacts: headerRows });
assert.equal(normalizedHeaders.contacts.length, 0);
assert.equal(match([], [sheet('NAME', 'MCH', '25150', 'Paeds Dr', 'PM')]).unmatched.length, 1, 'header exclusion must require the complete header pattern');
assert.equal(match([staff('Youstina NAN', 'MCH', 'Paeds', 'HMO', 'PM')], [sheet('Ungell', 'MCH', '25176', 'Paeds Dr', 'PM')]).matchedCount, 0);
const arnavRows = [sheet('Arnav - 25192', 'MMC', '25192', 'SEPSIS DR - MUST CARRY 25141', 'PM'), sheet('Arnav', 'MMC', '25192', 'Dr', 'PM')];
const arnavStaff = [staff('Arnav MEHTA', 'MMC', 'Sepsis', 'HMO', 'PM')];
const arnavResult = match(arnavStaff, arnavRows);
assert.equal(arnavResult.matchedCount, 1, 'identical handset/name repetitions are one allocation, not a conflict');
assert.equal(arnavResult.unmatched.length, 0);
assert.equal(arnavResult.assignments[0].contactAllocation.phone, '25192', 'do not replace the actual phone with the role instructions');
assert.equal(arnavResult.assignments[0].contactAllocation.uncertain, true);
assert.deepEqual(fingerprintForRows(arnavResult), fingerprintForRows(match(arnavStaff, [...arnavRows].reverse())));
function fingerprintForRows(result) { return result.assignments.map(a => [a.person.doctorKey, a.contactAllocation?.phone, a.contactAllocation?.contactKey]); }
const duplicateReject = { contactKey: arnavRows[1].contactKey, active: false, decision: 'rejected', revision: 1 };
assert.equal(match(arnavStaff, arnavRows, [duplicateReject]).matchedCount, 0, 'rejecting either duplicate must not resurrect through the other row');
const duplicateConfirmation = { ...duplicateReject, active: true, decision: 'assigned', doctorKey: 'ARNAV MEHTA', sourceType: 'mmc' };
assert.equal(match(arnavStaff, arnavRows, [duplicateConfirmation]).assignments[0].contactAllocation.matchMethod, 'manual');
assert.equal(match(arnavStaff, arnavRows, [duplicateConfirmation]).unmatched.length, 0);
assert.equal(validateContactResolutionSelection(arnavStaff, arnavRows, [], { contact: arnavRows[0], doctorKey: 'ARNAV MEHTA', now }).error, undefined, 'repeated rows remain editable');
const incompatibleConfirmation = { ...duplicateConfirmation, contactKey: arnavRows[0].contactKey, doctorKey: 'OTHER CLINICIAN' };
assert.equal(match(arnavStaff, arnavRows, [duplicateConfirmation, incompatibleConfirmation]).matchedCount, 0, 'coalescing cannot conceal conflicting human choices');
assert.equal(match(arnavStaff, [arnavRows[0], { ...arnavRows[1], name: 'Alex', contactKey: 'different-name' }]).matchedCount, 0, 'different names sharing a phone remain a conflict');
assert.equal(match(arnavStaff, [arnavRows[1], sheet('Arnav', 'MMC', '25193', 'Dr 2', 'PM')]).matchedCount, 0, 'same names with different handsets remain a conflict');
assert.equal(match(arnavStaff, [sheet('Arnav - 25193', 'MMC', '25192', 'Dr', 'PM')]).matchedCount, 0, 'strip an embedded phone only when it agrees with the actual handset');
const blankRows = Array.from({ length: 9 }, (_, index) => ({ ...sheet('*', 'MMC', '*', 'Dr', 'Night'), contactKey: `blank-${index}` }));
assert.equal(match([], blankRows).unmatched.length, 0, 'punctuation-only placeholder rows are empty allocations');
const normalizedBlanks = normaliseContactListExtract({ sourceId: 'mmc-shift-allocations', sourceDate: '2026-10-02', contacts: blankRows });
assert(normalizedBlanks.contacts.every(c => !c.isPopulated));
assert.equal(match([], normalizedBlanks.contacts).unmatched.length, 0);
assert.equal(match([], [sheet('Tara K', 'MMC', 'Call Switch - 92', 'ADULT SMS ON CALL', 'Night')]).unmatched.length, 1, 'keep the requested on-call detail visible');
assert.equal(match([staff('Craig PROMPEN', 'MMC')], [sheet('Craig Jirayut', 'MMC')]).matchedCount, 0, 'confirmed social names apply only to their approved site and identity');
assert.equal(match([staff('Craig JONES')], [sheet('Craig Jirayut')]).matchedCount, 0, 'the approved alias cannot bypass surname contradictions for another Craig');
assert.equal(match([staff('Craig', 'DDH', 'Silver', 'HMO', 'PM')], [sheet('Craig Jirayut', 'DDH', '0478068178', 'Orange Dr 8', 'PM')]).matchedCount, 1, 'the reported DDH PM stream disagreement must not prevent the given-name-only match');
assert.equal(match([staff('Oliver DEANS', 'DDH', 'Silver', 'HMO', 'PM')], [sheet('Ollie', 'DDH', '49916', 'Silver Dr 2', 'PM')]).matchedCount, 1);
assert.equal(match([staff('Jacqueline MOREL'), staff('Jacquline OTHER')], [sheet('Jacquline')]).matchedCount, 0, 'a plausible competitor needs a clear margin');
assert.equal(match([staff('Katherine SMITHSON'), staff('Katharine JONES')], [sheet('Kathrine')]).matchedCount, 0, 'two plausible spelling candidates must remain for review');
assert.equal(match([staff('Tea GUNASEN', 'DDH')], [sheet('Thisun (Tea)', 'MMC')]).matchedCount, 0);
assert.equal(match([staff('Tea GUNASEN', 'DDH', 'Orange', 'HMO', 'PM')], [sheet('Thisun (Tea)')]).matchedCount, 0);
assert.equal(match([staff('Tea GUNASEN', 'DDH', 'Silver')], [sheet('Thisun (Tea)')]).matchedCount, 1, 'changed stream must not defeat strong name evidence');
const streamSame = contactAllocationCandidates([staff('Jacqueline MOREL')], sheet('Jacquline'), { now })[0];
const streamDifferent = contactAllocationCandidates([staff('Jacqueline MOREL', 'DDH', 'Silver')], sheet('Jacquline'), { now })[0];
assert.equal(streamSame.score - streamDifferent.score, 3);
assert.equal(match([staff('Alex JONES')], [sheet('Alex SMITH', 'DDH', '49900', 'Orange HMO')]).matchedCount, 0, 'stream and grade cannot rescue a contradictory name');
assert.equal(match([staff('Tea GUNASEN')], [sheet('Thisun (HMO)')]).matchedCount, 0, 'annotations are not alternate names');
assert.equal(match([staff('Tea GUNASEN')], [sheet('Thisun (10:00)')]).matchedCount, 0);

// Decisions must be invariant under contact/event order and event duplicates.
const roster = [staff('Tea GUNASEN'), staff('Jacqueline MOREL'), staff('Katherine SMITHSON')];
const contacts = [sheet('Thisun (Tea)', 'DDH', '49900'), sheet('Jacquline', 'DDH', '49901'), sheet('Kathrine', 'DDH', '49902')];
const fingerprint = result => [...new Set(result.assignments.filter(a => a.contactAllocation).map(a => `${a.person.doctorKey}:${a.contactAllocation.phone}:${a.contactAllocation.matchMethod}`))].sort();
const expected = fingerprint(match(roster, contacts));
for (const rosterOrder of [roster, [...roster].reverse(), [roster[1], roster[2], roster[0]], [...roster, { ...roster[0] }]]) {
  for (const contactOrder of [contacts, [...contacts].reverse(), [contacts[1], contacts[2], contacts[0]]]) assert.deepEqual(fingerprint(match(rosterOrder, contactOrder)), expected);
}
assert.equal(match([roster[0], { ...roster[0] }], [contacts[0]]).matchedCount, 1, 'duplicate events count as one holder');
const competing = [sheet('Thisun (Tea)', 'DDH', '49900'), sheet('Tea (Thisun)', 'DDH', '49901')];
assert.equal(match([roster[0]], competing).matchedCount, 0, 'two tentative contacts must not claim one clinician');
const cascade = [sheet('Jacquline', 'DDH', '49900'), sheet('Jacqueline (Jane)', 'DDH', '49901')];
assert.equal(match([staff('Jacqueline MOREL'), staff('Jacqueline OTHER')], cascade).matchedCount, 0, 'guess consumption must not turn ambiguity into certainty');
assert.equal(match([staff('Tea GUNASEN'), staff('Jacqueline MOREL')], [sheet('Tea', 'DDH', '49900'), sheet('Jacqueline', 'DDH', '49900')]).matchedCount, 0, 'same handset cannot have two holders in a shift');

// Server roster conversion and browser display assignments must score identically,
// even if a display/parser grouping differs from original event evidence.
const rawRows = roster.map(a => ({ sourceType: a.person.sourceType, doctorKey: a.person.doctorKey, displayName: a.person.displayName, seniority: a.person.seniority, event: a.event }));
const server = contactRosterAssignments(rawRows);
const browser = roster.map(a => ({ ...a, team: 'Display override', period: 'PM', suggestedTitle: 'Display override' }));
assert.deepEqual(fingerprint(match(server, contacts)), fingerprint(match(browser, contacts)));
for (const contact of contacts) {
  assert.deepEqual(contactAllocationCandidates(server, contact, { now }).map(c => [c.identity, c.score, c.method]), contactAllocationCandidates(browser, contact, { now }).map(c => [c.identity, c.score, c.method]));
}
const suggestion = contacts[0];
const key = suggestion.contactKey;
const rejected = { id: 'r1', contactKey: key, decision: 'rejected', active: false, revision: 1 };
assert.equal(match(roster, [suggestion], [rejected]).matchedCount, 0);
assert.match(match(roster, [suggestion], [rejected]).unmatched[0].reviewReason, /rejected/);
assert.equal(match(roster, [suggestion], [{ ...rejected, decision: 'cleared', revision: 2 }]).matchedCount, 1, 'clearing a rejection explicitly allows automatic matching again');
const assigned = { ...rejected, decision: 'assigned', active: true, doctorKey: roster[0].person.doctorKey, sourceType: 'ddh', revision: 3 };
const confirmed = match(roster, [suggestion], [assigned]);
assert.equal(confirmed.assignments[0].contactAllocation.matchMethod, 'manual');
assert.equal(confirmed.assignments[0].contactAllocation.uncertain, false);
const reassigned = { ...assigned, doctorKey: roster[1].person.doctorKey };
assert.deepEqual(matchedNames(match(roster, [suggestion], [reassigned])), ['Jacqueline MOREL'], 'confirmation takes precedence over fuzzy matching');
assert.equal(validateContactResolutionSelection(roster, [suggestion], [], { contact: suggestion, doctorKey: roster[0].person.doctorKey, now }).target.person.doctorKey, roster[0].person.doctorKey);
assert.equal(validateContactResolutionSelection(roster, [suggestion], [], { contact: suggestion, decision: 'rejected', now }).error, undefined);
const exactContact = sheet('Tea GUNASEN');
assert.equal(validateContactResolutionSelection(roster, [exactContact], [], { contact: exactContact, decision: 'rejected', now }).status, 409, 'strong automatic matches remain protected');
assert.equal(validateContactResolutionSelection(roster, [suggestion], [], { contact: suggestion, doctorKey: 'ABSENT', now }).status, 400);
assert.equal(validateContactResolutionSelection(roster, [suggestion], [], { contact: suggestion, doctorKey: roster[0].person.doctorKey, decision: 'rejected', now }).status, 400);
const duplicatePhone = { ...suggestion, contactKey: 'duplicate-phone', name: 'Jacqueline' };
assert.equal(validateContactResolutionSelection(roster, [suggestion, duplicatePhone], [], { contact: suggestion, doctorKey: roster[0].person.doctorKey, now }).status, 409, 'a successful correction must not silently remain unusable because of duplicate sheet entries');

const vhhStaff = staff('Tea GUNASEN', 'VHH');
const vhhContact = sheet('Thisun (Tea)', 'VHH', '12018', 'SSU Dr');
assert.equal(match([vhhStaff], [vhhContact]).matchedCount, 1);
assert.equal(validateContactResolutionSelection([vhhStaff], [vhhContact], [], { contact: vhhContact, doctorKey: vhhStaff.person.doctorKey, now }).error, undefined);
const vhhResolution = { ...assigned, contactKey: vhhContact.contactKey, sourceType: 'vhh' };
assert.equal(match([vhhStaff], [vhhContact], [vhhResolution]).matchedCount, 1);
const finish = new Date('2026-10-02T07:30:00Z');
assert.equal(attachContactAllocations([vhhStaff], [vhhContact], [vhhResolution], { now: finish }).matchedCount, 0, 'confirmation cannot extend beyond roster finish');
assert.equal(validateContactResolutionSelection([vhhStaff], [vhhContact], [], { contact: vhhContact, doctorKey: vhhStaff.person.doctorKey, now: finish }).status, 400);
assert.equal(match([{ ...vhhStaff, event: { ...vhhStaff.event, allDay: true } }], [vhhContact], [vhhResolution]).matchedCount, 0);
assert.equal(match([vhhStaff], [{ ...vhhContact, sourceDate: '2026-10-01' }], [vhhResolution]).matchedCount, 0);
assert.equal(assignmentMatchesContactContext(vhhStaff, vhhContact, now), true);

// Actual SQLite persistence: rejection is distinct from clear, revision conflicts
// are checked, and each decision keeps the existing one-read/two-write save cost.
class TracedD1 {
  constructor(sqlite) { this.sqlite = sqlite; this.reads = 0; this.writes = 0; this.sql = []; this.failHistoryOnce = false; }
  prepare(sql) {
    const db = this; db.sql.push(sql);
    return { args: [], bind(...args) { this.args = args; return this; },
      async first() { db.reads++; return db.sqlite.prepare(sql).get(...this.args) || null; },
      async all() { db.reads++; return { results: db.sqlite.prepare(sql).all(...this.args) }; },
      async run() { if (db.failHistoryOnce && sql.includes("INSERT INTO contact_allocation_resolution_history")) { db.failHistoryOnce = false; throw new Error("Injected audit failure"); } const result = db.sqlite.prepare(sql).run(...this.args); db.writes += Number(result.changes); return { meta: { changes: Number(result.changes) } }; },
    };
  }
  async batch(statements) {
    if (this.beforeBatch) { const callback = this.beforeBatch; this.beforeBatch = null; callback(); }
    const before = this.writes;
    this.sqlite.exec('BEGIN');
    try { const results = []; for (const statement of statements) results.push(await statement.run()); this.sqlite.exec('COMMIT'); return results; }
    catch (error) { this.sqlite.exec('ROLLBACK'); this.writes = before; throw error; }
  }
}
const sqlite = new DatabaseSync(':memory:');
sqlite.exec(await readFile(new URL('../migrations/0024_contact_allocation_resolutions.sql', import.meta.url), 'utf8'));
const db = new TracedD1(sqlite);
const saveBase = { sourceId: 'ddh-daily-contact-sheet', sourceDate: '2026-10-02', sourceType: 'ddh', contactKey: key, actorEmail: 'fixture@example.test' };
const save = async (decision, expectedRevision, doctorKey = '') => {
  const before = [db.reads, db.writes];
  const result = await saveContactAllocationResolution(db, { ...saveBase, decision, expectedRevision, doctorKey });
  assert.deepEqual([db.reads - before[0], db.writes - before[1]], [1, 2]);
  return result;
};
const savedReject = await save('rejected', 0);
assert.equal(savedReject.decision, 'rejected');
assert.equal(sqlite.prepare('SELECT active FROM contact_allocation_resolutions').get().active, -1);
assert.deepEqual(await queryContactAllocationResolutions(db, saveBase), [], 'active-only callers still see only assignments');
const savedDecisions = await queryContactAllocationResolutions(db, { ...saveBase, includeInactive: true });
assert.equal(savedDecisions[0].decision, 'rejected');
assert.equal(match(roster, [suggestion], savedDecisions).matchedCount, 0);
await assert.rejects(saveContactAllocationResolution(db, { ...saveBase, expectedRevision: 0, doctorKey: roster[0].person.doctorKey }), /changed while/);
const savedAssign = await save('assigned', 1, roster[0].person.doctorKey);
assert.equal(savedAssign.active, true);
const savedClear = await save('cleared', 2);
assert.equal(savedClear.decision, 'cleared');
assert.equal(sqlite.prepare('SELECT active FROM contact_allocation_resolutions').get().active, 0);
assert.deepEqual(sqlite.prepare('SELECT action FROM contact_allocation_resolution_history ORDER BY revision').all().map(r => r.action), ['rejected', 'reassigned', 'cleared']);
assert.ok(db.sql.every(sql => !/\b(?:CREATE|ALTER|sqlite_master)\b/i.test(sql)), 'no schema queries or migrations during saves');

class MemoryR2 {
  constructor() { this.objects = new Map(); this.puts = 0; this.version = 0; this.failPublication = false; }
  async get(key) { const item = this.objects.get(key); return item ? { etag: item.etag, text: async () => item.text } : null; }
  async put(key, value, options = {}) {
    if (this.failPublication && key.endsWith('/resolutions.json')) throw new Error('Injected publication failure');
    const current = this.objects.get(key);
    if (options.onlyIf?.etagMatches && current?.etag !== options.onlyIf.etagMatches) return null;
    if (options.onlyIf?.etagDoesNotMatch === '*' && current) return null;
    this.puts++; this.version++;
    const item = { text: new TextDecoder().decode(value), etag: `fixture-${this.version}` };
    this.objects.set(key, item); return { etag: item.etag };
  }
}
const r2 = new MemoryR2();
const recentDate = new Intl.DateTimeFormat('en-CA', { timeZone: 'Australia/Melbourne', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
await publishFacilityContactExtract(r2, { sourceId: 'ddh-daily-contact-sheet', sourceDate: recentDate, contacts: [suggestion] });
await publishFacilityContactResolutions(r2, saveBase.sourceId, recentDate, savedDecisions);
const budgetBefore = [db.reads, db.writes, r2.puts];
for (let i = 0; i < 1000; i++) {
  const snapshot = await loadPublishedFacilityContacts(r2, { date: recentDate, facilityKeys: ['DDH'], now: new Date(`${recentDate}T01:00:00Z`) });
  assert.equal(snapshot.resolutions[0].decision, 'rejected');
  // Repeated scoring and rejected matching must perform no storage operations.
  match(roster, contacts, savedDecisions);
}
assert.deepEqual([db.reads, db.writes, r2.puts], budgetBefore);

// Execute the actual UI render functions with real matcher results, without
// importing the app's startup/network lifecycle into this local unit test.
const app = await readFile(new URL('../public/static/app.js', import.meta.url), 'utf8');
const escapeHtml = text => String(text ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const ui = vm.createContext({ escapeHtml, contactAllocationCandidates: (a,c) => contactAllocationCandidates(a,c,{ now }), facilityOverviewState: { contactList: { resolutions: [] }, contactResolutionSaving: false }, facilityOverviewOnShiftTimeLabel: () => '' });
for (const name of ['renderFacilityOverviewContactAllocation', 'renderFacilityOverviewContactResolutionMenu', 'renderFacilityOverviewContactReviewRow']) {
  const start = app.indexOf(`function ${name}(`), end = app.indexOf('\nfunction ', start + 1), asyncEnd = app.indexOf('\nasync function ', start + 1);
  vm.runInContext(app.slice(start, Math.min(end < 0 ? Infinity : end, asyncEnd < 0 ? Infinity : asyncEnd)), ui);
}
const tentative = match(roster, [suggestion]);
const html = ui.renderFacilityOverviewContactAllocation(tentative.assignments[0].contactAllocation);
assert.match(html, /is-tentative/);
assert.match(html, /\*<\/sup>/);
assert.match(html, /Automatic tentative match/);
const menu = ui.renderFacilityOverviewContactResolutionMenu(suggestion, tentative.assignments);
assert.match(menu, /Confirm Tea GUNASEN/);
assert.match(menu, /Sheet name: Thisun \(Tea\)/);
assert.match(menu, /Return to review/);
assert.match(menu, /confidence score/);
assert.doesNotMatch(menu, /probability|%/i);
const conflictMenu = ui.renderFacilityOverviewContactResolutionMenu(match(roster, [suggestion, duplicatePhone]).unmatched[0], roster);
assert.match(conflictMenu, /Correct the conflicting names or phone numbers/);
assert.doesNotMatch(conflictMenu, /contact-resolution-target=/);
ui.facilityOverviewState.contactList.resolutions = [rejected];
assert.match(ui.renderFacilityOverviewContactResolutionMenu(suggestion, match(roster, [suggestion], [rejected]).assignments), /Allow automatic matching again/);
assert.match(ui.renderFacilityOverviewContactAllocation(confirmed.assignments[0].contactAllocation), /Manually confirmed/);
assert.equal(ui.renderFacilityOverviewContactAllocation({ phone: '' }), '');
assert.match(ui.renderFacilityOverviewContactResolutionMenu(vhhContact, match([vhhStaff], [vhhContact]).assignments), /Confirm Tea GUNASEN/);
const hostile = ui.renderFacilityOverviewContactAllocation({ phone: '<img>', contactKey: '" onclick="bad', uncertain: true });
assert.doesNotMatch(hostile, /<img>|contact-resolution="" onclick/);
console.log(`Contact name matching passed: ${positiveCases.length} positive cases; ambiguity, conflicts, ordering, parity, corrections, VHH expiry, rendered controls and zero automatic D1 usage. Explicit saves remain 1 indexed read + 2 row writes.`);

// Execute the actual save action with its existing access gate and real SQLite/
// R2 persistence. Surrounding authentication is stubbed to isolate this action's
// incremental SQL budget; the existing access suite covers account boundaries.
const stateSource = await readFile(new URL('../functions/api/state.js', import.meta.url), 'utf8');
const actionStart = stateSource.indexOf('    if (action === "setContactAllocationResolution")');
const actionEnd = stateSource.indexOf('    if (action === "queryFacilityOverviewStaff")', actionStart);
let site = 'DDH';
let clock = now;
let actionRoster = rawRows;
class FixedDate extends Date { constructor(...args) { super(...(args.length ? args : [clock.getTime()])); } }
const actionUi = vm.createContext({
  Response, Date: FixedDate, facilityAccessKeys, facilityAccessAllows,
  facilityOverviewEnabled: () => true,
  facilityOverviewAccess: async () => ({ mode: 'site', facilityKey: site }),
  facilityOverviewAccessDeniedResponse: () => Response.json({ error: 'Site access denied' }, { status: 403 }),
  normalizeRosterName: value => String(value || '').toUpperCase(),
  sanitizeSourceTypes: values => values.filter(value => ['MMC','MCH','DDH','VHH'].includes(value)),
  sharedReadRouteFor: () => 'shared', sharedFacilityContactsEnabled: true,
  loadPublishedFacilityContacts: (store, options) => loadPublishedFacilityContacts(store, { ...options, now: clock }),
  loadPublishedFacilityDays: async () => ({ rows: actionRoster }), australianDateKey: () => '2026-10-02',
  queryContactAllocationResolutions, contactRosterAssignments, validateContactResolutionSelection, saveContactAllocationResolution,
  facilityBuildSources: () => [site.toLowerCase()], publishFacilityContactResolutions,
});
vm.runInContext(`async function saveAction(context, body) { const action = body.action; const email = 'fixture@example.test'; const account = { record: { email } }; ${stateSource.slice(actionStart, actionEnd)} }`, actionUi);
const actionR2 = new MemoryR2();
await publishFacilityContactExtract(actionR2, { sourceId: 'ddh-daily-contact-sheet', sourceDate: '2026-10-02', contacts: [suggestion] });
const published = await loadPublishedFacilityContacts(actionR2, { date: '2026-10-02', facilityKeys: ['DDH'], now });
const actionContact = published.contacts[0];
const actionContext = { env: { ROSTER_DB: db, ROSTER_FILES: actionR2, FACILITY_SHARED_CONTACTS_BUILD_ENABLED: 'true' } };
const runAction = async (decision, expectedRevision, doctorKey = '', facilityKey = site) => {
  const before = [db.reads, db.writes, actionR2.puts];
  const response = await actionUi.saveAction(actionContext, { action: 'setContactAllocationResolution', facilityKey, date: '2026-10-02', contactKey: actionContact.contactKey, doctorKey, decision, expectedRevision });
  return { status: response.status, data: await response.json(), cost: [db.reads - before[0], db.writes - before[1], actionR2.puts - before[2]] };
};
const actionConfirm = await runAction('assigned', 0, roster[0].person.doctorKey);
assert.equal(actionConfirm.status, 200);
assert.deepEqual(actionConfirm.cost, [3,2,1], 'full confirm action retains baseline 3 reads, 2 row writes, 1 R2 publication');
const actionReject = await runAction('rejected', 1);
assert.equal(actionReject.status, 200);
assert.deepEqual(actionReject.cost, [3,2,1], 'reject is bounded like an existing assignment action');
const publishedReject = await loadPublishedFacilityContacts(actionR2, { date: '2026-10-02', facilityKeys: ['DDH'], now });
assert.equal(match(roster, publishedReject.contacts, publishedReject.resolutions).matchedCount, 0, 'rejection survives publication and reload');
const actionClear = await runAction('cleared', 2);
assert.equal(actionClear.status, 200);
assert.deepEqual(actionClear.cost, [2,2,1], 'clear retains baseline 2 reads, 2 writes, 1 publication');
const actionReassign = await runAction('assigned', 3, roster[1].person.doctorKey);
assert.equal(actionReassign.status, 200);
assert.deepEqual(actionReassign.cost, [3,2,1]);
const staleAction = await runAction('assigned', 3, roster[0].person.doctorKey);
assert.equal(staleAction.status, 409);
assert.equal(staleAction.cost[1], 0);
assert.equal(staleAction.cost[2], 0);
const deniedAction = await runAction('assigned', 4, roster[0].person.doctorKey, 'MMC');
assert.equal(deniedAction.status, 403);
assert.deepEqual(deniedAction.cost, [0,0,0]);
// A changed name/phone produces a new contact key, not inherited confirmation.
await publishFacilityContactExtract(actionR2, { sourceId: 'ddh-daily-contact-sheet', sourceDate: '2026-10-02', contacts: [exactContact] });
const changed = await loadPublishedFacilityContacts(actionR2, { date: '2026-10-02', facilityKeys: ['DDH'], now });
const strongRequest = await actionUi.saveAction(actionContext, { action: 'setContactAllocationResolution', facilityKey: 'DDH', date: '2026-10-02', contactKey: changed.contacts[0].contactKey, doctorKey: roster[1].person.doctorKey, decision: 'assigned', expectedRevision: 0 });
assert.equal(strongRequest.status, 409, 'HTTP save action cannot override a strong automatic match');
site = 'VHH';
actionRoster = [{ sourceType: 'vhh', doctorKey: vhhStaff.person.doctorKey, displayName: vhhStaff.person.displayName, seniority: 'HMO', event: vhhStaff.event }];
await publishFacilityContactExtract(actionR2, { sourceId: 'vhh-shift-phone-allocations', sourceDate: '2026-10-02', doctors: [{ role: 'SSU Dr', name: 'Thisun (Tea)', phone: '12018' }] });
const vhhPublished = await loadPublishedFacilityContacts(actionR2, { date: '2026-10-02', facilityKeys: ['VHH'], now });
const vhhBody = { action: 'setContactAllocationResolution', facilityKey: 'VHH', date: '2026-10-02', contactKey: vhhPublished.contacts[0].contactKey, doctorKey: vhhStaff.person.doctorKey, decision: 'assigned', expectedRevision: 0 };
const vhhSave = await actionUi.saveAction(actionContext, vhhBody);
assert.equal(vhhSave.status, 200, 'Current allocation correction works through the actual save action');
clock = finish;
const expiredSave = await actionUi.saveAction(actionContext, { ...vhhBody, expectedRevision: 1 });
assert.equal(expiredSave.status, 400, 'actual save action rejects a clinician whose shift has finished');
console.log('Actual correction action passed: confirmations, reassignment, durable rejection, stale revisions, site boundaries, VHH active shifts; successful assignment budget unchanged at 3 indexed reads + 2 row writes.');

// Cache publication must not roll back a newer correction, even when concurrent
// snapshots arrive out of order. Conditional R2 writes retry without D1 access.
const raceR2 = new MemoryR2();
await publishFacilityContactExtract(raceR2, { sourceId: 'ddh-daily-contact-sheet', sourceDate: '2026-10-02', contacts: [suggestion] });
const raceBefore = [db.reads, db.writes];
await Promise.all([
  publishFacilityContactResolutions(raceR2, 'ddh-daily-contact-sheet', '2026-10-02', [{ ...assigned, revision: 5 }]),
  publishFacilityContactResolutions(raceR2, 'ddh-daily-contact-sheet', '2026-10-02', [{ ...rejected, revision: 6 }]),
]);
await publishFacilityContactResolutions(raceR2, 'ddh-daily-contact-sheet', '2026-10-02', [{ ...assigned, revision: 4 }]);
const raceLoaded = await loadPublishedFacilityContacts(raceR2, { date: '2026-10-02', facilityKeys: ['DDH'], now });
assert.equal(raceLoaded.resolutions[0].revision, 6);
assert.equal(raceLoaded.resolutions[0].decision, 'rejected');
assert.deepEqual([db.reads, db.writes], raceBefore);

// A publication failure is reported honestly, while the acknowledged local
// rejection is retained until shared storage catches up. Retrying is explicit.
site = 'DDH'; clock = now; actionRoster = rawRows;
await publishFacilityContactExtract(actionR2, { sourceId: 'ddh-daily-contact-sheet', sourceDate: '2026-10-02', contacts: [suggestion] });
actionR2.failPublication = true;
const failedPublication = await runAction('rejected', 4);
assert.equal(failedPublication.status, 503);
assert.equal(failedPublication.data.publicationPending, true);
assert.equal(failedPublication.data.resolution.decision, 'rejected');
assert.deepEqual(failedPublication.cost, [3,2,0], 'failure performs no extra SQL retry or storage write');
const pending = { ...failedPublication.data.resolution, pendingPublication: true };
const localPending = { ...published, resolutions: [pending] };
const stalePublished = await loadPublishedFacilityContacts(actionR2, { date: '2026-10-02', facilityKeys: ['DDH'], now });
const retained = mergeContactResolutionRefresh(localPending, stalePublished);
assert.equal(retained.resolutions.find(r => r.contactKey === pending.contactKey).pendingPublication, true);
assert.equal(match(roster, retained.contacts, retained.resolutions).matchedCount, 0);
assert.deepEqual(mergeContactResolutionRefresh(localPending, { ...stalePublished, sourceDate: '2026-10-03', resolutions: [] }).resolutions, [], 'pending decisions never carry to a different date');
assert.deepEqual(mergeContactResolutionRefresh(localPending, { ...stalePublished, contacts: [exactContact], resolutions: [] }).resolutions, [], 'changed contact keys do not inherit decisions');
actionR2.failPublication = false;
const retried = await runAction('rejected', 5);
assert.equal(retried.status, 200);
assert.deepEqual(retried.cost, [3,2,1]);
const acknowledged = await loadPublishedFacilityContacts(actionR2, { date: '2026-10-02', facilityKeys: ['DDH'], now });
assert.equal(mergeContactResolutionRefresh(localPending, acknowledged).resolutions.find(r => r.contactKey === pending.contactKey).pendingPublication, undefined);
ui.facilityOverviewState.contactList.resolutions = [pending];
assert.match(ui.renderFacilityOverviewContactResolutionMenu(actionContact, match(roster, retained.contacts, retained.resolutions).assignments), /Retry shared update/);
console.log('Concurrent publication and injected failure checks passed: no stale resurrection, no automatic SQL retries, local rejection retained, explicit retry shares the saved decision.');

// Atomic save guarantees: audit failure rolls back the decision; a concurrent
// first save losing the compare-and-swap race creates no phantom history.
db.failHistoryOnce = true;
await assert.rejects(saveContactAllocationResolution(db, { ...saveBase, contactKey: 'audit-failure', expectedRevision: 0, doctorKey: roster[0].person.doctorKey }), /Injected audit failure/);
assert.equal(sqlite.prepare("SELECT COUNT(*) AS n FROM contact_allocation_resolutions WHERE contact_key='audit-failure'").get().n, 0);
db.beforeBatch = () => sqlite.prepare(`INSERT INTO contact_allocation_resolutions (id, source_id, source_date, contact_key, active, revision) VALUES (?, ?, ?, ?, -1, 1)`).run('other-writer', saveBase.sourceId, saveBase.sourceDate, 'first-save-race');
await assert.rejects(saveContactAllocationResolution(db, { ...saveBase, contactKey: 'first-save-race', expectedRevision: 0, doctorKey: roster[0].person.doctorKey }), /changed while/);
assert.equal(sqlite.prepare("SELECT active FROM contact_allocation_resolutions WHERE contact_key='first-save-race'").get().active, -1);
assert.equal(sqlite.prepare("SELECT COUNT(*) AS n FROM contact_allocation_resolution_history WHERE resolution_id LIKE '%first-save-race%'").get().n, 0);
console.log('Atomic correction checks passed: failed audit rolls back both rows; a losing first-save revision race cannot create an allocation or history.');

const allDayCs = { ...staff('Support DOCTOR'), event: { title: 'DDH: CS onsite', start: '2026-10-02', end: '2026-10-03', allDay: true } };
const csContact = sheet('', 'DDH', '49908', 'Clinical Support on-site');
const allDayServer = contactRosterAssignments([{ sourceType: 'ddh', doctorKey: allDayCs.person.doctorKey, displayName: allDayCs.person.displayName, seniority: 'HMO', event: allDayCs.event }]);
assert.deepEqual(fingerprint(match([allDayCs], [csContact])), fingerprint(match(allDayServer, [csContact])));
assert.equal(match(allDayServer, [csContact]).matchedCount, 1, 'all-day DDH support service retains the existing default AM period; VHH still requires timed events');

// First-load/unavailable contact responses must never prevent the roster view.
for (const previous of [null, undefined]) {
  for (const next of [{}, {status:'unavailable',reason:'no-extract'}, {status:'not-current',contacts:[],resolutions:[]}]) {
    assert.equal(mergeContactResolutionRefresh(previous,next),next,'missing previous contacts accept the new empty state');
  }
}
assert.equal(mergeContactResolutionRefresh(null,null),null);
const unidentifiedNext={contacts:[],resolutions:[]};
assert.equal(mergeContactResolutionRefresh({resolutions:[pending]},unidentifiedNext),unidentifiedNext,'unidentified sheets never inherit prior resolutions');
