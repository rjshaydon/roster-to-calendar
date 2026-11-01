import assert from 'node:assert/strict';
import { DatabaseSync } from 'node:sqlite';
import { readFile, readdir } from 'node:fs/promises';
import { runInNewContext } from 'node:vm';
import { onRequestPost } from '../functions/api/state.js';
import { seedSupportFixture } from './support-request-fixture.mjs';
import { AUTH_PROJECT_ORIGIN } from '../functions/_lib/supabase-auth.js';

// Exercise real authenticated responses using a disposable database, no services.
const sql = new DatabaseSync(':memory:');
sql.exec('PRAGMA foreign_keys=ON');
for (const name of (await readdir(new URL('../migrations/', import.meta.url))).filter(name => name.endsWith('.sql')).sort()) {
  sql.exec(await readFile(new URL('../migrations/' + name, import.meta.url), 'utf8'));
}
const queries = [];
const db = {
  prepare(query) {
    return {
      args: [],
      bind(...args) { this.args = args; return this; },
      async first() { queries.push(query); return sql.prepare(query).get(...this.args) || null; },
      async all() { queries.push(query); return { results: sql.prepare(query).all(...this.args) }; },
      async run() { queries.push(query); return { meta: { changes: Number(sql.prepare(query).run(...this.args).changes) } }; },
    };
  },
  async batch(statements) { return Promise.all(statements.map(statement => statement.run())); },
};
const actors = await seedSupportFixture(db);
const env = {
  ROSTER_DB: db, BOUNDED_MANUAL_ROSTER_ENABLED: 'true', CREATOR_STARTUP_HYDRATION_ENABLED: 'false',
  ROSTER_AUTH_CONFIG: JSON.stringify({ enabled: true, sessionsEnabled: true, url: AUTH_PROJECT_ORIGIN, publishableKey: 'sb_publishable_fixture', secretKey: 'sb_secret_fixture' }),
};
async function request(actor, body) {
  const response = await onRequestPost({
    env, data: {}, waitUntil() {},
    request: new Request('https://fixture.test/api/state', {
      method: 'POST', headers: { Origin: 'https://fixture.test', Cookie: actors[actor].cookie, 'Content-Type': 'application/json' },
      body: JSON.stringify({ expectedAccount: actors[actor].email, ...body }),
    }),
  });
  const data = await response.json();
  assert.equal(response.status, 200, JSON.stringify(data));
  return data;
}
const responses = [];
for (const actor of ['admin', 'holder']) {
  queries.length = 0;
  const data = await request(actor, { action: 'login', responseMode: 'fast' });
  assert.deepEqual(data.subscription, { token: 'synthetic-feed-' + actor, enabled: true }, 'fast login must supply the existing subscription');
  assert.equal(queries.some(query => /^\s*(INSERT|UPDATE|DELETE)\b/i.test(query)), false, 'reading a token must not write account or roster state');
  responses.push(data);
}
responses.push(await request('admin', { action: 'adminLoadUser', targetEmail: actors.holder.email, responseMode: 'fast' }));
assert.equal(responses.at(-1).subscription.token, 'synthetic-feed-holder', 'entered accounts use their own feed');
assert.equal((await request('admin', { action: 'loadAccountContext' })).subscription.token, 'synthetic-feed-admin', 'contained Creator context retains its feed');

const source = await readFile(new URL('../public/static/app.js', import.meta.url), 'utf8');
function section(start, end) {
  const index = source.indexOf(start);
  assert.ok(index >= 0, start);
  const finish = source.indexOf(end, index);
  assert.ok(finish > index, end);
  return source.slice(index, finish);
}
const noop = () => {};
const calls = [];
const browser = {
  currentSubscription: null, currentCalendarRevision: '', currentUserRole: 'user', currentDefaultDoctorKey: '',
  currentSnapshotOwnerType: '', currentSnapshotOwnerId: '', currentUserEmail: actors.admin.email,
  normalizeEmail: value => String(value || '').toLowerCase(), normalizeRosterName: value => value,
  viewedAccountEmail: () => actors.admin.email, sanitizeRosterClaims: value => value,
  rememberSessionBaseline: noop, primeInsightsAccessForCurrentView: noop, sanitizeFacilityOverviewAccess: value => value,
  applyFacilityOverviewSiteScope: noop, syncFacilityOverviewAccess: noop, saveLocalAccountIdentity: noop, markFacilityOverviewAccessReady: noop,
  calendarTransitionStillCurrent: () => true, applyAvailableRosterDoctorsFromData: noop,
  applyCloudStateContext: () => assert.fail('fast login must not need deferred context'), applyCloudStateSnapshot: async () => {},
  activeCalendarMode: () => 'creator-account', selectedDoctor: () => ({ key: 'DOCTOR', displayName: 'Synthetic Doctor' }),
  pendingExportMode: 'full', pendingExportRange: {}, pendingExportHospitals: ['mmc'], exportModalBody: {},
  exportHospitalOptions: () => ['mmc'], escapeHtml: value => value, currentSnapshotStale: true,
  exportConfigForMode: () => browser.config, config: { mode: 'full' },
  setStatus: message => calls.push(['status', message]), closeExportModal: () => calls.push(['close']),
  refreshSnapshotInBackground: async () => { calls.push(['refresh']); },
  snapshotCloudSavePayload: () => ({ session: {} }), normalizeExportRangeState: value => value,
  defaultExportRangeState: () => ({}), exportSourceEntries: () => [], importRefForWorkspace: value => value,
  saveCloudState: async payload => calls.push(['save', payload]), queueBackgroundCloudStateSave: payload => calls.push(['save', payload]),
  navigator: { clipboard: { writeText: async url => calls.push(['copy', url]) } },
  window: { location: { origin: 'https://fixture.test' } }, URL,
  buildBrowserExportEvents: async () => [{ id: 'shift' }], exportIcs: () => 'calendar', downloadIcs: () => calls.push(['download']),
};
runInNewContext([
  section('function sanitizeSubscription(', 'function applyIssueConfig('),
  section('function applyCloudStateIdentity(', 'function applyAvailableRosterDoctorsFromData('),
  section('async function applyCloudStateData(', 'function calendarRevisionView('),
  section('function renderExportModal(', 'function downloadIcs('),
  section('function subscriptionUrl(', 'function formatUserSites('),
].join('\n'), browser);

for (const data of responses) {
  await browser.applyCloudStateData(data, { deferContext: true });
  assert.equal(browser.currentSubscription.token, data.subscription.token, 'subscription is applied before deferred context');
  browser.renderExportModal();
  assert.doesNotMatch(browser.exportModalBody.innerHTML, /data-export-action="copy" disabled/, 'Copy URL is enabled after fast login');
  for (const config of [{ mode: 'full' }, { mode: 'range', startDate: '2026-11-01', endDate: '2026-11-30', allFuture: false }]) {
    browser.config = config;
    for (const action of ['apple', 'copy']) {
      calls.length = 0;
      const pending = browser.handleExportAction(action);
      const protocol = action === 'apple' ? 'webcal:' : 'https:';
      const expected = `${protocol}//fixture.test/api/feed?token=${data.subscription.token}&view=${config.mode}`;
      if (action === 'apple') assert.equal(browser.window.location.href, expected, 'Apple Calendar navigation starts in the click turn');
      else assert.deepEqual(calls.find(call => call[0] === 'copy'), ['copy', expected], 'clipboard access starts in the click turn');
      await pending;
      assert.equal(calls.some(call => call[0] === 'refresh'), false, 'subscription actions do not await calendar refresh');
      assert.ok(calls.some(call => call[0] === 'save'), 'feed settings are still saved');
    }
  }
}
browser.activeCalendarMode = () => 'doctor-profile';
assert.equal(browser.canCopySubscriptionUrl(), false, 'unclaimed profiles cannot reuse the Creator token');
for (const action of ['copy', 'apple']) {
  calls.length = 0;
  browser.window.location.href = '';
  await browser.handleExportAction(action);
  assert.equal(calls.some(call => call[0] === 'copy'), false);
  assert.equal(browser.window.location.href, '');
}
browser.activeCalendarMode = () => 'creator-account';
browser.currentSubscription = browser.sanitizeSubscription({ token: '', enabled: true });
assert.equal(browser.canCopySubscriptionUrl(), false, 'a missing token cannot enable subscription exports');
browser.config = { mode: 'full' };
calls.length = 0;
await browser.handleExportAction('download');
assert.ok(calls.find(call => call[0] === 'refresh'), 'file exports still refresh stale events');
assert.ok(calls.find(call => call[0] === 'download'));
sql.close();
console.log('Export options passed fast Creator/user login, entered accounts, full/range URLs, immediate Apple/clipboard actions, profile restrictions and file exports.');
