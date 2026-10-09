import assert from 'node:assert/strict';
import worker from '../worker/roster-queue-watchdog.js';
import { findmyshiftPollingRosterRanges } from '../functions/_lib/findmyshift.js';
const calls = [], pending = [];
const original = globalThis.fetch;
const OriginalDate = globalThis.Date;
try {
 globalThis.fetch = async (url, options) => { calls.push({ url, options }); return Response.json({ ok: true, status: 'unchanged' }); };
 await worker.scheduled({}, {}, { waitUntil: promise => pending.push(promise) });
 assert.equal(calls.length, 0, 'disabled worker must do no work');
 const env = { ROSTER_AUTOMATION_ENABLED: 'true', ROSTER_WATCHDOG_TOKEN: 'synthetic-token' };
 await worker.scheduled({}, env, { waitUntil: promise => pending.push(promise) });
 await Promise.all(pending);
 assert.equal(calls.length, 1, 'one metadata poll; no unscoped queue kick');
 assert.match(calls[0].url, /findmyshift-check$/);
 assert.deepEqual(JSON.parse(calls[0].options.body), { poll: true });
 assert.equal((await worker.fetch(new Request('https://watchdog/check', { method: 'POST' }), env)).status, 401);
 assert.equal(findmyshiftPollingRosterRanges({}, new Date('2026-10-03T00:00:00Z')).length, 2);
 assert.deepEqual(findmyshiftPollingRosterRanges({}, new Date('2026-10-19T00:00:00Z')).map(r => r.from), ['2026-08-03', '2026-11-02'], 'next-term checking must not stop current-term updates');
 assert.equal(findmyshiftPollingRosterRanges({ FINDMYSHIFT_FROM: '2026-08-03', FINDMYSHIFT_TO: '2026-11-01' }, new Date('2026-10-19T00:00:00Z')).length, 1);
 // Exercise the real scheduler: registration every tick, audit only during
 // the Melbourne Sunday window, with no second call or retry per mode.
 const identityEnv = {...env, IDENTITY_REGISTRY_ENABLED:'true', IDENTITY_SCHEDULED_AUDIT_ENABLED:'true'};
 for (const [instant,expected] of [
  ['2026-10-10T16:29:00Z',['register']],
  ['2026-10-10T16:30:00Z',['register','audit']],
  ['2026-10-10T17:29:00Z',['register','audit']],
  ['2026-10-10T17:30:00Z',['register']],
  ['2026-10-09T16:30:00Z',['register']],
 ]) {
  globalThis.Date = class extends OriginalDate { constructor(...args) { super(...(args.length?args:[instant])); } };
  calls.length=0; pending.length=0;
  await worker.scheduled({},identityEnv,{waitUntil:promise=>pending.push(promise)});
  await Promise.all(pending);
  assert.equal(calls.filter(call=>/findmyshift-check$/.test(call.url)).length,1);
  assert.deepEqual(calls.filter(call=>/identity-maintenance$/.test(call.url)).map(call=>JSON.parse(call.options.body).mode),expected,instant);
 }
 calls.length=0;pending.length=0;
 await worker.scheduled({}, {...identityEnv,ROSTER_AUTOMATION_ENABLED:'false'}, {waitUntil:promise=>pending.push(promise)});
 await Promise.all(pending);
 assert.equal(calls.length,0,'paused automation also stops identity maintenance');
 console.log('Automatic identity registration and Melbourne weekly audit boundaries passed.');
 console.log('Five-minute DDH watchdog and current/next-term checks passed.');
} finally { globalThis.fetch = original; globalThis.Date = OriginalDate; }
