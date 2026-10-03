import assert from 'node:assert/strict';
import worker from '../worker/roster-queue-watchdog.js';
import { findmyshiftPollingRosterRanges } from '../functions/_lib/findmyshift.js';
const calls = [], pending = [];
const original = globalThis.fetch;
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
 assert.equal(findmyshiftPollingRosterRanges({}, new Date('2026-10-03T00:00:00Z')).length, 1);
 assert.deepEqual(findmyshiftPollingRosterRanges({}, new Date('2026-10-19T00:00:00Z')).map(r => r.from), ['2026-08-03', '2026-11-02'], 'next-term checking must not stop current-term updates');
 assert.equal(findmyshiftPollingRosterRanges({ FINDMYSHIFT_FROM: '2026-08-03', FINDMYSHIFT_TO: '2026-11-01' }, new Date('2026-10-19T00:00:00Z')).length, 1);
 console.log('Five-minute DDH watchdog and current/next-term checks passed.');
} finally { globalThis.fetch = original; }
