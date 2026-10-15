import assert from 'node:assert/strict';
import { DatabaseSync } from 'node:sqlite';
import { readFile, readdir } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { AUTH_PROJECT_ORIGIN, supabaseAuthenticationConfig, signInWithSupabasePassword, createSupabaseMigrationUser, getSupabaseMigrationUser } from '../functions/_lib/supabase-auth.js';
import { loadAuthenticationLink, reserveAuthenticationMigration, finishAuthenticationMigration } from '../functions/_lib/authentication-migration.js';
import { onRequestPost as stateHandler } from '../functions/api/state.js';

const subject = '11111111-2222-4333-8444-555555555555', otherSubject = 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee';
const email = 'person@example.test', password = 'Synthetic migration fixture only';
const env = { ROSTER_AUTH_CONFIG: JSON.stringify({ enabled: true, url: AUTH_PROJECT_ORIGIN, publishableKey: 'sb_publishable_fixture', secretKey: 'sb_secret_fixture' }) };
const user = { id: subject, email, email_confirmed_at: '2026-10-16T00:00:00Z' };
const login = { user, access_token: 'synthetic-access', refresh_token: 'synthetic-refresh', expires_in: 3600 };
const originalFetch = globalThis.fetch;
const requests = [];
let response = () => Response.json(login);
globalThis.fetch = async (url, init) => { requests.push({ url, init }); return response(); };
try {
  assert.throws(() => supabaseAuthenticationConfig({}), /not enabled/);
  assert.throws(() => supabaseAuthenticationConfig({ ROSTER_AUTH_CONFIG: '{' }), /not configured/);
  for (const url of ['https://attacker.invalid', AUTH_PROJECT_ORIGIN + '/', AUTH_PROJECT_ORIGIN.replace('https:', 'http:'), 'https://numqebixqasrfkyuafzi.supabase.co']) {
    assert.throws(() => supabaseAuthenticationConfig({ ROSTER_AUTH_CONFIG: JSON.stringify({ ...JSON.parse(env.ROSTER_AUTH_CONFIG), url }) }), /not configured/);
  }
  const verified = await signInWithSupabasePassword(env, { email: ' Person@Example.Test ', password, expectedSubject: subject });
  assert.equal(verified.subject, subject);
  assert.equal(verified.emailVerified, true);
  assert.equal(requests[0].url, AUTH_PROJECT_ORIGIN + '/auth/v1/token?grant_type=password');
  assert.equal(requests[0].init.headers.apikey, 'sb_publishable_fixture');
  assert.equal(requests[0].init.headers.Authorization, undefined, 'new keys are not used as bearer JWTs');
  assert.equal(requests[0].init.redirect, 'error', 'credentials cannot follow redirects');
  assert.deepEqual(JSON.parse(requests[0].init.body), { email, password });
  const beforeLocal = requests.length;
  await assert.rejects(() => signInWithSupabasePassword({ ...env, LOCAL_ONLY: 'true' }, { email, password }), e => e.code === 'LOCAL_OUTBOUND_BLOCKED');
  assert.equal(requests.length, beforeLocal, 'local mode cannot transmit even synthetic credentials');
  await assert.rejects(() => signInWithSupabasePassword(env, { email, password: 'x'.repeat(1025) }), e => e.code === 'invalid-credentials');
  await assert.rejects(() => signInWithSupabasePassword(env, { email, password, expectedSubject: otherSubject }), e => e.code === 'auth-identity-mismatch');
  response = () => Response.json({ ...login, user: { ...user, email: 'other@example.test' } });
  await assert.rejects(() => signInWithSupabasePassword(env, { email, password }), e => e.code === 'auth-identity-mismatch');
  response = () => Response.json({ ...login, user: { ...user, email_confirmed_at: null } });
  await assert.rejects(() => signInWithSupabasePassword(env, { email, password }), e => e.code === 'email-verification-required');
  response = () => Response.json({ ...login, user: { ...user, banned_until: '2099-01-01' } });
  await assert.rejects(() => signInWithSupabasePassword(env, { email, password }), e => e.status === 403);
  response = () => Response.json({ ...login, refresh_token: undefined });
  await assert.rejects(() => signInWithSupabasePassword(env, { email, password }), e => e.code === 'auth-provider-response');
  response = () => Response.json({ user });
  await createSupabaseMigrationUser(env, { email, password, subject });
  const creation = requests.at(-1);
  assert.equal(creation.init.headers.apikey, 'sb_secret_fixture');
  assert.deepEqual(JSON.parse(creation.init.body), { id: subject, email, password, email_confirm: false }, 'legacy email ownership is not fabricated');
  await getSupabaseMigrationUser(env, { email, subject });
  assert.equal(requests.at(-1).init.method, 'GET');
  assert.equal(requests.at(-1).init.body, undefined);
  await assert.rejects(() => getSupabaseMigrationUser(env, { email, subject: '../users' }), e => e.code === 'auth-migration-input');
  for (const [status, code] of [[400, 'invalid-credentials'], [401, 'invalid-credentials'], [429, 'auth-rate-limited'], [503, 'auth-provider-unavailable']]) {
    response = () => Response.json({ message: password + ' sb_secret_fixture synthetic-refresh' }, { status });
    await assert.rejects(() => signInWithSupabasePassword(env, { email, password }), e => e.code === code && !e.message.includes(password) && !e.message.includes('sb_secret_'));
  }
  response = () => Response.json({ message: password }, { status: 422 });
  await assert.rejects(() => createSupabaseMigrationUser(env, { email, password, subject }), e => e.code === 'auth-migration-unconfirmed' && !e.message.includes(password));
  response = () => { throw Error('fetch failed with ' + password); };
  await assert.rejects(() => signInWithSupabasePassword(env, { email, password }), e => e.code === 'auth-provider-unavailable' && !e.message.includes(password));
  response = () => new Response('not json');
  await assert.rejects(() => signInWithSupabasePassword(env, { email, password }), e => e.code === 'auth-provider-response');
  response = () => new Response('x'.repeat(65537));
  await assert.rejects(() => signInWithSupabasePassword(env, { email, password }), e => e.code === 'auth-provider-response');
} finally { globalThis.fetch = originalFetch; }

const sqlite = new DatabaseSync(':memory:');
sqlite.exec('PRAGMA foreign_keys = ON');
for (const name of (await readdir(new URL('../migrations/', import.meta.url))).filter(n => n.endsWith('.sql')).sort()) {
  sqlite.exec(await readFile(new URL('../migrations/' + name, import.meta.url), 'utf8'));
}
let failSql = '', beforeBatch;
const queries = [];
const db = { prepare(sql) {
  queries.push(sql);
  return { args: [], bind(...args) { this.args = args; return this; },
    async run() { if (failSql && sql.includes(failSql)) throw Error('Injected transaction failure'); const x = sqlite.prepare(sql).run(...this.args); return { meta: { changes: Number(x.changes) } }; },
    async first() { return sqlite.prepare(sql).get(...this.args) || null; },
    async all() { return { results: sqlite.prepare(sql).all(...this.args) }; },
  };
}, async batch(statements) {
  if (beforeBatch) { const fn = beforeBatch; beforeBatch = null; fn(); }
  sqlite.exec('BEGIN');
  try { const results = []; for (const statement of statements) results.push(await statement.run()); sqlite.exec('COMMIT'); return results; }
  catch (error) { sqlite.exec('ROLLBACK'); throw error; }
} };
const salt = 'a'.repeat(32), hash = createHash('sha256').update(salt + ':' + password).digest('hex');
const record = { email, passwordSalt: salt, passwordHash: hash };
const insertAccount = () => sqlite.prepare('INSERT INTO account_profiles(email,real_name,password_hash,password_salt,role,subscription_token) VALUES(?,?,?,?,?,?)')
  .run(email, 'Fixture Person', hash, salt, 'user', 'preserve-feed-token');
insertAccount();
const reservation = await reserveAuthenticationMigration(db, record);
const retry = await reserveAuthenticationMigration(db, record);
assert.equal(retry.provider_subject, reservation.provider_subject, 'uncertain outcomes keep the same external UUID');
assert.equal(sqlite.prepare('SELECT COUNT(*) n FROM account_authentication_links').get().n, 1);
const proof = { subject: reservation.provider_subject, email, emailVerified: true };
for (const wrong of [{ ...proof, subject: otherSubject }, { ...proof, email: 'other@example.test' }, { ...proof, emailVerified: false }]) {
  await assert.rejects(() => finishAuthenticationMigration(db, { reservation, record, proof: wrong }), e => e.code === 'auth-migration-conflict');
}
assert.equal(sqlite.prepare('SELECT password_hash FROM account_profiles WHERE email=?').get(email).password_hash, hash);
failSql = 'UPDATE account_profiles SET password_hash';
await assert.rejects(() => finishAuthenticationMigration(db, { reservation, record, proof }), /Injected/);
failSql = '';
assert.equal((await loadAuthenticationLink(db, email)).phase, 'pending', 'failed retirement rolls back ready state');
assert.equal(sqlite.prepare('SELECT password_hash FROM account_profiles WHERE email=?').get(email).password_hash, hash);
beforeBatch = () => sqlite.prepare('UPDATE account_profiles SET password_hash=? WHERE email=?').run('b'.repeat(64), email);
await assert.rejects(() => finishAuthenticationMigration(db, { reservation, record, proof }), e => e.code === 'auth-migration-conflict');
assert.equal(sqlite.prepare('SELECT password_hash FROM account_profiles WHERE email=?').get(email).password_hash, 'b'.repeat(64), 'a concurrent password change is never erased');
await assert.rejects(() => reserveAuthenticationMigration(db, record), e => e.code === 'auth-migration-conflict', 'a retry revalidates the current credential');
sqlite.prepare('UPDATE account_profiles SET password_hash=? WHERE email=?').run(hash, email);
await finishAuthenticationMigration(db, { reservation, record, proof });
await finishAuthenticationMigration(db, { reservation, record, proof });
assert.equal((await loadAuthenticationLink(db, email)).phase, 'ready');
const final = sqlite.prepare('SELECT password_hash,password_salt,role,subscription_token,real_name FROM account_profiles WHERE email=?').get(email);
assert.equal(final.password_hash, ''); assert.equal(final.password_salt, '');
assert.equal(final.subscription_token, 'preserve-feed-token'); assert.equal(final.role, 'user'); assert.equal(final.real_name, 'Fixture Person');
// Retired credentials must not be silently replaced by a new arbitrary password
// through the still-present legacy endpoint (a critical migration bypass).
const request = new Request('https://fixture.test/api/state', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ action: 'login', email, password: 'attacker-selected-password', responseMode: 'fast' }) });
const denied = await stateHandler({ request, env: { ROSTER_DB: db }, data: {}, waitUntil() {} });
assert.equal(denied.status, 401);
assert.equal(sqlite.prepare('SELECT password_hash FROM account_profiles WHERE email=?').get(email).password_hash, '');
const plan = sqlite.prepare(`EXPLAIN QUERY PLAN SELECT l.email FROM account_authentication_links l JOIN account_profiles a ON a.email=l.email WHERE l.email=? AND a.password_hash=? AND a.password_salt=?`).all(email, hash, salt).map(x => x.detail).join(';');
assert(!/SCAN\s+(?:l|a|account_profiles|account_authentication_links)\b/.test(plan), plan);
sqlite.prepare('DELETE FROM account_profiles WHERE email=?').run(email);
assert.equal(await loadAuthenticationLink(db, email), null, 'deleting the account removes its provider mapping');
assert(!queries.some(q => /UPDATE.*(?:role|subscription_token)\s*=/i.test(q)), 'migration does not change privileges or feeds');
console.log('Authentication foundation passed: provider destination pinning, bounded/redacted failures, local isolation, verified UUID/email binding, pending retries, transactional credential retirement, concurrent password change, legacy bypass denial and indexed queries. No external network or real accounts.');
