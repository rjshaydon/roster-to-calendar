import { normalizeEmail } from './email.js';
import { AuthenticationError } from './supabase-auth.js';

const UUID = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i;

function oldCredential(email, record) {
  if (normalizeEmail(record?.email) !== email || !/^[a-f0-9]{64}$/i.test(record?.passwordHash || '')
      || !/^[a-f0-9]{32}$/i.test(record?.passwordSalt || '')) {
    throw new AuthenticationError('auth-migration-input', 'An authenticated legacy account is required.', 409);
  }
}

async function credentialRevision(record) {
  const bytes = new TextEncoder().encode(JSON.stringify([record.passwordSalt, record.passwordHash]));
  return Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), x => x.toString(16).padStart(2, '0')).join('');
}

export async function loadAuthenticationLink(db, email) {
  return db.prepare('SELECT email, provider, provider_subject, phase, credential_revision FROM account_authentication_links WHERE email = ?')
    .bind(normalizeEmail(email)).first();
}

// Caller must already have verified the submitted password. This is an internal
// coordination primitive, not a public API that accepts client-supplied records.
export async function reserveAuthenticationMigration(db, record, options = {}) {
  const email = normalizeEmail(record?.email); oldCredential(email, record);
  const revision = await credentialRevision(record), subject = crypto.randomUUID(), now = options.now || new Date().toISOString();
  await db.prepare(`INSERT INTO account_authentication_links
    (email, provider, provider_subject, phase, credential_revision, created_at, updated_at)
    SELECT email, 'supabase', ?, 'pending', ?, ?, ? FROM account_profiles
    WHERE email = ? AND password_hash = ? AND password_salt = ?
    ON CONFLICT(email) DO NOTHING`)
    .bind(subject, revision, now, now, email, record.passwordHash, record.passwordSalt).run();
  const link = await db.prepare(`SELECT l.email, l.provider, l.provider_subject, l.phase, l.credential_revision
    FROM account_authentication_links l JOIN account_profiles a ON a.email = l.email
    WHERE l.email = ? AND a.password_hash = ? AND a.password_salt = ?`)
    .bind(email, record.passwordHash, record.passwordSalt).first();
  if (!link || link.credential_revision !== revision || link.phase !== 'pending') {
    throw new AuthenticationError('auth-migration-conflict', 'The account changed during its upgrade. Please sign in again.', 409);
  }
  return link;
}

// Proof must come from provider verification, never from browser fields or a
// decoded-but-unverified JWT. Finalize only the exact reserved provider subject.
export async function finishAuthenticationMigration(db, { reservation, record, proof }, options = {}) {
  const email = normalizeEmail(record?.email); oldCredential(email, record);
  if (!reservation || reservation.email !== email || reservation.provider !== 'supabase'
      || !UUID.test(reservation.provider_subject || '') || proof?.subject !== reservation.provider_subject
      || normalizeEmail(proof?.email) !== email || proof?.emailVerified !== true || proof?.disabled === true
      || reservation.credential_revision !== await credentialRevision(record)) {
    throw new AuthenticationError('auth-migration-conflict', 'The account upgrade has not been verified.', 409);
  }
  const now = options.now || new Date().toISOString();
  // D1 batch is transactional. A password change/delete or competing linkage
  // between verification and commit must leave the legacy credential untouched.
  await db.batch([
    db.prepare(`UPDATE account_authentication_links SET phase = 'ready', updated_at = ?
      WHERE email = ? AND provider_subject = ? AND credential_revision = ? AND phase = 'pending'
      AND EXISTS (SELECT 1 FROM account_profiles WHERE email = ? AND password_hash = ? AND password_salt = ?)`)
      .bind(now, email, reservation.provider_subject, reservation.credential_revision, email, record.passwordHash, record.passwordSalt),
    db.prepare(`UPDATE account_profiles SET password_hash = '', password_salt = '', updated_at = ?
      WHERE email = ? AND password_hash = ? AND password_salt = ?
      AND EXISTS (SELECT 1 FROM account_authentication_links WHERE email = ? AND provider_subject = ?
        AND credential_revision = ? AND phase = 'ready')`)
      .bind(now, email, record.passwordHash, record.passwordSalt, email, reservation.provider_subject, reservation.credential_revision),
  ]);
  const finalized = await db.prepare(`SELECT l.provider_subject FROM account_authentication_links l
    JOIN account_profiles a ON a.email = l.email WHERE l.email = ? AND l.provider_subject = ?
    AND l.credential_revision = ? AND l.phase = 'ready' AND a.password_hash = '' AND a.password_salt = ''`)
    .bind(email, reservation.provider_subject, reservation.credential_revision).first();
  if (!finalized) throw new AuthenticationError('auth-migration-conflict', 'The account changed during its upgrade. Please sign in again.', 409);
  return { email, subject: finalized.provider_subject, ready: true };
}
