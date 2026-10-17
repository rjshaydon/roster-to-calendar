import {authenticationSettings} from './authentication-settings.js';
import { AuthenticationError } from './supabase-auth.js';
import { normalizeEmail } from './email.js';

export const SESSION_COOKIE = '__Host-roster_session';
export const SESSION_IDLE_MS = 30 * 86400000;
export const SESSION_OUTER_MS = 180 * 86400000;
export const SESSION_RENEW_MS = 6 * 3600000;
export const RECENT_AUTH_MS = 10 * 60000;
const TOKEN = /^[a-f0-9]{64}$/;
const MAX_SESSIONS = 20;

export function accountSessionsEnabled(env = {}) {
  try { const c = authenticationSettings(env); return c.enabled === true && c.sessionsEnabled === true; }
  catch { return false; }
}

export function requireSameOrigin(request) {
  const origin = new URL(request.url).origin;
  if (new URL(request.url).protocol !== 'https:' || request.headers.get('Origin') !== origin
      || !['same-origin', 'none', null].includes(request.headers.get('Sec-Fetch-Site'))) {
    throw new AuthenticationError('auth-origin', 'Open the app directly and try again.', 403);
  }
}

// An assertion from the client, never account authority. Shared browser cookies
// may change before another tab receives its storage notification.
export function requireExpectedAccount(session,expected) {
  if(expected===undefined)return;
  if(typeof expected!=='string'||!expected||normalizeEmail(expected)!==session.email) {
    throw new AuthenticationError('auth-identity-mismatch','The signed-in account changed. Please sign in again.',403);
  }
}

export function hasSessionCookie(request) {
  return (request.headers.get('Cookie') || '').split(';').some(part => part.trim().split('=')[0] === SESSION_COOKIE);
}

function cookieToken(request) {
  const values = (request.headers.get('Cookie') || '').split(';').map(x => x.trim())
    .filter(x => x.split('=')[0] === SESSION_COOKIE).map(x => x.slice(SESSION_COOKIE.length + 1));
  if (values.length !== 1 || !TOKEN.test(values[0])) throw expired();
  return values[0];
}

function expired() { return new AuthenticationError('auth-session-expired', 'Please sign in again.', 401); }

async function tokenHash(token) {
  return Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(token))), x => x.toString(16).padStart(2, '0')).join('');
}

function time(options) {
  const now = options.now ?? Date.now();
  if (!Number.isSafeInteger(now) || now < 0) throw Error('Invalid server time.');
  return now;
}

function cookie(token, persistent, lifetime) {
  return `${SESSION_COOKIE}=${token}; Path=/; Secure; HttpOnly; SameSite=Lax${persistent ? `; Max-Age=${Math.max(0, Math.floor(lifetime / 1000))}` : ''}`;
}
export function clearSessionCookie() { return `${SESSION_COOKIE}=; Path=/; Secure; HttpOnly; SameSite=Lax; Max-Age=0`; }

// Called only after server-side authentication. The caller never accepts a
// browser-supplied provider subject or authentication timestamp as proof.
export async function issueAccountSession(db, { email, subject, persistent = false, credentialVersion = 0 }, options = {}) {
  email = normalizeEmail(email);
  if (!email || !/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i.test(subject || '')) throw expired();
  const now = time(options), token = Array.from(crypto.getRandomValues(new Uint8Array(32)), x => x.toString(16).padStart(2, '0')).join('');
  const hash = await tokenHash(token);
  const results = await db.batch([
    db.prepare(`INSERT INTO account_sessions(token_hash,email,authenticated_at,idle_expires_at,absolute_expires_at,renew_after,persistent)
      SELECT ?,a.email,?,?,?,?,? FROM account_profiles a JOIN account_authentication_links l ON l.email=a.email
      WHERE a.email=? AND l.provider_subject=? AND l.phase='ready' AND a.password_hash='' AND a.password_salt='' AND l.credential_version=? AND l.credential_operation_id=''`)
      .bind(hash, now, now + SESSION_IDLE_MS, now + SESSION_OUTER_MS, now + SESSION_RENEW_MS, persistent ? 1 : 0, email, subject, credentialVersion),
    db.prepare(`DELETE FROM account_sessions WHERE email=? AND token_hash IN
      (SELECT token_hash FROM account_sessions WHERE email=? AND token_hash<>? ORDER BY authenticated_at DESC,token_hash DESC LIMIT -1 OFFSET ?)`)
      .bind(email, email, hash, MAX_SESSIONS - 1),
  ]);
  if (Number(results[0]?.meta?.changes) !== 1) throw expired();
  return { setCookie: cookie(token, persistent, SESSION_IDLE_MS), email };
}

export async function readAccountSession(db, request, options = {}) {
  const hash = await tokenHash(cookieToken(request)), now = time(options);
  const session = await db.prepare(`SELECT s.token_hash,s.email,s.authenticated_at,s.idle_expires_at,s.absolute_expires_at,s.renew_after,s.persistent,
    a.role FROM account_sessions s JOIN account_profiles a ON a.email=s.email WHERE s.token_hash=?
    AND s.idle_expires_at>? AND s.absolute_expires_at>?`).bind(hash, now, now).first();
  if (!session) throw expired();
  return session;
}

export function requireRecentAuthentication(session, options = {}) {
  if (time(options) - session.authenticated_at > RECENT_AUTH_MS) {
    throw new AuthenticationError('auth-reauthentication-required', 'Please sign in again to make this change.', 401);
  }
}

// Explicit foreground user activity only. Routine reads/polls never call this.
// A stable opaque token avoids rotation races across tabs and lost responses.
// Reauthentication replaces it; renewal cannot reset authenticated_at/outer cap.
export async function renewAccountSession(db, request, session, options = {}) {
  requireSameOrigin(request);
  const now = time(options);
  if (session.idle_expires_at <= now || session.absolute_expires_at <= now) throw expired();
  if (session.renew_after > now) return { renewed: false };
  const expiry = Math.min(now + SESSION_IDLE_MS, session.absolute_expires_at);
  const result = await db.prepare(`UPDATE account_sessions SET idle_expires_at=?,renew_after=?
    WHERE token_hash=? AND renew_after<=? AND idle_expires_at>? AND absolute_expires_at>?`)
    .bind(expiry, now + SESSION_RENEW_MS, session.token_hash, now, now, now).run();
  return { renewed: Number(result.meta?.changes) === 1,
    setCookie: session.persistent ? cookie(cookieToken(request), true, expiry - now) : undefined };
}

export async function revokeAccountSession(db, session) {
  await db.prepare('DELETE FROM account_sessions WHERE token_hash=?').bind(session.token_hash).run();
}

export async function revokeOtherAccountSessions(db, session, options = {}) {
  requireRecentAuthentication(session, options);
  const result=await db.prepare('DELETE FROM account_sessions WHERE email=? AND token_hash<>?').bind(session.email, session.token_hash).run();
  const revoked=Number(result.meta?.changes);
  if(!Number.isSafeInteger(revoked)||revoked<0)throw new AuthenticationError('auth-revocation-unconfirmed','Sign-out could not be confirmed. Please try again.');
  return revoked;
}

export async function cleanExpiredAccountSessions(db, options = {}) {
  const result = await db.prepare(`DELETE FROM account_sessions WHERE token_hash IN
    (SELECT token_hash FROM account_sessions WHERE idle_expires_at<=? ORDER BY idle_expires_at,token_hash LIMIT 100)`)
    .bind(time(options)).run();
  return Number(result.meta?.changes || 0);
}
