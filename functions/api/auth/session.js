import { AuthenticationError } from '../../_lib/supabase-auth.js';
import { authenticateExistingAccount } from '../../_lib/account-authentication.js';
import { normalizeEmail } from '../../_lib/email.js';
import {
  accountSessionsEnabled, requireSameOrigin, issueAccountSession, readAccountSession,
  renewAccountSession, revokeAccountSession, revokeOtherAccountSessions, clearSessionCookie, hasSessionCookie,
} from '../../_lib/account-sessions.js';

const headers = { 'Cache-Control': 'no-store', 'Vary': 'Cookie' };
function reply(body, status = 200, setCookie) {
  return Response.json(body, { status, headers: { ...headers, ...(setCookie ? { 'Set-Cookie': setCookie } : {}) } });
}

async function boundedBody(request) {
  if (!request.headers.get('Content-Type')?.toLowerCase().startsWith('application/json')) {
    throw new AuthenticationError('auth-request', 'A JSON request is required.', 400);
  }
  const reader = request.body?.getReader();
  if (!reader) throw new AuthenticationError('auth-request', 'Invalid request.', 400);
  let length = 0; const chunks = [];
  try {
    while (true) {
      const next = await reader.read(); if (next.done) break;
      length += next.value.byteLength;
      if (length > 8192) { await reader.cancel(); throw new AuthenticationError('auth-request', 'Request is too large.', 413); }
      chunks.push(next.value);
    }
    const bytes = new Uint8Array(length); let offset = 0;
    for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.byteLength; }
    try { const body = JSON.parse(new TextDecoder().decode(bytes)); if (!body || Array.isArray(body) || typeof body !== 'object') throw Error(); return body; }
    catch { throw new AuthenticationError('auth-request', 'Invalid request.', 400); }
  } finally { reader.releaseLock(); }
}

export async function onRequestPost(context) {
  if (!accountSessionsEnabled(context.env)) return reply({ error: 'The sign-in upgrade is not enabled.', code: 'auth-disabled' }, 503);
  try {
    requireSameOrigin(context.request);
    if (!context.env.ROSTER_DB) throw new AuthenticationError('auth-unavailable', 'Sign-in is currently unavailable. Please try again later.');
    const body = await boundedBody(context.request), action = String(body.action || '');
    const db = context.env.ROSTER_DB;
    if (action === 'login' || action === 'reauthenticate') {
      // No legacy fallback here. Accounts without a completed verified migration
      // keep their existing journey until that separate rollout is ready.
      const email = normalizeEmail(body.email);
      if (!email || email.length > 254 || typeof body.password !== 'string' || !body.password || body.password.length > 1024) {
        throw new AuthenticationError('invalid-credentials', 'Check your email and password.', 400);
      }
      // Admission precedes D1 and password work. No D1 counter per failed guess.
      // Fail closed until a supported rate-limit binding is staged and verified.
      if (!context.env.AUTH_LOGIN_LIMITER?.limit) throw new AuthenticationError('auth-unavailable', 'Sign-in is currently unavailable. Please try again later.');
      const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(email));
      const key = Array.from(new Uint8Array(digest), x => x.toString(16).padStart(2, '0')).join('');
      const admission = await context.env.AUTH_LOGIN_LIMITER.limit({ key: 'login:' + key });
      if (admission?.success !== true) throw new AuthenticationError('auth-rate-limited', 'Too many sign-in attempts. Please wait before trying again.', 429);
      let oldSession;
      if (action === 'reauthenticate') {
        oldSession = await readAccountSession(db, context.request);
        if (email !== oldSession.email) throw new AuthenticationError('auth-identity-mismatch', 'Sign in using the current account.', 403);
      }
      const proof = await authenticateExistingAccount(context.env, { email, password: body.password });
      const issued = await issueAccountSession(db, { email: proof.email, subject: proof.subject, persistent: body.remember === true });
      if (oldSession) await revokeAccountSession(db, oldSession);
      return reply({ ok: true, email: issued.email }, 200, issued.setCookie);
    }
    if (action === 'logout') {
      // An expired/invalid cookie may always be cleared locally, but cannot
      // select an account or revoke anybody else's sessions.
      if (hasSessionCookie(context.request)) {
        try { await revokeAccountSession(db, await readAccountSession(db, context.request)); }
        catch (error) { if (error.code !== 'auth-session-expired') throw error; }
      }
      return reply({ ok: true }, 200, clearSessionCookie());
    }
    const session = await readAccountSession(db, context.request);
    if (action === 'status') return reply({ ok: true, email: session.email, role: session.role });
    if (action === 'activity') {
      const renewed = await renewAccountSession(db, context.request, session);
      return reply({ ok: true }, 200, renewed.setCookie);
    }
    if (action === 'signOutOtherDevices') {
      await revokeOtherAccountSessions(db, session);
      return reply({ ok: true });
    }
    throw new AuthenticationError('auth-request', 'Unknown account action.', 400);
  } catch (error) {
    if (['d1-statement-budget-exceeded', 'd1-request-deadline-exceeded'].includes(error?.code)) throw error;
    if (error instanceof AuthenticationError) return reply({ error: error.message, code: error.code }, error.status);
    // Never reflect provider, database or request error details into auth replies.
    return reply({ error: 'Sign-in is currently unavailable. Please try again later.', code: 'auth-unavailable' }, 503);
  }
}

export function onRequestGet() { return reply({ error: 'Use POST for account requests.' }, 405); }
