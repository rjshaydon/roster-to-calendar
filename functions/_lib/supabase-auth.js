import { normalizeEmail } from './email.js';
import { guardedFetch } from './outbound-network.js';

// Server-only. A single secret binding avoids expanding the Pages variable count.
// This module is not an API route and is not enabled in production yet.
export const AUTH_PROJECT_ORIGIN = 'https://nncmejyfkujpfuivlart.supabase.co';
const UUID = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i;
const MAX_RESPONSE_BYTES = 64 * 1024;

export class AuthenticationError extends Error {
  constructor(code, message, status = 503) {
    super(message);
    this.name = 'AuthenticationError';
    this.code = code;
    this.status = status;
  }
}

export function supabaseAuthenticationConfig(env = {}) {
  let config;
  try { config = JSON.parse(env.ROSTER_AUTH_CONFIG || '{}'); }
  catch { throw new AuthenticationError('auth-configuration', 'Authentication is not configured.'); }
  if (config.enabled !== true) throw new AuthenticationError('auth-disabled', 'The authentication upgrade is not enabled.');
  // Pin the approved project. Neither request bodies nor arbitrary environment
  // URLs may redirect a password or administrative credential elsewhere.
  if (config.url !== AUTH_PROJECT_ORIGIN || !/^sb_publishable_[A-Za-z0-9_-]+$/.test(config.publishableKey || '')
      || !/^sb_secret_[A-Za-z0-9_-]+$/.test(config.secretKey || '')) {
    throw new AuthenticationError('auth-configuration', 'Authentication is not configured.');
  }
  return config;
}

function credentials(email, password) {
  const normalized = normalizeEmail(email);
  if (!normalized || normalized.length > 254 || !normalized.includes('@') || /[\s\x00-\x1f\x7f]/.test(normalized)
      || typeof password !== 'string' || !password || new TextEncoder().encode(password).length > 1024) {
    throw new AuthenticationError('invalid-credentials', 'Check your email and password.', 400);
  }
  return { email: normalized, password };
}

function userIdentity(value, expected = {}) {
  const user = value?.user || value;
  const email = normalizeEmail(user?.email), subject = String(user?.id || '');
  if (!UUID.test(subject) || !email || expected.email && email !== normalizeEmail(expected.email)
      || expected.subject && subject !== expected.subject) {
    throw new AuthenticationError('auth-identity-mismatch', 'The account could not be verified.');
  }
  return { subject, email, emailVerified: Boolean(user.email_confirmed_at), disabled: Boolean(user.banned_until && Date.parse(user.banned_until) > Date.now()) };
}

async function boundedJson(response) {
  const declared = Number(response.headers.get('content-length') || 0);
  if (declared > MAX_RESPONSE_BYTES) throw new AuthenticationError('auth-provider-response', 'Authentication returned an invalid response.');
  const reader = response.body?.getReader();
  if (!reader) throw new AuthenticationError('auth-provider-response', 'Authentication returned an invalid response.');
  const chunks = []; let length = 0;
  try {
    while (true) {
      const next = await reader.read();
      if (next.done) break;
      length += next.value.byteLength;
      if (length > MAX_RESPONSE_BYTES) {
        await reader.cancel();
        throw new AuthenticationError('auth-provider-response', 'Authentication returned an invalid response.');
      }
      chunks.push(next.value);
    }
    const bytes = new Uint8Array(length); let offset = 0;
    for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; }
    return JSON.parse(new TextDecoder().decode(bytes));
  } catch (error) {
    if (error instanceof AuthenticationError) throw error;
    throw new AuthenticationError('auth-provider-response', 'Authentication returned an invalid response.');
  } finally { reader.releaseLock(); }
}

async function providerRequest(env, path, { admin = false, method = 'POST', body, allowMissing = false } = {}) {
  const config = supabaseAuthenticationConfig(env);
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 10000);
  try {
    // Modern project keys belong in apikey, not Authorization: Bearer.
    // Never log request/response bodies: these paths handle passwords and tokens.
    const response = await guardedFetch(env, config.url + '/auth/v1/' + path, {
      // Workers supports manual redirect handling, not redirect: 'error'.
      // Never forward credentials to a Location supplied by the provider.
      method, redirect: 'manual', signal: controller.signal,
      headers: { apikey: admin ? config.secretKey : config.publishableKey, 'Content-Type': 'application/json' },
      ...(body === undefined ? {} : { body: JSON.stringify(body) }),
    }, { label: 'Supabase authentication' });
    if (response.status >= 300 && response.status < 400) {
      await response.body?.cancel();
      throw new AuthenticationError('auth-provider-response', 'Authentication returned an invalid response.');
    }
    if (allowMissing && admin && method === 'GET' && response.status === 404) {
      await response.body?.cancel();
      return null;
    }
    if (response.status === 429) throw new AuthenticationError('auth-rate-limited', 'Too many sign-in attempts. Please wait before trying again.', 429);
    if (response.status >= 500) throw new AuthenticationError('auth-provider-unavailable', 'Sign-in is temporarily unavailable. Please try again later.');
    const data = await boundedJson(response);
    if (!response.ok) {
      if (data?.error_code === 'email_not_confirmed' || data?.code === 'email_not_confirmed') {
        throw new AuthenticationError('email-verification-required', 'Confirm your email address before continuing.', 403);
      }
      // No provider messages are reflected: an error might echo private inputs.
      if (!admin && [400, 401, 403].includes(response.status)) {
        throw new AuthenticationError('invalid-credentials', 'Incorrect email or password.', 401);
      }
      throw new AuthenticationError(admin ? 'auth-migration-unconfirmed' : 'auth-provider-response',
        admin ? 'The account upgrade could not be confirmed. Your existing account has not been replaced.' : 'Sign-in could not be completed.');
    }
    return data;
  } catch (error) {
    if (error instanceof AuthenticationError || error?.code === 'LOCAL_OUTBOUND_BLOCKED') throw error;
    // Do not include error.message/cause; network errors can contain request data.
    throw new AuthenticationError('auth-provider-unavailable', 'Sign-in is temporarily unavailable. Please try again later.');
  } finally { clearTimeout(timeout); }
}

export async function signInWithSupabasePassword(env, input) {
  const body = credentials(input.email, input.password);
  const data = await providerRequest(env, 'token?grant_type=password', { body });
  const identity = userIdentity(data, { email: body.email, subject: input.expectedSubject });
  if (identity.disabled) throw new AuthenticationError('auth-account-unavailable', 'This account is unavailable. Contact the administrator.', 403);
  if (!identity.emailVerified) throw new AuthenticationError('email-verification-required', 'Confirm your email address before continuing.', 403);
  if (typeof data.access_token !== 'string' || !data.access_token || typeof data.refresh_token !== 'string' || !data.refresh_token
      || !Number.isFinite(data.expires_in) || data.expires_in <= 0) {
    throw new AuthenticationError('auth-provider-response', 'Authentication returned an invalid response.');
  }
  // Server-only result. Never serialize this object directly to an app response.
  return { ...identity, accessToken: data.access_token, refreshToken: data.refresh_token, expiresIn: data.expires_in };
}

export async function createSupabaseMigrationUser(env, input) {
  const body = credentials(input.email, input.password);
  if (!UUID.test(input.subject || '')) throw new AuthenticationError('auth-migration-input', 'Invalid migration reference.', 400);
  // Do not silently mark legacy emails verified: password possession proves the
  // old app account, not ownership of its email address. Verification is a gate.
  const data = await providerRequest(env, 'admin/users', { admin: true,
    body: { id: input.subject, ...body, email_confirm: false },
  });
  return userIdentity(data, { email: body.email, subject: input.subject });
}

export async function getSupabaseMigrationUser(env, { subject, email, allowMissing = false }) {
  if (!UUID.test(subject || '')) throw new AuthenticationError('auth-migration-input', 'Invalid migration reference.', 400);
  const data = await providerRequest(env, 'admin/users/' + subject, { admin: true, method: 'GET', allowMissing });
  if (data === null && allowMissing) return null;
  return userIdentity(data, { email, subject });
}
