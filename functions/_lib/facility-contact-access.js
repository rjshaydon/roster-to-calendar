const TOKEN_VERSION = 1;
const MAX_TOKEN_LIFETIME_MS = 15 * 60 * 1000;

export async function issueFacilityContactAccessToken(secretValue, options = {}) {
  const secret = String(secretValue || "");
  const facilityKey = String(options.facilityKey || "").trim().toUpperCase();
  if (secret.length < 32 || !/^[A-Z0-9_-]{2,12}$/.test(facilityKey)) return "";
  const now = Number(options.now || Date.now());
  const date = String(options.date || "");
  if (date && !/^\d{4}-\d{2}-\d{2}$/.test(date)) return "";
  const requestedExpiry = Date.parse(String(options.expiresAt || ""));
  const expiresAt = Math.min(
    Number.isFinite(requestedExpiry) ? requestedExpiry : now + MAX_TOKEN_LIFETIME_MS,
    now + MAX_TOKEN_LIFETIME_MS,
  );
  if (expiresAt <= now) return "";
  const payload = encodeBase64Url(JSON.stringify({ v: TOKEN_VERSION, facilityKey, ...(date ? { date } : {}), exp: Math.floor(expiresAt / 1000) }));
  const signature = await sign(secret, payload);
  return signature ? `${payload}.${signature}` : "";
}

export async function verifyFacilityContactAccessToken(secretValue, tokenValue, options = {}) {
  const secret = String(secretValue || "");
  const token = String(tokenValue || "");
  if (secret.length < 32 || !token.includes(".")) return null;
  const [payload, signature, extra] = token.split(".");
  if (!payload || !signature || extra || !(await verify(secret, payload, signature))) return null;
  try {
    const claims = JSON.parse(decodeBase64Url(payload));
    const nowSeconds = Math.floor(Number(options.now || Date.now()) / 1000);
    const facilityKey = String(claims?.facilityKey || "").trim().toUpperCase();
    if (Number(claims?.v) !== TOKEN_VERSION || !facilityKey || Number(claims?.exp || 0) <= nowSeconds) return null;
    if (claims.date && !/^\d{4}-\d{2}-\d{2}$/.test(claims.date)) return null;
    return { facilityKey, ...(claims.date ? { date: claims.date } : {}), expiresAt: new Date(Number(claims.exp) * 1000).toISOString() };
  } catch {
    return null;
  }
}

async function signingKey(secret, usages) {
  return crypto.subtle.importKey("raw", new TextEncoder().encode(secret), { name: "HMAC", hash: "SHA-256" }, false, usages);
}

async function sign(secret, payload) {
  const key = await signingKey(secret, ["sign"]);
  const bytes = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(payload));
  return encodeBase64UrlBytes(new Uint8Array(bytes));
}

async function verify(secret, payload, signature) {
  try {
    const key = await signingKey(secret, ["verify"]);
    return crypto.subtle.verify("HMAC", key, decodeBase64UrlBytes(signature), new TextEncoder().encode(payload));
  } catch {
    return false;
  }
}

function encodeBase64Url(value) {
  return encodeBase64UrlBytes(new TextEncoder().encode(value));
}

function encodeBase64UrlBytes(bytes) {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function decodeBase64Url(value) {
  return new TextDecoder().decode(decodeBase64UrlBytes(value));
}

function decodeBase64UrlBytes(value) {
  const normalized = String(value || "").replace(/-/g, "+").replace(/_/g, "/");
  const binary = atob(normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "="));
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
}
