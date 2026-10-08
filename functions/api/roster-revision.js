import { melbourneDateKey } from '../../public/static/roster-term-policy.js';

const sources = new Set(['mmc', 'mch', 'ddh', 'vhh']);

// Public change fingerprint only. No names, shifts, claims, revisions/keys or
// account identifiers leave this endpoint; actual calendars still authenticate.
export async function onRequestGet(context) {
  if (context.env.ROSTER_VISIBLE_REFRESH_ENABLED !== 'true') return new Response('Unavailable', { status: 503 });
  const url = new URL(context.request.url);
  const selected = [...new Set(String(url.searchParams.get('sites') || '').split(','))].sort();
  if (!selected.length || selected.length > 4 || selected.some(source => !sources.has(source))) return new Response('Invalid sites', { status: 400 });
  const today = melbourneDateKey();
  const key = new Request(`${url.origin}/api/roster-revision?sites=${selected.join(',')}&day=${today}`);
  const cache = globalThis.caches?.default;
  const cached = await cache?.match(key);
  if (cached) return cached;
  // Only storage version tags are needed here. Reading/decompressing whole
  // manifests made this cheap polling route exceed the free CPU allowance.
  const fingerprints = await Promise.all(selected.map(async source => {
    const object = await context.env.ROSTER_FILES.head(`facility-overview/v1/${source}/manifest.json`);
    return object?.etag ? [source, object.etag] : null;
  }));
  if (fingerprints.some(value => !value)) return new Response('Preparing', { status: 503, headers: { 'Cache-Control': 'no-store' } });
  const identity = context.env.IDENTITY_REVIEW_ENABLED === 'true'
    ? await context.env.ROSTER_FILES.head('identity/revision.json') : null;
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(JSON.stringify([today, fingerprints, ...(context.env.IDENTITY_REVIEW_ENABLED==='true'?[identity?.etag||'']:[])])));
  const revision = [...new Uint8Array(digest)].map(byte => byte.toString(16).padStart(2, '0')).join('');
  const response = Response.json({ ok: true, revision }, { headers: { 'Cache-Control': 'public, max-age=30, must-revalidate' } });
  if (cache) context.waitUntil(cache.put(key, response.clone()));
  return response;
}
