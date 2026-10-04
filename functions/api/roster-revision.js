import { loadCachedSnapshot } from '../_lib/d1-calendar.js';
import { facilityMetadataManifestKey } from '../_lib/facility-overview-cache.js';
import { melbourneDateKey, rosterTermVisible } from '../../public/static/roster-term-policy.js';

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
  const fingerprints = [];
  for (const source of selected) {
    const manifest = await loadCachedSnapshot(context.env.ROSTER_FILES, facilityMetadataManifestKey(source));
    if (!manifest || (manifest.terms || []).length > 64) return new Response('Preparing', { status: 503, headers: { 'Cache-Control': 'no-store' } });
    fingerprints.push([source, manifest.revision || '', (manifest.terms || []).filter(term => rosterTermVisible(term, today)).map(term => [term.termStart, term.staffRevision || ''])]);
  }
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(JSON.stringify([today, fingerprints])));
  const revision = [...new Uint8Array(digest)].map(byte => byte.toString(16).padStart(2, '0')).join('');
  const response = Response.json({ ok: true, revision }, { headers: { 'Cache-Control': 'public, max-age=30, must-revalidate' } });
  if (cache) context.waitUntil(cache.put(key, response.clone()));
  return response;
}
