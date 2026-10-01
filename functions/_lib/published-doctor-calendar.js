import { loadPublishedFacilityRange } from './facility-overview-cache.js';
import { applyAccountHospitalLocations, buildPreviewFromDerivedEvents } from './d1-calendar.js';
import { applyEventOverrides, customEventsToEvents, defaultSettings } from './roster.js';

// Explicit Creator profile views read shared artifacts; they never build D1 snapshots.
export async function loadPublishedDoctorCalendar(r2, profile, { range, today, locations = {}, schemaVersion = 1 } = {}) {
  const known = new Set(['mmc', 'mch', 'ddh', 'vhh']);
  const sources = [...new Set(profile.sourceTypes || [])];
  const aliases = profile.aliases?.length ? profile.aliases : sources.map(sourceType => ({ sourceType, key: profile.doctorKey }));
  if (!sources.length || sources.length > 4 || sources.some(source => !known.has(source)) || aliases.length > 16) {
    throw new Error('Published doctor calendar requires bounded site identities.');
  }
  const published = await loadPublishedFacilityRange(r2, sources, range.startDate, range.endDate, today);
  if (published.preparing || sources.some(source => !published.sourceTypes?.includes(source)) || published.events.length > 50000) {
    return { snapshot: null, snapshotAvailable: false, snapshotStale: false, stale: false, snapshotStatus: 'missing', snapshotSource: 'published-roster', calendarRevision: '' };
  }
  const markers = new Set(aliases.map(alias => `${alias.sourceType}|${alias.key}`));
  const seen = new Set();
  const roster = published.events.filter(row => {
    if (!markers.has(`${row.sourceType}|${row.doctorKey}`)) return false;
    const marker = `${row.sourceType}|${row.doctorKey}|${row.event?.id || JSON.stringify(row.event)}`;
    if (seen.has(marker)) return false;
    seen.add(marker);
    return true;
  }).map(row => row.event);
  const session = profile.state?.session || {};
  const settings = { ...defaultSettings(), ...(session.settings || {}) };
  const events = applyEventOverrides(applyAccountHospitalLocations(roster, locations, { includeLocations: settings.includeLocations !== false }), session.overrides || {});
  const custom = customEventsToEvents(session.customEvents || [], settings, events).filter(event => String(event.start || '').slice(0, 10) <= range.endDate && String(event.end || event.start || '').slice(0, 10) >= range.startDate);
  const bytes = new TextEncoder().encode(JSON.stringify([published.revision, aliases, session, locations]));
  const revision = 'published-profile:' + [...new Uint8Array(await crypto.subtle.digest('SHA-256', bytes))].map(value => value.toString(16).padStart(2, '0')).join('');
  const builtAt = new Date().toISOString();
  return {
    snapshotAvailable: true, snapshotStale: false, stale: false, snapshotStatus: 'ready', snapshotSource: 'published-roster', snapshotBuiltAt: builtAt, snapshotRevision: revision, calendarRevision: revision,
    snapshot: {
      ownerType: 'doctor-profile', ownerId: profile.profileId, schemaVersion, builtAt, buildStamp: 'published-roster',
      preview: buildPreviewFromDerivedEvents([...events, ...custom], { customEventsMaterialized: true }),
      session: { ...session, doctorKey: profile.doctorKey },
      doctorOptions: [{ key: profile.doctorKey, displayName: profile.displayName, sourceTypes: sources, aliases }],
      detectedSources: Object.fromEntries(["mmc", "mch", "ddh", "vhh", "casey"].map(source => [source, sources.includes(source) ? [source] : []])), fileRefs: [], subscriptionFeeds: {}, insightCache: null,
    },
  };
}
