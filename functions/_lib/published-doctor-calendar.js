import { loadPublishedFacilityRange } from './facility-overview-cache.js';
import { applyAccountHospitalLocations, buildPreviewFromDerivedEvents } from './d1-calendar.js';
import { applyEventOverrides, customEventsToEvents, defaultSettings, normalizeRosterName } from './roster.js';

// Explicit Creator profile views read shared artifacts; they never build D1 snapshots.
export async function loadPublishedDoctorCalendar(r2, profile, { range, today, locations = {}, schemaVersion = 1, previousSnapshot = null, ownerType = "doctor-profile", ownerId = profile.profileId } = {}) {
  const known = new Set(['mmc', 'mch', 'ddh', 'vhh', 'casey']);
  const sources = [...new Set(profile.sourceTypes || [])];
  const aliases = profile.aliases?.length ? profile.aliases : sources.map(sourceType => ({ sourceType, key: profile.doctorKey }));
  if (!sources.length && !profile.identityAliasesEnforced || sources.length > 5 || sources.some(source => !known.has(source)) || aliases.length > 16) {
    throw new Error('Published doctor calendar requires bounded site identities.');
  }
  const published = sources.length?await loadPublishedFacilityRange(r2, sources, range.startDate, range.endDate, today):{events:[],revision:'empty-identity',missing:[],visibleTerms:[]};
  if (published.preparing || published.events.length > 50000
    || (published.missing || []).some(item => ["month-unavailable", "staff-unavailable"].includes(item.reason))) {
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
  // Retain historical cached shifts outside published windows. Never retain a
  // cached current-term shift after a replacement or removal.
  const historical = (previousSnapshot?.preview?.events || []).filter(event => {
    const date = String(event.start || '').slice(0, 10);
    const source = String(event.source || '').toLowerCase();
    if(profile.identityAliasesEnforced && !markers.has(`${source}|${normalizeRosterName(event.doctorKey || event.doctor || '')}`)) return false;
    return known.has(source) && date >= range.startDate && date <= range.endDate
      && !(published.visibleTerms || []).some(term => term.sourceType === source && term.termStart <= date && term.termEnd >= date);
  });
  const session = profile.state?.session || {};
  const settings = { ...defaultSettings(), ...(session.settings || {}) };
  const events = applyEventOverrides(applyAccountHospitalLocations([...historical, ...roster], locations, { includeLocations: settings.includeLocations !== false }), session.overrides || {});
  const custom = customEventsToEvents(session.customEvents || [], settings, events).filter(event => String(event.start || '').slice(0, 10) <= range.endDate && String(event.end || event.start || '').slice(0, 10) >= range.startDate);
  const bytes = new TextEncoder().encode(JSON.stringify([published.revision, aliases, session, locations, historical]));
  const revision = 'published-profile:' + [...new Uint8Array(await crypto.subtle.digest('SHA-256', bytes))].map(value => value.toString(16).padStart(2, '0')).join('');
  const builtAt = new Date().toISOString();
  return {
    partial: Boolean(published.missing?.length), missing: published.missing || [],
    snapshotAvailable: true, snapshotStale: false, stale: false, snapshotStatus: 'ready', snapshotSource: 'published-roster', snapshotBuiltAt: builtAt, snapshotRevision: revision, calendarRevision: revision,
    snapshot: {
      ownerType, ownerId, schemaVersion, builtAt, buildStamp: 'published-roster',
      preview: { ...buildPreviewFromDerivedEvents([...events, ...custom], { customEventsMaterialized: true }), publicationMissing: published.missing || [] },
      session: { ...session, doctorKey: profile.doctorKey },
      doctorOptions: [{ key: profile.doctorKey, displayName: profile.displayName, sourceTypes: sources, aliases }],
      detectedSources: Object.fromEntries(["mmc", "mch", "ddh", "vhh", "casey"].map(source => [source, sources.includes(source) ? [source] : []])), fileRefs: [], subscriptionFeeds: {}, insightCache: null,
    },
  };
}

// Never serve an earlier person's cached roster after an explicit reassignment,
// even while a new publication is incomplete. Personal custom events survive.
export function filterSnapshotByIdentityAliases(snapshot,aliases,displayName='') {
 if(!snapshot) return snapshot;
 const markers=new Set(aliases.map(a=>`${a.sourceType}|${normalizeRosterName(a.key)}`));
 const sources=new Set(['mmc','mch','ddh','vhh','casey']);
 const events=(snapshot.preview?.events||[]).filter(event=>!sources.has(String(event.source||'').toLowerCase()) || markers.has(`${String(event.source||'').toLowerCase()}|${normalizeRosterName(event.doctorKey||event.doctor||'')}`));
 return {...snapshot,preview:{...buildPreviewFromDerivedEvents(events,{customEventsMaterialized:true}),publicationMissing:snapshot.preview?.publicationMissing||[]},doctorOptions:aliases.length?[{key:snapshot.session?.doctorKey||'',displayName,sourceTypes:[...new Set(aliases.map(a=>a.sourceType))],aliases}]:[]};
}
