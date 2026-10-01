import { loadPublishedFacilityRange } from './facility-overview-cache.js';
import { normalizeRosterName } from './roster.js';

const SOURCES = ['mmc','mch','ddh','vhh'];
const DAY = 86400000;
const key = value => normalizeRosterName(String(value || ''));
const dateMs = value => {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return NaN;
  const time = Date.parse(`${value}T12:00:00Z`);
  return Number.isFinite(time) && new Date(time).toISOString().slice(0,10) === value ? time : NaN;
};

export function validateCachedInsightOptions(options) {
  const start = dateMs(options.startDate), end = dateMs(options.endDate);
  if (!Number.isFinite(start) || !Number.isFinite(end) || end < start || end-start > 180*DAY) return 'Choose a date range of up to six months.';
  for (const field of ['doctorKeys','excludeDoctorKeys','overlapDoctorKeys']) {
    if (!Array.isArray(options[field] || []) || (options[field] || []).length > 64) return 'Choose up to 64 roster identities.';
  }
  if (!Array.isArray(options.sourceTypes || [])) return 'Choose valid hospitals.';
  if ((options.sourceTypes || []).some(source => !SOURCES.includes(source))) return 'This hospital does not have published colleague data.';
  return '';
}

// Matches the existing date-overlap semantics, using published facts only.
// Leave/unknown entries are filtered by the reviewed caller's working predicate.
export function selectCachedInsightRows(rows, options = {}) {
  const include = new Set((options.doctorKeys || []).map(key));
  const exclude = new Set((options.excludeDoctorKeys || []).map(key));
  const mine = new Set((options.overlapDoctorKeys || []).map(key));
  const presence = new Set();
  const dates = row => {
    const start = Math.max(dateMs(String(row.event?.start || '').slice(0,10)), dateMs(options.startDate));
    const end = Math.min(dateMs(String(row.event?.end || row.event?.start || '').slice(0,10)), dateMs(options.endDate));
    if (!Number.isFinite(start) || !Number.isFinite(end) || end-start > 120*DAY) return [];
    const result=[];
    for(let time=start;time<=end;time+=DAY) result.push(`${row.sourceType}|${new Date(time).toISOString().slice(0,10)}`);
    return result;
  };
  if (mine.size) for (const row of rows) if (mine.has(key(row.doctorKey))) for (const day of dates(row)) presence.add(day);
  return rows.filter(row => !exclude.has(key(row.doctorKey)) && (!include.size || include.has(key(row.doctorKey))) && (!mine.size || dates(row).some(day => presence.has(day))));
}

export async function loadCachedRosterInsights(r2, options, today, isWorking) {
  const issue = validateCachedInsightOptions(options);
  if (issue) return { ok: false, invalid: true, error: issue, coworkers: [], doctors: [] };
  const sources = options.sourceTypes?.length ? options.sourceTypes : SOURCES;
  // Include the bounded maximum attendance span for a shift starting before
  // the requested range; trim by attendance below. No event-history SQL.
  const lookback = new Date(dateMs(options.startDate)-120*DAY).toISOString().slice(0,10);
  const published = await loadPublishedFacilityRange(r2, sources, lookback, options.endDate, today);
  if (published.preparing) return { ok: false, unavailable: true, coworkers: [], doctors: [] };
  if (published.events.length > 50000) return { ok: false, unavailable: true, coworkers: [], doctors: [] };
  const rows = selectCachedInsightRows(published.events.filter(row => isWorking(row.event, row.sourceType) && String(row.event?.start || "").slice(0,10) <= options.endDate && String(row.event?.end || row.event?.start || "").slice(0,10) >= options.startDate), options);
  const doctors = [...new Map(rows.map(row => [`${row.sourceType}|${key(row.doctorKey)}`, { doctorKey: row.doctorKey, displayName: row.displayName, sourceType: row.sourceType }])).values()]
    .sort((a,b) => String(a.displayName).localeCompare(String(b.displayName)) || a.sourceType.localeCompare(b.sourceType));
  return { ok: true, coworkers: rows, doctors, revision: published.revision, source: 'published-roster' };
}
