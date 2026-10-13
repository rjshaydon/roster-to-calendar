import { melbourneMidnight, nextFacilityDate } from './facility-access-policy.js';
import { isClinicalSupportRosterEvent } from './roster.js';
// Published working facts used by both comparison readers. UI matching modes
// remain explicit: a shared roster day is different from an intersecting shift.
export function comparisonWorkingEvent(event, { facilityKey = '', includeClinicalSupport = true } = {}) {
  const text = `${event?.title || ''} ${event?.rawValue || ''}`.toLowerCase();
  if (['unknown'].includes(String(event?.status || '').toLowerCase()) || String(event?.kind || '').toLowerCase() === 'unknown') return false;
  if (/\b(?:leave|conference|cme|annual|sick|personal|study|exam|sabbatical|parental|long service)\b/.test(text) || text.includes('phnw') || text.includes('public holiday')) return false;
  if (String(facilityKey).toUpperCase() === 'DDH' && /\b(?:hith|vhh)\b/.test(text)) return false;
  if (!includeClinicalSupport && isClinicalSupportRosterEvent(event)) return false;
  return true;
}

export function normalizeComparisonEvent(event) {
  if (!event?.allDay && /T00:00/.test(String(event?.end || '')) && Date.parse(event.end) <= Date.parse(event.start)) {
    return { ...event, end: melbourneMidnight(nextFacilityDate(String(event.end).slice(0,10))) };
  }
  return event;
}

// Discovery historically includes both calendar dates of an overnight shift.
// Its final UI matcher still uses the roster date, preserving existing results.
export function comparisonPresence(events, sourceType) {
  const months = Object.create(null);
  for (const original of events) {
    const event=normalizeComparisonEvent(original);
    if (!comparisonWorkingEvent(event, { facilityKey: sourceType })) continue;
    const start = String(event.start || '').slice(0, 10), end = String(event.end || event.start || '').slice(0, 10);
    const a = Date.parse(start + 'T12:00:00Z'), b = Date.parse(end + 'T12:00:00Z');
    if (!Number.isFinite(a) || !Number.isFinite(b) || b < a || b - a > 120 * 86400000) continue;
    for (let time = a; time <= b; time += 86400000) {
      const day = new Date(time).toISOString().slice(0, 10), month = day.slice(0, 7);
      months[month] = ((months[month] || 0) | (1 << (Number(day.slice(8)) - 1))) >>> 0;
    }
  }
  return months;
}

export function comparisonDayMask(month, startDate, endDate) {
  let mask = 0;
  for (let day = 1; day <= 31; day++) {
    const date = `${month}-${String(day).padStart(2, '0')}`;
    if (date >= startDate && date <= endDate) mask = (mask | (1 << (day - 1))) >>> 0;
  }
  return mask;
}

export function comparisonIntervalMatches(indexed, range, compare) {
  if (!indexed.length || indexed.some(entry => !entry.intervals.length)) return [];
  const lower = new Date(`${range.startDate}T00:00:00`), upper = new Date(`${range.endDate}T00:00:00`);
  upper.setDate(upper.getDate() + 1);
  const facilities = [...new Set(indexed[0].intervals.map(interval => interval.facility))];
  const matches = [];
  for (const facility of facilities) {
    const groups = indexed.map(entry => entry.intervals.filter(interval => interval.facility === facility).sort((a, b) => a.start - b.start));
    if (groups.some(group => !group.length)) continue;
    const visit = (index, start, end, people) => {
      if (index === indexed.length) { if (start < end) matches.push({ facility, start, end, people }); return; }
      for (const interval of groups[index]) {
        if (interval.start >= end) break;
        if (interval.end <= start) continue;
        const a = interval.start > start ? interval.start : start, b = interval.end < end ? interval.end : end;
        if (a < b) visit(index + 1, a, b, [...people, { doctor: indexed[index].doctor, event: interval.event }]);
      }
    };
    for (const interval of groups[0]) {
      const start = interval.start > lower ? interval.start : lower, end = interval.end < upper ? interval.end : upper;
      if (start < end) visit(1, start, end, [{ doctor: indexed[0].doctor, event: interval.event }]);
    }
  }
  return matches.sort(compare);
}
