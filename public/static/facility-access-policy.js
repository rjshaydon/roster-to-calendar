// Shared scope vocabulary. An empty restricted list never means all hospitals.
export const FACILITY_ACCESS_VERSION = 2;
export function isAllSiteSeniority(value) {
  return ["SMS", "CMO"].includes(String(value || "").trim().toUpperCase());
}
export function facilityAccessKeys(access) {
  return [...new Set((access?.facilityKeys || [access?.facilityKey]).map(value => String(value || "").toUpperCase()).filter(Boolean))].sort();
}
export function facilityAccessAllows(access, facility) {
  return access?.mode === "all" || (["site", "sites"].includes(access?.mode) && facilityAccessKeys(access).includes(String(facility || "").toUpperCase()));
}
export function restrictedFacilityScope(keys) {
  const facilityKeys = [...new Set(keys.map(value => String(value).toUpperCase()))].sort();
  return { mode: facilityKeys.length > 1 ? "sites" : facilityKeys.length ? "site" : "denied", facilityKeys, facilityKey: facilityKeys.length === 1 ? facilityKeys[0] : "" };
}
export function validFacilityDateRange(start, end, maximumDays = 370) {
  const date = value => /^\d{4}-\d{2}-\d{2}$/.test(value || "") && Number.isFinite(Date.parse(`${value}T00:00:00Z`)) && new Date(`${value}T00:00:00Z`).toISOString().slice(0,10) === value;
  return date(start) && date(end) && end >= start && (Date.parse(end) - Date.parse(start)) / 86400000 <= maximumDays;
}
export function nextFacilityDate(date) {
  return new Date(Date.parse(`${date}T00:00:00Z`) + 86400000).toISOString().slice(0,10);
}
let melbourneOffsetFormatter;
function melbourneMidnight(date) {
  // Melbourne midnight occurs before the 02:00/03:00 DST transition.
  const sample = new Date(Date.parse(`${date}T00:00:00Z`) - 10 * 60 * 60 * 1000);
  melbourneOffsetFormatter ||= new Intl.DateTimeFormat("en", { timeZone: "Australia/Melbourne", timeZoneName: "longOffset" });
  const offset = melbourneOffsetFormatter.formatToParts(sample).find(part => part.type === "timeZoneName").value.replace("GMT", "");
  return `${date}T00:00:00${offset}`;
}
// Compute timezone boundaries per segment/date rather than per roster row.
// A term-wide request can contain thousands of rows sharing the same bounds.
export function filterFacilityRowsBySegments(rows, segments) {
  const midnights = new Map();
  const midnight = date => {
    if (!midnights.has(date)) midnights.set(date, melbourneMidnight(date));
    return midnights.get(date);
  };
  const bySource = new Map();
  for (const segment of segments || []) {
    const next = nextFacilityDate(segment.endDate);
    const lower = midnight(segment.startDate), upper = midnight(next);
    const bounds = { lower, upper, lowerStamp: Date.parse(lower), upperStamp: Date.parse(upper),
      allDayLower: segment.startDate, allDayUpper: next,
      allDayLowerStamp: Date.parse(`${segment.startDate}T00:00:00Z`), allDayUpperStamp: Date.parse(`${next}T00:00:00Z`) };
    if (!bySource.has(segment.sourceType)) bySource.set(segment.sourceType, []);
    bySource.get(segment.sourceType).push(bounds);
  }
  const result = [];
  for (const row of rows || []) {
    if (!row?.event) continue;
    const source = String(row.sourceType || row.event.sourceType || row.event.source || "").toLowerCase();
    const authorised = bySource.get(source);
    if (!authorised) continue;
    const event = row.event, allDay = event.allDay === true;
    const start = String(event.start || ""), end = String(event.end || event.start || "");
    const stamp = value => allDay ? Date.parse(`${value.slice(0,10)}T00:00:00Z`)
      : Date.parse(/[Zz]|[+-]\d\d:\d\d$/.test(value) ? value : `${value}${midnight(value.slice(0,10)).slice(-6)}`);
    const startStamp = stamp(start), endStamp = stamp(end);
    for (const bounds of authorised) {
      const lowerStamp = allDay ? bounds.allDayLowerStamp : bounds.lowerStamp;
      const upperStamp = allDay ? bounds.allDayUpperStamp : bounds.upperStamp;
      if (startStamp >= upperStamp || endStamp <= lowerStamp) continue;
      result.push({ ...row, event: { ...event,
        start: startStamp < lowerStamp ? (allDay ? bounds.allDayLower : bounds.lower) : start,
        end: endStamp > upperStamp ? (allDay ? bounds.allDayUpper : bounds.upper) : end } });
    }
  }
  return result;
}
