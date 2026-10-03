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
function melbourneMidnight(date) {
  // Melbourne midnight occurs before the 02:00/03:00 DST transition.
  const sample = new Date(Date.parse(`${date}T00:00:00Z`) - 10 * 60 * 60 * 1000);
  const offset = new Intl.DateTimeFormat("en", { timeZone: "Australia/Melbourne", timeZoneName: "longOffset" }).formatToParts(sample).find(part => part.type === "timeZoneName").value.replace("GMT", "");
  return `${date}T00:00:00${offset}`;
}
// Clip at authorised date boundaries, including overnight shifts. A range
// spanning two rotations must not return shifts from the intervening periods.
export function filterFacilityRowsBySegments(rows, segments) {
  const result = [];
  for (const row of rows || []) {
    if (!row?.event) continue;
    const source = String(row.sourceType || row.event.sourceType || row.event.source || "").toLowerCase();
    for (const segment of segments || []) {
      if (source !== segment.sourceType) continue;
      const event = row.event;
      const allDay = event.allDay === true;
      const start = String(event.start || "");
      const end = String(event.end || event.start || "");
      const lower = allDay ? segment.startDate : melbourneMidnight(segment.startDate);
      const upper = allDay ? nextFacilityDate(segment.endDate) : melbourneMidnight(nextFacilityDate(segment.endDate));
      const stamp = value => allDay ? Date.parse(`${value.slice(0,10)}T00:00:00Z`) : Date.parse(/[Zz]|[+-]\d\d:\d\d$/.test(value) ? value : `${value}${melbourneMidnight(value.slice(0,10)).slice(-6)}`);
      if (stamp(start) >= stamp(upper) || stamp(end) <= stamp(lower)) continue;
      result.push({ ...row, event: { ...event, start: stamp(start) < stamp(lower) ? lower : start, end: stamp(end) > stamp(upper) ? upper : end } });
    }
  }
  return result;
}
