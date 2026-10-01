import { australianTermStartForDate } from "./d1-calendar.js";
import { facilityBuildSources } from "./facility-rollout.js";

export function automaticFacilityPublicationEnabled(env, sourceType) {
  return String(env.FACILITY_AUTOMATIC_PUBLICATION_ENABLED || "") === "true" && facilityBuildSources(env, [sourceType]).length === 1;
}

export function facilityRefreshStatements(db, sourceType, dates, signature) {
  const terms = new Map();
  for (const date of [...new Set(dates)].sort()) {
    const term = australianTermStartForDate(date);
    if (!term || !/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error("Invalid facility refresh date.");
    if (!terms.has(term)) terms.set(term, []);
    terms.get(term).push(date);
  }
  if (dates.length > 180 || terms.size > 3) throw new Error("Facility refresh exceeds its affected-date budget.");
  return [...terms].map(([term, affected]) => db.prepare(`INSERT INTO facility_refresh_jobs
    (source_type, term_start, dates_json, content_signature, request_revision, updated_at) VALUES (?, ?, ?, ?, ?, ?)
    ON CONFLICT(source_type, term_start) DO UPDATE SET
      dates_json = CASE WHEN facility_refresh_jobs.status = 'complete' THEN excluded.dates_json ELSE (SELECT json_group_array(value) FROM (SELECT value FROM json_each(facility_refresh_jobs.dates_json) UNION SELECT value FROM json_each(excluded.dates_json) ORDER BY value)) END,
      content_signature = excluded.content_signature, request_revision = excluded.request_revision,
      status = 'pending', plan_json = '', next_batch = 0, next_month = 0, last_error = '', updated_at = excluded.updated_at
    WHERE facility_refresh_jobs.content_signature <> excluded.content_signature`)
    .bind(sourceType, term, JSON.stringify(affected), signature, crypto.randomUUID(), new Date().toISOString()));
}

export function facilityTermDates(start, end) {
  const dates = [];
  let cursor = start;
  while (cursor <= end && dates.length <= 180) {
    dates.push(cursor);
    const next = new Date(`${cursor}T12:00:00Z`);
    next.setUTCDate(next.getUTCDate() + 1);
    cursor = next.toISOString().slice(0, 10);
  }
  if (dates.length > 180) throw new Error("Facility refresh range exceeds 180 days.");
  return dates;
}
