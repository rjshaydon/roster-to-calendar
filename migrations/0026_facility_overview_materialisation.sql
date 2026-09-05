-- Phase 1: compact ingestion-time facts for At a glance.
-- This migration only creates storage; live reads remain on the old path until
-- the materialised data and cost gates have been verified locally.

ALTER TABLE roster_file_doctors ADD COLUMN provider_staff_id TEXT NOT NULL DEFAULT '';
ALTER TABLE roster_events ADD COLUMN provider_staff_id TEXT NOT NULL DEFAULT '';

CREATE TABLE IF NOT EXISTS roster_file_coverage (
  file_id TEXT PRIMARY KEY,
  source_type TEXT NOT NULL,
  coverage_start TEXT NOT NULL DEFAULT '',
  coverage_end TEXT NOT NULL DEFAULT '',
  content_revision TEXT NOT NULL DEFAULT '',
  staff_digest TEXT NOT NULL DEFAULT '',
  daily_digest TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_roster_file_coverage_active_lookup
ON roster_file_coverage (source_type, coverage_start, coverage_end, file_id);

CREATE TABLE IF NOT EXISTS facility_term_staff_contributions (
  source_type TEXT NOT NULL,
  term_start TEXT NOT NULL,
  doctor_key TEXT NOT NULL,
  file_id TEXT NOT NULL,
  display_name TEXT NOT NULL DEFAULT '',
  seniority TEXT NOT NULL DEFAULT '',
  membership_source TEXT NOT NULL DEFAULT 'roster',
  provider_staff_id TEXT NOT NULL DEFAULT '',
  first_applicable_date TEXT NOT NULL DEFAULT '',
  last_applicable_date TEXT NOT NULL DEFAULT '',
  fact_digest TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT '',
  PRIMARY KEY (source_type, term_start, doctor_key, file_id)
);

CREATE INDEX IF NOT EXISTS idx_facility_term_staff_lookup
ON facility_term_staff_contributions (source_type, term_start, doctor_key);

CREATE TABLE IF NOT EXISTS facility_term_visibility (
  source_type TEXT NOT NULL,
  term_start TEXT NOT NULL,
  visible_from TEXT NOT NULL,
  revision TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT '',
  PRIMARY KEY (source_type, term_start)
);
