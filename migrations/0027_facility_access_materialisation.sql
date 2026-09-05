-- Short-lived, account-scoped At a glance access decisions. These records
-- replace repeated current-term roster-event scans once the guarded reader is
-- enabled. One row per subject keeps expiry refreshes bounded.

CREATE TABLE IF NOT EXISTS facility_access_sessions (
  subject_email TEXT PRIMARY KEY,
  access_date TEXT NOT NULL,
  subject_revision TEXT NOT NULL,
  access_json TEXT NOT NULL,
  expires_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_facility_access_sessions_expiry
ON facility_access_sessions (expires_at);
