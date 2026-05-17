CREATE TABLE IF NOT EXISTS canonical_doctors (
  canonical_key TEXT PRIMARY KEY,
  display_name TEXT NOT NULL,
  source_type TEXT NOT NULL DEFAULT '',
  source_types_json TEXT NOT NULL DEFAULT '[]',
  aliases_json TEXT NOT NULL DEFAULT '[]',
  has_events INTEGER NOT NULL DEFAULT 0,
  updated_at TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS custom_events (
  owner_email TEXT NOT NULL,
  id TEXT NOT NULL,
  title TEXT NOT NULL,
  start_date TEXT NOT NULL,
  end_date TEXT NOT NULL,
  all_day INTEGER NOT NULL DEFAULT 0,
  start_time TEXT NOT NULL DEFAULT '',
  end_time TEXT NOT NULL DEFAULT '',
  location TEXT NOT NULL DEFAULT '',
  include INTEGER NOT NULL DEFAULT 1,
  updated_at TEXT NOT NULL DEFAULT '',
  PRIMARY KEY (owner_email, id)
);

CREATE INDEX IF NOT EXISTS idx_custom_events_owner_range ON custom_events (owner_email, start_date, end_date);
