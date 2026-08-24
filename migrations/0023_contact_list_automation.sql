CREATE TABLE IF NOT EXISTS contact_list_files (
  id TEXT PRIMARY KEY,
  source_id TEXT NOT NULL,
  name TEXT NOT NULL,
  size INTEGER NOT NULL DEFAULT 0,
  last_modified INTEGER NOT NULL DEFAULT 0,
  object_key TEXT NOT NULL,
  content_type TEXT NOT NULL DEFAULT '',
  content_hash TEXT NOT NULL,
  provider_version TEXT NOT NULL DEFAULT '',
  provider_modified_at TEXT NOT NULL DEFAULT '',
  received_at TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_contact_list_files_source_version
ON contact_list_files (source_id, provider_version, name);

CREATE INDEX IF NOT EXISTS idx_contact_list_files_source_hash
ON contact_list_files (source_id, content_hash, name);
