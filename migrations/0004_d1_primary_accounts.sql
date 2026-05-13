ALTER TABLE account_profiles ADD COLUMN password_salt TEXT NOT NULL DEFAULT '';
ALTER TABLE account_profiles ADD COLUMN password_hash TEXT NOT NULL DEFAULT '';
ALTER TABLE account_profiles ADD COLUMN admin_issues_json TEXT NOT NULL DEFAULT '[]';
ALTER TABLE account_profiles ADD COLUMN local_parser_extensions_json TEXT NOT NULL DEFAULT '[]';
ALTER TABLE account_profiles ADD COLUMN created_at TEXT NOT NULL DEFAULT '';
