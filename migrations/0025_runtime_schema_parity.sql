-- Complete the explicit migration history for tables that were historically
-- created by runtime schema repair. Ordinary requests must never execute this DDL.

CREATE TABLE IF NOT EXISTS subscription_tokens (
  token TEXT PRIMARY KEY,
  email TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_subscription_tokens_email
ON subscription_tokens (email);

CREATE TABLE IF NOT EXISTS parser_rules (
  id TEXT PRIMARY KEY,
  scope TEXT NOT NULL DEFAULT 'global',
  email TEXT NOT NULL DEFAULT '',
  source_type TEXT NOT NULL,
  seniority TEXT NOT NULL DEFAULT '',
  code TEXT NOT NULL,
  title TEXT NOT NULL DEFAULT '',
  rule_json TEXT NOT NULL DEFAULT '{}',
  updated_at TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_parser_rules_scope
ON parser_rules (scope, email, source_type);

CREATE TABLE IF NOT EXISTS parser_rule_suggestions (
  id TEXT PRIMARY KEY,
  email TEXT NOT NULL DEFAULT '',
  status TEXT NOT NULL DEFAULT 'pending',
  suggestion_json TEXT NOT NULL DEFAULT '{}',
  created_at TEXT NOT NULL DEFAULT '',
  updated_at TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS issue_dismissals (
  email TEXT NOT NULL,
  fingerprint TEXT NOT NULL,
  dismissed_at TEXT NOT NULL DEFAULT '',
  PRIMARY KEY (email, fingerprint)
);

CREATE TABLE IF NOT EXISTS console_messages (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  actor_email TEXT NOT NULL DEFAULT '',
  message TEXT NOT NULL DEFAULT '',
  is_error INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_console_messages_created_at
ON console_messages (created_at DESC, id DESC);

CREATE TABLE IF NOT EXISTS issue_ignores (
  fingerprint TEXT PRIMARY KEY,
  ignored_at TEXT NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS roster_dispatches (
  id TEXT PRIMARY KEY,
  status TEXT NOT NULL DEFAULT 'requested',
  reason TEXT NOT NULL DEFAULT '',
  github_run_id TEXT NOT NULL DEFAULT '',
  requested_at TEXT NOT NULL DEFAULT '',
  accepted_at TEXT NOT NULL DEFAULT '',
  started_at TEXT NOT NULL DEFAULT '',
  completed_at TEXT NOT NULL DEFAULT '',
  retry_after TEXT NOT NULL DEFAULT '',
  attempt_count INTEGER NOT NULL DEFAULT 0,
  last_error TEXT NOT NULL DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_roster_dispatches_status_retry
ON roster_dispatches (status, retry_after DESC);
