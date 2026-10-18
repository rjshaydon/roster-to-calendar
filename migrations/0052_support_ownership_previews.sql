-- Short-lived, owner-only prepared review evidence. A successful decision
-- consumes its preview inside the same atomic case/identity transaction.
CREATE TABLE support_ownership_previews (
 preview_id TEXT PRIMARY KEY,
 actor_email TEXT NOT NULL REFERENCES account_profiles(email) ON DELETE CASCADE,
 case_id TEXT NOT NULL REFERENCES support_cases(case_id) ON DELETE CASCADE,
 case_version INTEGER NOT NULL,
 expires_at INTEGER NOT NULL,
 payload_json TEXT NOT NULL
);
CREATE INDEX idx_support_previews_expiry ON support_ownership_previews(actor_email,expires_at,preview_id);
