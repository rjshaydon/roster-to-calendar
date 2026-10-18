-- Account-wide preference, independent of roster snapshots and device storage.
ALTER TABLE account_profiles ADD COLUMN passkey_reminder_dismissed INTEGER NOT NULL DEFAULT 0
  CHECK (passkey_reminder_dismissed IN (0, 1));
