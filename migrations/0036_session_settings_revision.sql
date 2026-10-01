-- Constant-cost settings CAS; no data rewrite or roster backfill.
ALTER TABLE account_states ADD COLUMN settings_revision TEXT NOT NULL DEFAULT '';
ALTER TABLE doctor_profiles ADD COLUMN settings_revision TEXT NOT NULL DEFAULT '';
