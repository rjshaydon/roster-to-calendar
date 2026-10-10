-- Preserve unknown costs. New receipts carry a bounded execution deadline;
-- recovery uses settled account totals, never invented per-request row counts.
ALTER TABLE roster_maintenance_receipts ADD COLUMN purpose TEXT NOT NULL DEFAULT '';
ALTER TABLE roster_maintenance_receipts ADD COLUMN deadline_at TEXT NOT NULL DEFAULT '';
ALTER TABLE roster_maintenance_receipts ADD COLUMN recover_after TEXT NOT NULL DEFAULT '';
ALTER TABLE roster_maintenance_receipts ADD COLUMN reconciled_at TEXT NOT NULL DEFAULT '';
CREATE INDEX idx_maintenance_recovery ON roster_maintenance_receipts(utc_day,reconciled_at,recover_after);
