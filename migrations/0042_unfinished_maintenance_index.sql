-- Admission inspects only unfinished/unknown receipts. Completed receipts
-- must not make every optional checkpoint scan all of today's bookkeeping.
CREATE INDEX IF NOT EXISTS idx_maintenance_unfinished
ON roster_maintenance_receipts(utc_day,reconciled_at,request_id)
WHERE finished_at='' OR metadata_complete<>1;
