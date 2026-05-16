CREATE INDEX IF NOT EXISTS idx_roster_events_source_range ON roster_events (source_type, start_date, end_date);
