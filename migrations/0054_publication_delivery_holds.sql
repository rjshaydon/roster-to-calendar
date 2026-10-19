-- Capacity commitments, not usage. Actual spend remains in metered receipts.
CREATE TABLE IF NOT EXISTS roster_publication_holds (
 source_type TEXT NOT NULL, term_start TEXT NOT NULL, owner_id TEXT NOT NULL,
 owner_kind TEXT NOT NULL CHECK(owner_kind IN ('staging','publication')),
 remaining_writes INTEGER NOT NULL CHECK(remaining_writes>=0),
 remaining_reads INTEGER NOT NULL CHECK(remaining_reads>=0),
 PRIMARY KEY(source_type,term_start)
);
-- Resume pre-existing jobs conservatively; the handler pins exact remaining
-- steps before spending. Publication can progress even if commitments exceed a
-- fresh day's grant; imports cannot spend the committed capacity.
INSERT OR IGNORE INTO roster_publication_holds
SELECT j.source_type,j.term_start,j.request_revision,'publication',
 (4+(json_array_length(j.dates_json)+6)/7+(SELECT COUNT(DISTINCT substr(value,1,7)) FROM json_each(j.dates_json)))*112,
 (4+(json_array_length(j.dates_json)+6)/7+(SELECT COUNT(DISTINCT substr(value,1,7)) FROM json_each(j.dates_json)))*125048
FROM facility_refresh_jobs j WHERE j.source_type IN ('mmc','mch','ddh','vhh') AND j.status='pending' AND EXISTS(
 SELECT 1 FROM facility_term_visibility v WHERE v.source_type=j.source_type AND v.term_start=j.term_start);
