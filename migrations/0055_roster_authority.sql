-- Compact site/term authority. No clinical-row rewrite or history backfill.
CREATE TABLE roster_term_authority (
 source_type TEXT NOT NULL, term_start TEXT NOT NULL, generation INTEGER NOT NULL,
 file_id TEXT NOT NULL, source_id TEXT NOT NULL, content_revision TEXT NOT NULL,
 activated_at TEXT NOT NULL, PRIMARY KEY(source_type,term_start)
);
CREATE TABLE roster_authority_fences (
 run_id TEXT PRIMARY KEY, source_type TEXT NOT NULL, generations_json TEXT NOT NULL
);
CREATE TABLE roster_provider_observations (
 source_id TEXT NOT NULL, file_name TEXT NOT NULL, provider_version TEXT NOT NULL,
 generations_json TEXT NOT NULL, PRIMARY KEY(source_id,file_name)
);
CREATE TABLE roster_authority_policy (source_type TEXT PRIMARY KEY, enabled INTEGER NOT NULL CHECK(enabled IN (0,1)));
INSERT INTO roster_authority_policy VALUES('mmc',1),('mch',1),('ddh',1),('vhh',1);
INSERT INTO roster_term_authority
 SELECT f.source_type,c.term_start,1,f.id,f.source_id,s.content_revision,f.parsed_at
 FROM roster_files f JOIN facility_stream_catalog_contributions c ON c.file_id=f.id
 JOIN roster_file_status_summaries s ON s.file_id=f.id
 WHERE f.active=1 AND f.source_id IN ('manual-mmc','manual-mch','manual-ddh','manual-vhh')
 GROUP BY f.source_type,c.term_start HAVING COUNT(DISTINCT f.id)=1;
