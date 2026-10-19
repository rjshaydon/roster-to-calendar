-- Legacy historical gap-fill files can coexist with automatic term fragments.
-- They were never a whole-term takeover. Do not infer new authority for them.
-- New manual activations have an immutable manual fence and are excluded.
DELETE FROM roster_term_authority
 WHERE generation=1 AND source_id LIKE 'manual-%'
 AND NOT EXISTS(SELECT 1 FROM roster_authority_fences p WHERE p.run_id='manual:'||roster_term_authority.file_id)
 AND EXISTS(SELECT 1 FROM facility_stream_catalog_contributions c JOIN roster_files f ON f.id=c.file_id
  WHERE c.source_type=roster_term_authority.source_type AND c.term_start=roster_term_authority.term_start
  AND f.active=1 AND f.source_id NOT LIKE 'manual-%');
