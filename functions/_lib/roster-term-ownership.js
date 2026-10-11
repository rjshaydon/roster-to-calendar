// Catalogue bounds use shift START dates. Attendance coverage includes
// overnight ends, so it must not decide ownership or retained-window overlap.
export async function rosterTermOwnershipForFiles(db,fileIds) {
  const ids=[...new Set(fileIds.map(String))];
  if(ids.length>32) throw new Error('Roster ownership exceeds its file ceiling.');
  if(!ids.length) return new Map();
  // Each exact file read keeps its 751-row sentinel. D1 aggregates the bounded
  // rows and returns one small summary per file, avoiding JS sorting/hydration.
  const bounded=ids.map(()=>`SELECT * FROM (SELECT file_id,term_start,first_date,last_date
    FROM facility_stream_catalog_contributions WHERE file_id=? LIMIT 751)`).join(' UNION ALL ');
  const rows=(await db.prepare(`WITH bounded AS (${bounded}) SELECT file_id,COUNT(*) AS row_count,
    MIN(NULLIF(term_start,'')) AS first_term,MAX(NULLIF(term_start,'')) AS last_term,
    MIN(first_date) AS first_date,MAX(last_date) AS last_date FROM bounded GROUP BY file_id`).bind(...ids).all()).results;
  const result=new Map();
  for(const row of rows) {
    if(Number(row.row_count)>750) throw new Error('Roster ownership exceeds its compact contribution ceiling.');
    if(!row.first_term || !row.last_term) throw new Error('Prepared roster term ownership is unavailable.');
    result.set(row.file_id,{firstTerm:row.first_term,lastTerm:row.last_term,firstDate:row.first_date,lastDate:row.last_date});
  }
  if(result.size!==ids.length) throw new Error('Prepared roster term ownership is unavailable.');
  return result;
}
export async function rosterTermOwnership(db,fileId) {
  return (await rosterTermOwnershipForFiles(db,[fileId])).get(String(fileId));
}

// One bounded preflight response replaces separate file, source, job and
// catalogue calls. Correlated catalogue reads run only for overlapping files
// from the incoming source; every read retains its 751-row sentinel.
export async function rosterTargetPreflight(db,{sourceType,sourceId,runId,start,end}) {
  const row=await db.prepare(`WITH files AS (
    SELECT f.id,f.source_id,s.derived_state,c.coverage_start,c.coverage_end
    FROM roster_files f INDEXED BY idx_roster_files_source_active
    LEFT JOIN roster_file_status_summaries s ON s.file_id=f.id
    LEFT JOIN roster_file_coverage c ON c.file_id=f.id
    WHERE f.source_type=? AND f.active=1 LIMIT 33
  ) SELECT (SELECT COALESCE(json_group_array(json_object(
    'id',id,'source_id',source_id,'derived_state',derived_state,'coverage_start',coverage_start,'coverage_end',coverage_end,
    'ownership',CASE WHEN source_id=? AND coverage_start<=? AND coverage_end>=? THEN
      (SELECT json_object('rowCount',COUNT(*),'firstTerm',MIN(NULLIF(term_start,'')),'lastTerm',MAX(NULLIF(term_start,'')),
        'firstDate',MIN(first_date),'lastDate',MAX(last_date)) FROM
        (SELECT term_start,first_date,last_date FROM facility_stream_catalog_contributions WHERE file_id=files.id LIMIT 751))
      ELSE NULL END)), '[]') FROM files) AS files_json,
    (SELECT next_batch FROM roster_import_jobs WHERE run_id=?) AS next_batch,
    (SELECT prepared_batch FROM roster_import_jobs WHERE run_id=?) AS prepared_batch,
    (SELECT active_file_id FROM roster_sources WHERE id=?) AS active_file_id`)
    .bind(sourceType,sourceId,end,start,runId,runId,sourceId).first();
  return {...row,files:JSON.parse(row.files_json).map(file=>({...file,ownership:typeof file.ownership==='string'?JSON.parse(file.ownership):file.ownership}))};
}
