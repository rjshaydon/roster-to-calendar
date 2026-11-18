// Catalogue bounds use shift START dates. Attendance coverage includes
// overnight ends, so it must not decide ownership or retained-window overlap.
export async function rosterTermOwnershipForFiles(db,fileIds) {
  const ids=[...new Set(fileIds.map(String))];
  if(ids.length>32) throw new Error('Roster ownership exceeds its file ceiling.');
  if(!ids.length) return new Map();
  // One bounded indexed read per file without a compound SELECT. D1 can have
  // a lower compound-term limit than local SQLite. Preserve the 751-row
  // sentinel separately for each file and return only the aggregate.
  const rows=(await db.prepare(`SELECT ids.value AS file_id,
    (SELECT json_object('row_count',COUNT(*),'first_term',MIN(NULLIF(term_start,'')),
      'last_term',MAX(NULLIF(term_start,'')),'first_date',MIN(first_date),'last_date',MAX(last_date))
      FROM (SELECT term_start,first_date,last_date FROM facility_stream_catalog_contributions
        WHERE file_id=ids.value LIMIT 751)) AS summary
    FROM json_each(?) ids`).bind(JSON.stringify(ids)).all()).results;
  const result=new Map();
  for(const row of rows) {
    const summary=typeof row.summary==='string'?JSON.parse(row.summary):row.summary;
    if(Number(summary.row_count)>750) throw new Error('Roster ownership exceeds its compact contribution ceiling.');
    if(!summary.first_term || !summary.last_term) throw new Error('Prepared roster term ownership is unavailable.');
    result.set(row.file_id,{firstTerm:summary.first_term,lastTerm:summary.last_term,firstDate:summary.first_date,lastDate:summary.last_date});
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
