// Catalogue bounds use shift START dates. Attendance coverage includes
// overnight ends, so it must not decide ownership or retained-window overlap.
export async function rosterTermOwnership(db, fileId) {
  const rows = await db.prepare("SELECT term_start, first_date, last_date FROM facility_stream_catalog_contributions WHERE file_id = ? LIMIT 751").bind(fileId).all();
  if (rows.results.length > 750) throw new Error("Roster ownership exceeds its compact contribution ceiling.");
  const terms = [...new Set(rows.results.map((row) => row.term_start).filter(Boolean))].sort();
  if (!terms.length) throw new Error("Prepared roster term ownership is unavailable.");
  return { firstTerm: terms[0], lastTerm: terms.at(-1),
    firstDate: rows.results.map((row) => row.first_date).sort()[0],
    lastDate: rows.results.map((row) => row.last_date).sort().at(-1) };
}
