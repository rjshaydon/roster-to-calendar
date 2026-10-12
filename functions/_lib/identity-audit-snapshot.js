// Internal, read-only cache export. No endpoint is enabled by this module.
export const IDENTITY_AUDIT_PAGE_SIZE = 32;
export async function readIdentityAuditSnapshotPage(db, input = {}) {
  const source = String(input.afterSource || ''), key = String(input.afterKey || '');
  const revision = input.identityRevision;
  if (source.length > 20 || key.length > 200 || revision != null && (!Number.isSafeInteger(revision) || revision < 0)) throw Error('Invalid identity snapshot cursor.');
  // D1 batch is a transaction: the revision and page share one database view.
  const results = await db.batch([
    db.prepare('SELECT revision FROM roster_identity_revision WHERE id=1'),
    db.prepare(`SELECT f.source_type,f.doctor_key,f.display_name,f.fingerprint,f.audited_fingerprint,
      a.person_id,p.status AS person_status FROM roster_identity_features f
      LEFT JOIN roster_person_aliases a ON a.source_type=f.source_type AND a.doctor_key=f.doctor_key
      LEFT JOIN roster_people p ON p.person_id=a.person_id
      WHERE (f.source_type,f.doctor_key)>(?,?) ORDER BY f.source_type,f.doctor_key LIMIT 33`).bind(source,key),
  ]);
  const current = Number(results[0]?.results?.[0]?.revision), rows = results[1]?.results;
  if (!Number.isSafeInteger(current) || !Array.isArray(rows)) throw Error('Identity snapshot unavailable.');
  if (revision != null && revision !== current) throw Object.assign(Error('Identity snapshot changed; restart its read-only export.'),{code:'IDENTITY_AUDIT_EVIDENCE'});
  const page = rows.slice(0, IDENTITY_AUDIT_PAGE_SIZE);
  return {identityRevision:current,rows:page,next:rows.length>IDENTITY_AUDIT_PAGE_SIZE
    ? {afterSource:page.at(-1).source_type,afterKey:page.at(-1).doctor_key} : null};
}
