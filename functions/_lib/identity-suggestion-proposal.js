// A subsequent guarded/pinned audit transaction will use this statement. It
// deliberately never reopens an unchanged human rejection, even across retries.
export function identitySuggestionProposalStatement(db,candidate,now) {
  return db.prepare(`INSERT INTO roster_identity_candidates(pair_key,left_json,right_json,evidence_json,fingerprint,status,updated_at)
    VALUES(?,?,?,json_set(?,'$.changedEvidence',json(CASE WHEN EXISTS(SELECT 1 FROM roster_identity_candidates WHERE pair_key=?) THEN 'true' ELSE 'false' END)),?,'pending',?)
    ON CONFLICT(pair_key) DO UPDATE SET left_json=excluded.left_json,right_json=excluded.right_json,evidence_json=excluded.evidence_json,
      fingerprint=excluded.fingerprint,status='pending',reviewer='',updated_at=excluded.updated_at
    WHERE roster_identity_candidates.fingerprint<>excluded.fingerprint`)
    .bind(candidate.pairKey,JSON.stringify(candidate.left),JSON.stringify(candidate.right),JSON.stringify(candidate.evidence),candidate.pairKey,candidate.fingerprint,now);
}
