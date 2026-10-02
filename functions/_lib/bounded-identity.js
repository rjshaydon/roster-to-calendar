import { loadPublishedRosterDoctors } from './facility-overview-cache.js';

export const MAX_ACCOUNT_CLAIMS = 16;
const sources = new Set(['mmc', 'mch', 'ddh', 'vhh', 'casey']);
const marker = claim => `${claim.sourceType}|${claim.key}`;

export async function publishedIdentityDirectory(r2, today) {
  try { return await loadPublishedRosterDoctors(r2, today); }
  catch (error) {
    console.warn('Published identity directory unavailable', { error: String(error.message || error).slice(0, 160) });
    return { preparing: true, doctors: [], missingSources: ['mmc', 'mch', 'ddh', 'vhh'] };
  }
}

export function publishedClaimSeniorities(claims, doctors) {
  const wanted = new Set((claims || []).map(marker));
  return [...new Set((doctors || []).filter(doctor => wanted.has(marker(doctor)))
    .flatMap(doctor => doctor.seniorities || []).filter(Boolean))].sort();
}

// Ownership is checked by exact site/key, including legacy duplicate owners.
// No profile table join or roster history is needed for suggestions.
export async function availableIdentitySuggestions(db, claims, email) {
  if (claims.length > MAX_ACCOUNT_CLAIMS) return [];
  if (!claims.length) return [];
  const predicates = claims.map(() => '(source_type=? AND doctor_key=?)').join(' OR ');
  const result = await db.prepare(`SELECT email,source_type,doctor_key FROM account_claims
    INDEXED BY idx_account_claims_source_doctor_email WHERE ${predicates} LIMIT 33`)
    .bind(...claims.flatMap(claim => [claim.sourceType, claim.key])).all();
  if ((result.results || []).length > 32) return [];
  const occupied = new Set((result.results || []).filter(row => row.email !== email)
    .map(row => `${row.source_type}|${row.doctor_key}`));
  return claims.filter(claim => !occupied.has(marker(claim)));
}

function normaliseClaims(claims) {
  if (!Array.isArray(claims) || claims.length > MAX_ACCOUNT_CLAIMS) throw new Error('Link at most 16 roster names per account.');
  const result = new Map();
  for (const claim of claims) {
    if (!sources.has(claim.sourceType) || !claim.key || claim.key.length > 200 || !claim.displayName || claim.displayName.length > 200) {
      throw new Error('Invalid roster identity.');
    }
    result.set(marker(claim), claim);
  }
  return [...result.values()].sort((a, b) => marker(a).localeCompare(marker(b)));
}

// D1 batches are atomic. Assertions run inside the same transaction as the
// incremental mutation, so concurrent claims cannot acquire the same name.
// json() deliberately aborts and rolls back the batch if an assertion fails.
export async function saveBoundedAccountClaims(db, record, incoming, { adminIssues = record.adminIssues } = {}) {
  const email = String(record.email || '').trim().toLowerCase();
  const previous = normaliseClaims(record.claims || []);
  const claims = normaliseClaims(incoming);
  const before = new Map(previous.map(claim => [marker(claim), claim]));
  const after = new Map(claims.map(claim => [marker(claim), claim]));
  const added = claims.filter(claim => !before.has(marker(claim)));
  const removed = previous.filter(claim => !after.has(marker(claim)));
  // Retain timestamps and display names for unchanged claims; replay is free.
  const next = claims.map(claim => before.get(marker(claim)) || claim);
  const issuesChanged = JSON.stringify(adminIssues || []) !== JSON.stringify(record.adminIssues || []);
  let storedIssues = '';
  if (issuesChanged) {
    const profile = await db.prepare('SELECT admin_issues_json FROM account_profiles WHERE email=?').bind(email).first();
    if (!profile || JSON.stringify(JSON.parse(profile.admin_issues_json || '[]')) !== JSON.stringify(record.adminIssues || [])) throw claimConflict();
    storedIssues = profile.admin_issues_json;
  }
  const ownershipPredicates = claims.map(() => '(source_type=? AND doctor_key=?)').join(' OR ');
  if (claims.length) {
    const conflict = await db.prepare(`SELECT email FROM account_claims INDEXED BY idx_account_claims_source_doctor_email
      WHERE (${ownershipPredicates}) AND email<>? LIMIT 1`)
      .bind(...claims.flatMap(claim => [claim.sourceType, claim.key]), email).first();
    if (conflict) throw claimConflict();
  }
  if (!added.length && !removed.length && !issuesChanged) return { ...record, claims: next, unchanged: true };
  const predicates = previous.map(() => '(source_type=? AND doctor_key=? AND display_name=? AND matched_at=?)').join(' OR ');
  const statements = [db.prepare(`SELECT CASE WHEN
    (SELECT COUNT(*) FROM account_claims WHERE email=?)=?
    ${previous.length ? `AND (SELECT COUNT(*) FROM account_claims WHERE email=? AND (${predicates}))=?` : ''}
    AND EXISTS(SELECT 1 FROM account_profiles WHERE email=?${issuesChanged ? ' AND admin_issues_json=?' : ''})
    THEN 1 ELSE json('identity-conflict') END AS identity_guard`)
    .bind(email, previous.length, ...(previous.length ? [email, ...previous.flatMap(claim => [claim.sourceType, claim.key, claim.displayName, claim.matchedAt || '']), previous.length] : []), email, ...(issuesChanged ? [storedIssues] : []))];
  if (claims.length) statements.push(db.prepare(`SELECT CASE WHEN NOT EXISTS(
    SELECT 1 FROM account_claims INDEXED BY idx_account_claims_source_doctor_email
    WHERE (${ownershipPredicates}) AND email<>?) THEN 1 ELSE json('identity-conflict') END AS identity_guard`)
    .bind(...claims.flatMap(claim => [claim.sourceType, claim.key]), email));
  const now = new Date().toISOString();
  for (const claim of removed) statements.push(db.prepare('DELETE FROM account_claims WHERE email=? AND source_type=? AND doctor_key=?').bind(email, claim.sourceType, claim.key));
  for (const claim of added) statements.push(db.prepare(`INSERT INTO account_claims(email,source_type,doctor_key,display_name,matched_at,updated_at)
    VALUES(?,?,?,?,?,?)`).bind(email, claim.sourceType, claim.key, claim.displayName, claim.matchedAt || now, now));
  if (issuesChanged) statements.push(db.prepare('UPDATE account_profiles SET admin_issues_json=?,updated_at=? WHERE email=? AND admin_issues_json=?')
    .bind(JSON.stringify(adminIssues || []), now, email, storedIssues));
  try { await db.batch(statements); }
  catch (error) {
    if (/malformed JSON/i.test(error.message || '')) throw claimConflict();
    throw error;
  }
  return { ...record, claims: next, adminIssues, updatedAt: now };
}

function claimConflict() {
  const error = new Error('Roster links changed or this name belongs to another account. Reload the account before linking again.');
  error.code = 'IDENTITY_CLAIM_CONFLICT';
  return error;
}
