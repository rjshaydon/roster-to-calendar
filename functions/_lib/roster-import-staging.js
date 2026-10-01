import { rosterTermOwnership } from "./roster-term-ownership.js";
import { boundedRosterEventStatements, boundedRosterPresenceStatements, startDerivedRosterFileSave, refreshFacilityOverviewMaterializationForFile, australianTermStartForDate, australianTermEndForStart, facilitySmsMembershipStatement } from "./d1-calendar.js";
import { rosterImportDigest, validateRosterImportManifest, validateRosterImportBatch, ROSTER_BATCH_FACT_LIMIT, ROSTER_BATCH_BYTE_LIMIT, ROSTER_IMPORT_WRITE_COST as COST } from "./roster-import-batches.js";
import { reserveRosterMaintenanceBudget } from "./roster-maintenance-budget.js";
import { facilityRefreshStatements, facilityTermDates } from "./facility-refresh-queue.js";

// Internal execution primitives used only by the source- and budget-gated protocol.
export async function beginBoundedRosterImport(db, runId, file, plan, options = {}) {
  validateRosterImportManifest(plan.manifest);
  if (!runId || plan.manifest.fileId !== file.id || plan.manifest.sourceId !== file.sourceId || await rosterImportDigest(plan.manifest) !== plan.revision) throw new Error("Staged import plan does not match its queued file.");
  if (plan.manifest.doctors.length > 512 || plan.manifest.eventCount > 25000 || plan.manifest.issueCount > 5000) throw new Error("Staged import plan exceeds total safety bounds.");
  const existing = await db.prepare("SELECT * FROM roster_import_jobs WHERE run_id = ?").bind(runId).first();
  if (existing && (existing.plan_revision !== plan.revision || existing.file_id !== file.id || existing.source_id !== file.sourceId)) throw new Error("Staged import plan changed; a new run is required.");
  if (existing?.initialized === 1) return { duplicate: true, nextBatch: existing.next_batch, preparedBatch: existing.prepared_batch, completed: Boolean(existing.activated), retiredFileIds: JSON.parse(existing.retired_file_ids_json || "[]") };
  const existingFile = await db.prepare("SELECT active FROM roster_files WHERE id = ?").bind(file.id).first();
  if (Number(existingFile?.active) === 1) throw new Error("Cannot stage over an active roster.");
  if (!await reserveImportWrites(db, plan.manifest.doctors.length * (COST.doctor + COST.fileDoctor) + COST.control)) return { deferred: true };
  await db.prepare("INSERT OR IGNORE INTO roster_import_jobs (run_id, file_id, source_id, plan_revision, manifest_json) VALUES (?, ?, ?, ?, ?)")
    .bind(runId, file.id, file.sourceId, plan.revision, JSON.stringify(plan.manifest)).run();
  if (options.allowReplacement) {
    const active = await db.prepare("SELECT id, parsed_at FROM roster_files INDEXED BY idx_roster_files_source_active WHERE source_type=? AND active=1 ORDER BY id LIMIT 33").bind(file.sourceType).all();
    if (active.results.length > 32) throw new Error("Roster replacement exceeds the active-file limit.");
    await db.prepare("UPDATE roster_import_jobs SET promotion_fence_json=? WHERE run_id=? AND initialized=0 AND promotion_fence_json=''").bind(JSON.stringify(active.results), runId).run();
  }
  const claimedAt = new Date().toISOString();
  const claim = await db.prepare("UPDATE roster_import_jobs SET initialize_started_at = ? WHERE run_id = ? AND initialized = 0 AND (initialize_started_at = '' OR initialize_started_at < ?)")
    .bind(claimedAt, runId, new Date(Date.now() - 600000).toISOString()).run();
  if (Number(claim.meta?.changes || 0) !== 1) return { deferred: true, reason: "initialization-in-progress" };
  try {
    await startDerivedRosterFileSave(db, { ...file, active: false }, plan.manifest.doctors);
    await db.prepare("UPDATE roster_import_jobs SET initialized = 1, initialize_started_at = '' WHERE run_id = ? AND plan_revision = ? AND initialize_started_at = ?").bind(runId, plan.revision, claimedAt).run();
  } catch (error) {
    await db.prepare("UPDATE roster_import_jobs SET initialize_started_at = '' WHERE run_id = ? AND initialized = 0 AND initialize_started_at = ?").bind(runId, claimedAt).run();
    throw error;
  }
  return { nextBatch: 0 };
}

export async function stageBoundedRosterBatch(db, runId, file, revision, batch) {
  const job = await db.prepare("SELECT * FROM roster_import_jobs WHERE run_id = ?").bind(runId).first();
  if (!job?.initialized || job.plan_revision !== revision || job.file_id !== file.id || job.source_id !== file.sourceId) throw new Error("Staged batch does not match its initialized job.");
  const manifest = JSON.parse(job.manifest_json);
  validateRosterImportBatch(batch, manifest);
  const expected = manifest.batches[batch.index];
  const hash = await rosterImportDigest(batch);
  if (!expected || expected.hash !== hash) throw new Error("Staged batch differs from the pinned plan.");
  if (batch.facts + 16 > ROSTER_BATCH_FACT_LIMIT || new TextEncoder().encode(JSON.stringify(batch)).length > ROSTER_BATCH_BYTE_LIMIT) throw new Error("Staged batch exceeds its safety budget.");
  const receipt = await db.prepare("SELECT * FROM roster_import_batch_receipts WHERE run_id = ? AND batch_index = ?").bind(runId, batch.index).first();
  if (receipt) {
    if (receipt.payload_hash !== hash || receipt.plan_revision !== revision) throw new Error("Conflicting staged batch receipt.");
    return { duplicate: true, nextBatch: job.next_batch };
  }
  if (batch.index !== job.next_batch) throw new Error("Staged batches must be processed in order.");
  const storedFile = await db.prepare("SELECT active FROM roster_files WHERE id = ?").bind(file.id).first();
  if (!storedFile || Number(storedFile.active) === 1) throw new Error("Staged file must remain inactive.");
  const rows = boundedRosterEventStatements(db, file, batch.doctors, batch.eventsByDoctor, batch.issuesByDoctor);
  if (rows.eventCount !== expected.eventCount || rows.issueCount !== expected.issueCount) throw new Error("Staged facts failed normalization; import remains inactive.");
  if (rows.statements.length + 12 > 256) throw new Error("Staged transaction exceeds the D1 statement budget.");
  if (!await reserveImportWrites(db, rows.eventCount * COST.event + rows.issueCount * COST.issue + COST.control)) return { deferred: true };
  // One db.batch, rather than the legacy helper's multiple transactions:
  // a receipt can never survive without all of its corresponding facts.
  await db.batch([...rows.statements,
    db.prepare("INSERT INTO roster_import_batch_receipts (run_id, batch_index, plan_revision, payload_hash) VALUES (?, ?, ?, ?)").bind(runId, batch.index, revision, hash),
    db.prepare("UPDATE roster_import_jobs SET next_batch = next_batch + 1, event_count = event_count + ?, issue_count = issue_count + ? WHERE run_id = ? AND next_batch = ?")
      .bind(rows.eventCount, rows.issueCount, runId, batch.index),
    db.prepare("UPDATE roster_file_status_summaries SET event_count = ? WHERE file_id = ? AND active = 0 AND derived_state = 'building'")
      .bind(expected.indexedEventCount, file.id),
  ]);
  return { nextBatch: batch.index + 1, eventCount: expected.indexedEventCount };
}

export async function prepareBoundedRosterPresence(db, runId, file, revision, batch) {
  const job = await db.prepare("SELECT * FROM roster_import_jobs WHERE run_id = ?").bind(runId).first();
  if (!job || job.plan_revision !== revision || job.file_id !== file.id || job.source_id !== file.sourceId) throw new Error("Presence preparation does not match the staged job.");
  const manifest = JSON.parse(job.manifest_json);
  validateRosterImportBatch(batch, manifest);
  if (job.next_batch !== manifest.batches.length) throw new Error("All facts must be staged before presence preparation.");
  if (manifest.batches[batch.index]?.hash !== await rosterImportDigest(batch)) throw new Error("Presence batch differs from the pinned plan.");
  if (batch.index < job.prepared_batch) return { duplicate: true, nextBatch: job.prepared_batch };
  if (batch.index !== job.prepared_batch) throw new Error("Presence batches must be processed in order.");
  const stored = await db.prepare("SELECT active FROM roster_files WHERE id = ?").bind(file.id).first();
  if (!stored || Number(stored.active) === 1) throw new Error("Presence preparation requires an inactive file.");
  const presence = boundedRosterPresenceStatements(db, file, batch.doctors, batch.eventsByDoctor);
  if (presence.count !== batch.presenceRows || presence.count + 16 > ROSTER_BATCH_FACT_LIMIT) throw new Error("Presence expansion exceeds its pinned batch budget.");
  if (!await reserveImportWrites(db, presence.count * COST.presence + COST.control)) return { deferred: true };
  await db.batch([...presence.statements, db.prepare("UPDATE roster_import_jobs SET prepared_batch = prepared_batch + 1 WHERE run_id = ? AND prepared_batch = ?").bind(runId, batch.index)]);
  return { nextBatch: batch.index + 1 };
}

export async function prepareBoundedRosterMetadata(db, runId, revision) {
  const job = await db.prepare("SELECT * FROM roster_import_jobs WHERE run_id = ?").bind(runId).first();
  if (!job || job.plan_revision !== revision) throw new Error("Metadata preparation does not match its job.");
  const manifest = JSON.parse(job.manifest_json);
  if (job.next_batch !== manifest.batches.length || job.prepared_batch !== manifest.batches.length || job.event_count !== manifest.eventCount || job.issue_count !== manifest.issueCount) throw new Error("Incomplete staged import cannot be prepared or activated.");
  if (job.compact_ready) return { duplicate: true };
  if (!await reserveRosterMaintenanceBudget(db, 750 * COST.compact + COST.control, 70000)) return { deferred: true };
  const result = await refreshFacilityOverviewMaterializationForFile(db, job.file_id, {
    maximumEventRows: 25000, maximumDoctorRows: 512, maximumExistingStaffRows: 750,
    maximumExistingCatalogRows: 750, maximumWrites: 750, contentRevision: revision,
  });
  if (result.overBudget || result.ok === false) throw new Error(`Compact metadata preparation blocked: ${result.reason}`);
  if (result.eventCount !== manifest.eventCount || result.doctorCount !== manifest.doctors.length) throw new Error("Persisted roster counts do not match the complete staged manifest.");
  await db.prepare("UPDATE roster_import_jobs SET compact_ready = 1 WHERE run_id = ? AND plan_revision = ?").bind(runId, revision).run();
  return result;
}

// Pinned jobs use atomic replacement promotion. Legacy jobs retain the disjoint-term gate.
export async function activateBoundedRosterTerm(db, runId, revision, options = {}) {
  const job = await db.prepare("SELECT * FROM roster_import_jobs WHERE run_id = ?").bind(runId).first();
  if (!job || job.plan_revision !== revision || !job.compact_ready) throw new Error("Staged roster is not ready for activation.");
  if (job.activated) return { duplicate: true, completed: true, fileId: job.file_id, retiredFileIds: JSON.parse(job.retired_file_ids_json || "[]") };
  if (job.promotion_fence_json) return activatePreparedRosterReplacement(db, job, revision, options);
  const coverage = await db.prepare("SELECT * FROM roster_file_coverage WHERE file_id = ?").bind(job.file_id).first();
  const file = await db.prepare("SELECT source_type, active FROM roster_files WHERE id = ?").bind(job.file_id).first();
  if (!coverage || !file || Number(file.active) === 1) throw new Error("Prepared inactive roster is unavailable.");
  const manifest = JSON.parse(job.manifest_json);
  if (!await reserveRosterMaintenanceBudget(db, manifest.doctors.length * COST.sms + COST.control, 100000)) return { deferred: true };
  const ownership = await rosterTermOwnership(db, job.file_id);
  const termStart = ownership.firstTerm;
  const termEnd = ownership.lastTerm;
  const existing = await db.prepare(`SELECT f.id, c.coverage_start, c.coverage_end FROM roster_files f INDEXED BY idx_roster_files_source_active
    LEFT JOIN roster_file_coverage c ON c.file_id = f.id WHERE f.source_type = ? AND f.active = 1 LIMIT 33`).bind(file.source_type).all();
  if (existing.results.length > 32) throw new Error("Bounded activation requires at most 32 active files.");
  for (const row of existing.results) {
    if (!row.coverage_start || !row.coverage_end) throw new Error("Bounded activation requires prepared existing coverage.");
    const old = await rosterTermOwnership(db, row.id);
    if (old.firstDate <= ownership.lastDate && old.lastDate >= ownership.firstDate) throw new Error("Bounded activation requires a disjoint new term.");
  }
  const now = new Date().toISOString();
  // Only small control rows change here. Events and presence are already
  // prepared; active readers switch to the complete file atomically.
  const activation = await db.batch([
    db.prepare(`UPDATE roster_files SET active = 1 WHERE id = ? AND active = 0 AND NOT EXISTS (
      SELECT 1 FROM roster_files f INDEXED BY idx_roster_files_source_active
      LEFT JOIN roster_file_coverage c ON c.file_id = f.id
      WHERE f.source_type = ? AND f.active = 1 AND (c.file_id IS NULL OR NOT EXISTS (SELECT 1 FROM facility_stream_catalog_contributions own WHERE own.file_id = f.id)
        OR EXISTS (SELECT 1 FROM facility_stream_catalog_contributions own WHERE own.file_id = f.id GROUP BY own.file_id HAVING MIN(own.first_date) <= ? AND MAX(own.last_date) >= ?)))`)
      .bind(job.file_id, file.source_type, ownership.lastDate, ownership.firstDate),
    db.prepare("UPDATE roster_file_status_summaries SET active = 1, derived_state = 'ready', content_revision = ?, status_revision = ?, updated_at = ? WHERE file_id = ? AND EXISTS (SELECT 1 FROM roster_files f WHERE f.id = file_id AND f.active = 1)")
      .bind(revision, crypto.randomUUID(), now, job.file_id),
    db.prepare("UPDATE roster_import_jobs SET activated = 1 WHERE run_id = ? AND plan_revision = ? AND EXISTS (SELECT 1 FROM roster_files f WHERE f.id = file_id AND f.active = 1)").bind(runId, revision),
    facilitySmsMembershipStatement(db, job.file_id, true),
    ...(options.publishFacility ? facilityRefreshStatements(db, file.source_type, facilityTermDates(coverage.coverage_start, coverage.coverage_end), `${job.file_id}:${revision}`) : []),
  ]);
  if (Number(activation[0]?.meta?.changes || 0) !== 1) throw new Error("New-term activation was superseded by a conflicting roster.");
  return { fileId: job.file_id, events: job.event_count };
}

// All sources share the admitted account budget; each request reserves its worst-case cost.
async function reserveImportWrites(db, writes) {
  return reserveRosterMaintenanceBudget(db, writes, writes + 128);
}

async function activatePreparedRosterReplacement(db, job, revision, options) {
  const manifest = JSON.parse(job.manifest_json);
  const file = await db.prepare("SELECT source_type,active FROM roster_files WHERE id=?").bind(job.file_id).first();
  const coverage = await db.prepare("SELECT coverage_start,coverage_end FROM roster_file_coverage WHERE file_id=?").bind(job.file_id).first();
  if (!file || file.active || !coverage) throw new Error("Prepared inactive replacement is unavailable.");
  if (!await reserveRosterMaintenanceBudget(db, manifest.doctors.length * COST.sms + 1024, 150000)) return { deferred: true };
  const fence = JSON.parse(job.promotion_fence_json);
  const active = (await db.prepare("SELECT id,parsed_at FROM roster_files INDEXED BY idx_roster_files_source_active WHERE source_type=? AND active=1 ORDER BY id LIMIT 33").bind(file.source_type).all()).results;
  if (JSON.stringify(active) !== JSON.stringify(fence)) throw new Error("Active rosters changed during import; replacement was not activated. Start a fresh import after reviewing the current roster.");
  const incoming = await rosterTermOwnership(db, job.file_id);
  const retire = [];
  let start = coverage.coverage_start, end = coverage.coverage_end;
  for (const old of active) {
    const own = await rosterTermOwnership(db, old.id);
    if (own.firstDate > incoming.lastDate || own.lastDate < incoming.firstDate) continue;
    if (own.firstTerm !== incoming.firstTerm || own.lastTerm !== incoming.lastTerm || own.firstDate < incoming.firstDate || own.lastDate > incoming.lastDate) throw new Error("Replacement would remove dates outside the incoming roster. Import the full retained term range.");
    const oldCoverage = await db.prepare("SELECT coverage_start,coverage_end FROM roster_file_coverage WHERE file_id=?").bind(old.id).first();
    start = start < oldCoverage.coverage_start ? start : oldCoverage.coverage_start;
    end = end > oldCoverage.coverage_end ? end : oldCoverage.coverage_end;
    retire.push(old.id);
  }
  const dates = facilityTermDates(start, end), token = crypto.randomUUID(), now = new Date().toISOString();
  const guard = { runId: job.run_id, token };
  const gate = "EXISTS(SELECT 1 FROM roster_import_jobs WHERE run_id=? AND activation_token=?)";
  const statements = [
    db.prepare(`UPDATE roster_import_jobs SET activation_token=? WHERE run_id=? AND plan_revision=? AND activated=0 AND compact_ready=1
      AND EXISTS(SELECT 1 FROM roster_files WHERE id=? AND active=0)
      AND (SELECT COUNT(*) FROM roster_files INDEXED BY idx_roster_files_source_active WHERE source_type=? AND active=1)=json_array_length(promotion_fence_json)
      AND NOT EXISTS(SELECT 1 FROM roster_files f INDEXED BY idx_roster_files_source_active WHERE f.source_type=? AND f.active=1
        AND NOT EXISTS(SELECT 1 FROM json_each(promotion_fence_json) pin WHERE json_extract(pin.value,'$.id')=f.id AND json_extract(pin.value,'$.parsed_at')=f.parsed_at))`)
      .bind(token,job.run_id,revision,job.file_id,file.source_type,file.source_type),
    db.prepare(`UPDATE roster_files SET active=1 WHERE id=? AND active=0 AND ${gate}`).bind(job.file_id,job.run_id,token),
    ...retire.flatMap(id => [
      db.prepare(`UPDATE roster_files SET active=0 WHERE id=? AND active=1 AND ${gate}`).bind(id,job.run_id,token),
      db.prepare(`UPDATE roster_file_status_summaries SET active=0,status_revision=?,updated_at=? WHERE file_id=? AND ${gate}`).bind(crypto.randomUUID(),now,id,job.run_id,token),
      db.prepare(`UPDATE roster_sources SET active_file_id=? WHERE active_file_id=? AND ${gate}`).bind(job.file_id,id,job.run_id,token),
    ]),
    db.prepare(`UPDATE roster_file_status_summaries SET active=1,derived_state='ready',content_revision=?,status_revision=?,updated_at=? WHERE file_id=? AND ${gate}`).bind(revision,crypto.randomUUID(),now,job.file_id,job.run_id,token),
    db.prepare(`UPDATE roster_import_jobs SET activated=1,retired_file_ids_json=? WHERE run_id=? AND activation_token=?`).bind(JSON.stringify(retire),job.run_id,token),
    facilitySmsMembershipStatement(db,job.file_id,true),
    ...(options.publishFacility ? facilityRefreshStatements(db,file.source_type,dates,`${job.file_id}:${revision}`,guard) : []),
  ];
  const results = await db.batch(statements);
  if (Number(results[0]?.meta?.changes || 0) !== 1) throw new Error("Concurrent roster change prevented replacement activation.");
  return { fileId:job.file_id, retiredFileIds:retire, activated:true };
}
