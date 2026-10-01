import { createD1Meter } from "../../_middleware.js";
import { runFacilityPublicationStep } from "../../_lib/facility-overview-cache.js";
import { automaticFacilityPublicationEnabled, facilityRefreshStatements, facilityTermDates } from "../../_lib/facility-refresh-queue.js";
import { reserveRosterMaintenanceBudget, configureRosterMaintenanceBudget, stopAccountMaintenance } from "../../_lib/roster-maintenance-budget.js";
import { automatedRosterSourceEnabled, rosterWritePausedResponse } from "../../_lib/roster-automation-guard.js";

import { australianTermStartForDate, australianTermEndForStart, refreshFacilityOverviewMaterializationForFile, upsertRosterFileStatusSummaryStatement } from "../../_lib/d1-calendar.js";

const SOURCES = Object.freeze({ "monash-adults": "mmc", "monash-paeds": "mch", "dandenong-findmyshift": "ddh", "vhh-active-medical-roster": "vhh" });
const READ_RESERVATION = 125000;
const WRITE_RESERVATION = 64;

// One durable publication step per request. Source-scoped indexed polling,
// reservations before work, and cursor CAS prevent unbounded retry work.
export async function onRequestPost(context) {
  const token = String(context.env.ROSTER_AUTOMATION_TOKEN || "");
  const provided = String(context.request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!token || token.length !== provided.length || [...token].reduce((n, c, i) => n | (c.charCodeAt(0) ^ provided.charCodeAt(i)), 0)) return Response.json({ error: "Unauthorized." }, { status: 401 });
  const body = await context.request.json().catch(() => ({}));
  const source = SOURCES[body.sourceId];
  if (!source || !automatedRosterSourceEnabled(context.env, body.sourceId) || !automaticFacilityPublicationEnabled(context.env, source)) return rosterWritePausedResponse();
  const db = context.env.ROSTER_DB;
  configureRosterMaintenanceBudget(db, body.maintenanceBudget);
  if (["prepare-coverage", "queue-seed"].includes(body.mode)) {
    if (body.seedCurrent === true) {
      if (!await reserveRosterMaintenanceBudget(db, 8, 32)) return Response.json({ ok: true, deferred: true });
      const today = new Intl.DateTimeFormat("en-CA", { timeZone: "Australia/Melbourne", year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date());
      await db.prepare("INSERT OR IGNORE INTO facility_refresh_jobs (source_type,term_start,dates_json,content_signature,request_revision,status,updated_at) VALUES (?,?,'[]','seed-required',?,'seed-pending',?)")
        .bind(source, australianTermStartForDate(today), crypto.randomUUID(), new Date().toISOString()).run();
    }
    if (body.mode === "queue-seed") return Response.json({ ok: true, seedQueued: true });
    const files = await db.prepare(`SELECT f.*, c.coverage_start, c.coverage_end, s.derived_state
      FROM roster_files f INDEXED BY idx_roster_files_source_active
      LEFT JOIN roster_file_coverage c ON c.file_id = f.id LEFT JOIN roster_file_status_summaries s ON s.file_id = f.id
      WHERE f.source_type = ? AND f.active = 1 LIMIT 33`).bind(source).all();
    if (files.results.length > 32) return Response.json({ ok: false, reason: "active-file-limit" }, { status: 409 });
    const missing = files.results.find((file) => file.source_id === body.sourceId && (!file.coverage_start || !file.coverage_end || file.derived_state !== "ready"));
    if (!missing) {
      const seed = await db.prepare("SELECT term_start FROM facility_refresh_jobs WHERE source_type=? AND status='seed-pending' LIMIT 1").bind(source).first();
      if (seed) return onRequestPost({ ...context, request: new Request(context.request.url, { method: "POST", headers: context.request.headers, body: JSON.stringify({ ...body, mode: "seed", seedCurrent: true }) }) });
      return Response.json({ ok: true, idle: true });
    }
    const day = new Date().toISOString().slice(0, 10);
    if (!await reserveRosterMaintenanceBudget(db, 64, 100000)) return Response.json({ ok: true, deferred: true });
    // Every first() in this exact-file primitive is a PK/scalar query. Opt in
    // to metadata-bearing all() so measured unused reservations can be refunded.
    const meter = createD1Meter(db, 760, { firstViaAll: true });
    let reservedMutationWrites = 0;
    const result = await refreshFacilityOverviewMaterializationForFile(meter.binding, missing.id, {
      sourceType: source, maximumEventRows: 25000, maximumDoctorRows: 512,
      maximumExistingStaffRows: 750, maximumExistingCatalogRows: 750, maximumWrites: 750,
      reserveMutationWrites: async (count) => {
        const cost = count * 6 + 4;
        if (!await reserveRosterMaintenanceBudget(db, cost, 0)) return false;
        reservedMutationWrites = cost;
        return true;
      },
    });
    if (result.deferred) {
      if (meter.metadataComplete && meter.rowsRead <= 100000) await db.prepare("UPDATE roster_import_daily_budget SET reserved_reads=MAX(0,reserved_reads-?) WHERE utc_day=?").bind(100000-meter.rowsRead, day).run();
      return Response.json(result);
    }
    if (result.overBudget || result.ok === false) return Response.json(result, { status: 409 });
    if (!result.eventCount) return Response.json({ ok: false, reason: "retained-file-has-no-shifts" }, { status: 409 });
    await upsertRosterFileStatusSummaryStatement(meter.binding, {
      id: missing.id, sourceType: source, sourceId: missing.source_id, name: missing.name,
      active: true, derivedState: "ready", expectedDoctorCount: result.doctorCount,
      indexedDoctorCount: result.doctorCount, eventCount: result.eventCount,
      contentRevision: result.contentRevision, size: missing.size, lastModified: missing.last_modified,
    }).run();
    if (meter.rowsRead > 100000 || meter.rowsWritten > reservedMutationWrites + 64) {
      await db.prepare("UPDATE roster_import_daily_budget SET reserved_reads=500000, reserved_writes=10000 WHERE utc_day=?").bind(day).run();
      throw new Error("Coverage preparation exceeded its reservation; maintenance stopped.");
    }
    if (meter.metadataComplete) await db.prepare("UPDATE roster_import_daily_budget SET reserved_reads=MAX(0,reserved_reads-?), reserved_writes=MAX(0,reserved_writes-?) WHERE utc_day=?")
      .bind(100000-meter.rowsRead, Math.max(0,reservedMutationWrites-meter.rowsWritten), day).run();
    return Response.json({ ok: true, prepared: true, sourceType: source, eventCount: result.eventCount });
  }
  if (body.seedCurrent === true) {
    if (!await reserveRosterMaintenanceBudget(db, 64, 2048)) return Response.json({ ok: true, deferred: true });
    const today = new Intl.DateTimeFormat("en-CA", { timeZone: "Australia/Melbourne", year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date());
    const term = australianTermStartForDate(today);
    const end = australianTermEndForStart(term);
    const coverage = await db.prepare(`SELECT f.id, c.coverage_start, c.coverage_end, c.content_revision
      FROM roster_files f INDEXED BY idx_roster_files_source_active LEFT JOIN roster_file_coverage c ON c.file_id = f.id
      WHERE f.source_type = ? AND f.active = 1 LIMIT 33`).bind(source).all();
    if (coverage.results.length > 32 || coverage.results.some((row) => !row.coverage_start || !row.coverage_end)) return Response.json({ ok: false, reason: "prepared-active-coverage-required" }, { status: 409 });
    const current = coverage.results.filter((row) => row.coverage_start <= end && row.coverage_end >= term).sort((a, b) => a.id.localeCompare(b.id));
    const dates = [...new Set(current.flatMap((row) => facilityTermDates(row.coverage_start < term ? term : row.coverage_start, row.coverage_end > end ? end : row.coverage_end)))];
    if (dates.length) await db.batch(facilityRefreshStatements(db, source, dates, JSON.stringify(current.map((row) => [row.id, row.content_revision]))));
    return Response.json({ ok: true, seeded: dates.length > 0, sourceType: source, termStart: term, dateCount: dates.length });
  }
  const job = await db.prepare("SELECT * FROM facility_refresh_jobs WHERE source_type = ? AND status = 'pending' ORDER BY term_start LIMIT 1").bind(source).first();
  if (!job) return Response.json({ ok: true, idle: true });
  const day = new Date().toISOString().slice(0, 10);
  if (!await reserveRosterMaintenanceBudget(db, WRITE_RESERVATION, READ_RESERVATION)) return Response.json({ ok: true, deferred: true });
  const plan = job.plan_json ? JSON.parse(job.plan_json) : null;
  let mode = "plan";
  if (plan) mode = job.next_batch < plan.batchCount ? "build-batch" : job.next_month < plan.months.length ? "build-month" : "finalize";
  const meter = createD1Meter(db, 48);
  let result;
  try {
    result = await runFacilityPublicationStep({ ...context, env: { ...context.env, ROSTER_DB: meter.binding } }, source, {
      mode, termStart: job.term_start, dates: JSON.parse(job.dates_json),
      operationRevision: plan?.operationRevision, batchIndex: job.next_batch, month: plan?.months[job.next_month],
    });
    if (meter.metadataComplete && (meter.rowsRead > READ_RESERVATION || meter.rowsWritten > WRITE_RESERVATION)) {
      if (context.env.ROSTER_ACCOUNT_BUDGET_ENABLED === "true") await stopAccountMaintenance(db, "cost-overrun:publication");
      else await db.prepare("UPDATE roster_import_daily_budget SET reserved_reads = 500000, reserved_writes = 10000 WHERE utc_day = ?").bind(day).run();
      throw new Error("Publication exceeded its reserved cost; maintenance stopped for this UTC day.");
    }
    if (result.reason === "term-not-prepared") result = { ...result, ok: true, deferred: true, waitingForTerm: true };
    if (result.ok && !result.deferred) {
      await db.prepare(`UPDATE facility_refresh_jobs SET plan_json = ?, next_batch = ?, next_month = ?, status = ?, last_error = '', updated_at = ?
        WHERE source_type = ? AND term_start = ? AND request_revision = ? AND next_batch = ? AND next_month = ?`)
        .bind(mode === "plan" ? JSON.stringify(result) : job.plan_json,
          job.next_batch + (mode === "build-batch" ? 1 : 0), job.next_month + (mode === "build-month" ? 1 : 0),
          mode === "finalize" ? "complete" : "pending", new Date().toISOString(), source, job.term_start, job.request_revision, job.next_batch, job.next_month).run();
    } else if (result.stalePlan) {
      await db.prepare("UPDATE facility_refresh_jobs SET plan_json = '', next_batch = 0, next_month = 0, last_error = ? WHERE source_type = ? AND term_start = ? AND request_revision = ?")
        .bind(result.reason, source, job.term_start, job.request_revision).run();
    }
  } catch (error) {
    await db.prepare("UPDATE facility_refresh_jobs SET last_error = ? WHERE source_type = ? AND term_start = ? AND request_revision = ?")
      .bind(String(error.message || error).slice(0, 300), source, job.term_start, job.request_revision).run();
    throw error;
  } finally {
    // Refund unused reads only when every result supplied D1 billing metadata.
    // The reservation remains intact after errors or incomplete instrumentation.
    if (result?.ok && meter.metadataComplete && meter.rowsRead <= READ_RESERVATION && meter.rowsWritten <= WRITE_RESERVATION) {
      await db.prepare("UPDATE roster_import_daily_budget SET reserved_reads = MAX(0, reserved_reads - ?), reserved_writes=MAX(0,reserved_writes-?) WHERE utc_day = ?")
        .bind(READ_RESERVATION - meter.rowsRead, Math.max(0,WRITE_RESERVATION-16-meter.rowsWritten), day).run();
    }
  }
  return Response.json({ ...result, completed: result.ok && !result.deferred && mode === "finalize" }, { status: result.ok ? 200 : 409 });
}
