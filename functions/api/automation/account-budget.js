import inventory from "../../_lib/d1-account-inventory.js";
import { d1AnalyticsQuery, settledUtcDayInterval, summarizeAnalyticsPayload } from "../../../scripts/d1-quota-budget-lib.mjs";
import { accountMaintenanceHeadroom, outstandingMaintenance } from "../../_lib/account-maintenance-policy.js";
import { stopAccountMaintenance } from "../../_lib/roster-maintenance-budget.js";

export async function onRequestPost(context) {
  const token = String(context.env.ROSTER_AUTOMATION_TOKEN || "");
  const provided = String(context.request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  if (!token || token.length !== provided.length || [...token].reduce((n,c,i) => n | (c.charCodeAt(0) ^ provided.charCodeAt(i)), 0)) return Response.json({ error: "Unauthorized." }, { status: 401 });
  if (context.env.ROSTER_ACCOUNT_BUDGET_ENABLED !== "true") return Response.json({ ok: true, legacy: true });
  const db = context.env.ROSTER_DB;
  const day = new Date().toISOString().slice(0, 10);
  try {
    const analyticsToken = String(context.env.ROSTER_ACCOUNT_ANALYTICS_TOKEN || "");
    if (!analyticsToken || !inventory.complete) throw new Error("Account analytics credentials or inventory unavailable.");
    const interval = settledUtcDayInterval();
    if (!interval.settled) throw new Error("Waiting for current UTC-day analytics settlement.");
    const response = await fetch("https://api.cloudflare.com/client/v4/graphql", {
      method: "POST", headers: { Authorization: `Bearer ${analyticsToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({ query: d1AnalyticsQuery(), variables: { accountTag: inventory.accountId, start: interval.start, end: interval.observedUntil } }),
      signal: AbortSignal.timeout(20000),
    });
    if (!response.ok) throw new Error("Account analytics request failed.");
    const analytics = { ...summarizeAnalyticsPayload(await response.json(), inventory), interval };
    if (!analytics.complete) throw new Error("Account analytics did not reconcile.");
    const receipts = (await db.prepare("SELECT reserved_reads,reserved_writes,actual_reads,actual_writes,finished_at,metadata_complete FROM roster_maintenance_receipts WHERE utc_day=? LIMIT 10001").bind(day).all()).results || [];
    if (receipts.length > 10000) throw new Error("Maintenance receipt inspection bound reached.");
    const outstanding = outstandingMaintenance(receipts, interval.observedUntil);
    const settledMaintenance = receipts.reduce((sum, row) => {
      if (Number(row.metadata_complete) === 1 && row.finished_at && row.finished_at <= interval.observedUntil) {
        // Stored actual costs include 24 estimated settlement units; subtract
        // only the measured route portion, never the estimate, from traffic.
        sum.reads += Math.max(0, Number(row.actual_reads)-24);
        sum.writes += Math.max(0, Number(row.actual_writes)-24);
      }
      return sum;
    }, { reads: 0, writes: 0 });
    const budget = accountMaintenanceHeadroom(analytics, outstanding, new Date(), settledMaintenance);
    // A concurrent reservation may occur after the SELECT above. Do not grant
    // headroom against a newer allocation baseline: fence it using this value.
    const previous = await db.prepare("SELECT allocated_reads,allocated_writes,stop_reason FROM roster_account_budget WHERE utc_day=?").bind(day).first();
    if (previous?.stop_reason && previous.stop_reason.startsWith("cost-overrun:")) return Response.json({ ok: true, deferred: true, reason: previous.stop_reason });
    const reads = Number(previous?.allocated_reads || 0), writes = Number(previous?.allocated_writes || 0);
    // Re-read receipts atomically by comparing their monotonic reserved totals.
    const receiptReads = receipts.reduce((n,r) => n + Number(r.reserved_reads), 0);
    const receiptWrites = receipts.reduce((n,r) => n + Number(r.reserved_writes), 0);
    const result = await db.prepare(`INSERT INTO roster_account_budget
      (utc_day,allocated_reads,allocated_writes,maximum_reads,maximum_writes,valid_until,stop_reason,observed_until,account_reads,account_writes)
      SELECT ?,?,?,?,?,?,'',?,?,? WHERE
        (SELECT COALESCE(SUM(reserved_reads),0) FROM roster_maintenance_receipts WHERE utc_day=?)=? AND
        (SELECT COALESCE(SUM(reserved_writes),0) FROM roster_maintenance_receipts WHERE utc_day=?)=?
      ON CONFLICT(utc_day) DO UPDATE SET maximum_reads=excluded.maximum_reads,maximum_writes=excluded.maximum_writes,
        valid_until=excluded.valid_until,stop_reason='',observed_until=excluded.observed_until,account_reads=excluded.account_reads,account_writes=excluded.account_writes
      WHERE allocated_reads=? AND allocated_writes=? AND stop_reason NOT LIKE 'cost-overrun:%'`)
      .bind(day, reads, writes, reads + budget.reads, writes + budget.writes, budget.validUntil, interval.observedUntil,
        analytics.totals.rowsRead, analytics.totals.rowsWritten, day, receiptReads, day, receiptWrites, reads, writes).run();
    return Response.json({ ok: true, deferred: Number(result.meta?.changes || 0) !== 1 || !budget.reads || !budget.writes,
      account: { reads: analytics.totals.rowsRead, writes: analytics.totals.rowsWritten }, outstanding, budget });
  } catch (error) {
    await stopAccountMaintenance(db, "analytics-unavailable-or-inconsistent");
    return Response.json({ ok: true, deferred: true, reason: String(error.message || error) });
  }
}
