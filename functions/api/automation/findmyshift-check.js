import { findmyshiftConfiguredRosterRange, findmyshiftLastModified, findmyshiftRosterWorkbook } from "../../_lib/findmyshift.js";
import { hasCalendarDb, loadRosterSource, upsertRosterSource } from "../../_lib/d1-calendar.js";

const SOURCE_ID = "dandenong-findmyshift";

export async function onRequestPost(context) {
  if (!hasValidToken(context.request, context.env)) return Response.json({ error: "Unauthorized." }, { status: 401 });
  if (!hasCalendarDb(context.env)) return Response.json({ error: "Roster database is unavailable." }, { status: 503 });
  const apiKey = String(context.env.FINDMYSHIFT_API_KEY || "").trim();
  const teamId = String(context.env.FINDMYSHIFT_TEAM_ID || "").trim();
  if (!apiKey || !teamId) return Response.json({ ok: false, status: "not-configured", error: "FindMyShift API key or team ID is not configured." });

  const now = new Date().toISOString();
  const current = await loadRosterSource(context.env.ROSTER_DB, SOURCE_ID);
  try {
    const providerVersion = await findmyshiftLastModified(apiKey, teamId);
    // A provider version is current only after it has completed the whole
    // retained-source → background-parser → active-calendar lifecycle.  A
    // failed import must be retried even when FindMyShift has not changed the
    // roster since the failed attempt.
    if (current?.providerVersion
      && current.providerVersion === providerVersion
      && current.lastSuccessAt
      && !current.lastError) {
      await saveSource(context, current, { lastCheckedAt: now, lastError: "" });
      return Response.json({ ok: true, status: "unchanged", providerModifiedAt: providerVersion });
    }
    const range = findmyshiftConfiguredRosterRange(context.env);
    const workbook = await findmyshiftRosterWorkbook(apiKey, teamId, range);
    const response = await fetch(new URL("/api/automation/ingest", context.request.url), {
      method: "POST",
      headers: { Authorization: `Bearer ${String(context.env.ROSTER_AUTOMATION_TOKEN || "")}`, "Content-Type": "application/json" },
      body: JSON.stringify({
        sourceId: SOURCE_ID,
        fileName: `Dandenong-FindMyShift-${range.from}-to-${range.to}.xlsx`,
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        contentBase64: bytesToBase64(new Uint8Array(workbook)),
        providerVersion,
        providerModifiedAt: providerVersion,
      }),
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(String(result.error || `Roster queue returned HTTP ${response.status}.`));
    return Response.json({ ok: true, status: String(result.status || "queued"), providerModifiedAt: providerVersion, queue: { runId: String(result.runId || ""), dispatched: result.processorDispatch?.dispatched === true } });
  } catch (error) {
    await saveSource(context, current, { lastCheckedAt: now, lastError: String(error?.message || error).slice(0, 300) });
    return Response.json({ ok: false, status: "failed", error: "FindMyShift roster check failed." }, { status: 502 });
  }
}

async function saveSource(context, existing, update) {
  const now = new Date().toISOString();
  await upsertRosterSource(context.env.ROSTER_DB, {
    ...(existing || {}), id: SOURCE_ID, provider: "findmyshift", sourceType: "ddh", label: "Dandenong (Findmyshift)", enabled: true,
    config: existing?.config || {}, cursor: existing?.cursor || {}, createdAt: existing?.createdAt || now, updatedAt: now, ...update,
  });
}

function hasValidToken(request, env) {
  const provided = String(request.headers.get("authorization") || "").replace(/^Bearer\s+/i, "");
  return [env.ROSTER_AUTOMATION_TOKEN, env.ROSTER_WATCHDOG_TOKEN].some((candidate) => timingSafeEqual(provided, String(candidate || "")));
}

function timingSafeEqual(left, right) {
  if (!left || !right || left.length !== right.length) return false;
  let mismatch = 0;
  for (let index = 0; index < left.length; index += 1) mismatch |= left.charCodeAt(index) ^ right.charCodeAt(index);
  return mismatch === 0;
}

function bytesToBase64(bytes) {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}
