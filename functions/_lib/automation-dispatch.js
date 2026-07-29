import {
  claimRosterDispatch,
  loadLatestRosterDispatch,
  updateRosterDispatch,
} from "./d1-calendar.js";

const GITHUB_WORKFLOW = "monash-roster-sync.yml";
const GITHUB_REPOSITORY = "rjshaydon/roster-to-calendar";
const REQUEST_LEASE_MS = 2 * 60 * 1000;
const ACCEPTED_LEASE_MS = 20 * 60 * 1000;
const TRANSIENT_RETRY_MS = 5 * 60 * 1000;
const AUTH_RETRY_MS = 60 * 60 * 1000;

export async function requestQueuedRosterProcessing(env, { reason = "source-update", now = new Date() } = {}) {
  const requestedAt = validDate(now);
  const claim = await claimRosterDispatch(env?.ROSTER_DB, {
    reason,
    now: requestedAt.toISOString(),
    retryAfter: addMilliseconds(requestedAt, REQUEST_LEASE_MS).toISOString(),
  });
  if (!claim.claimed) return { ok: true, dispatched: false, reason: claim.reason, dispatch: claim.dispatch || await loadLatestRosterDispatch(env?.ROSTER_DB) };

  const token = String(env?.GITHUB_ACTIONS_TOKEN || "").trim();
  if (!token) {
    const dispatch = await failDispatch(env, claim.dispatch, "GitHub dispatch is not configured.", AUTH_RETRY_MS, requestedAt);
    return { ok: false, dispatched: false, reason: "github-token-missing", dispatch };
  }

  try {
    const response = await fetch(`https://api.github.com/repos/${GITHUB_REPOSITORY}/actions/workflows/${GITHUB_WORKFLOW}/dispatches`, {
      method: "POST",
      headers: {
        Accept: "application/vnd.github+json",
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "roster-to-calendar-dispatcher",
      },
      body: JSON.stringify({ ref: "main", inputs: { dispatch_id: claim.dispatch.id } }),
    });
    if (response.status === 204) {
      const acceptedAt = new Date().toISOString();
      const updated = await updateRosterDispatch(env.ROSTER_DB, claim.dispatch.id, {
        status: "accepted",
        acceptedAt,
        retryAfter: addMilliseconds(new Date(acceptedAt), ACCEPTED_LEASE_MS).toISOString(),
        lastError: "",
      });
      return { ok: true, dispatched: true, reason: "accepted", dispatch: updated.dispatch };
    }
    const detail = sanitiseGitHubError(await response.text(), response.status);
    const retryMs = [401, 403].includes(response.status) ? AUTH_RETRY_MS : TRANSIENT_RETRY_MS;
    const dispatch = await failDispatch(env, claim.dispatch, detail, retryMs, requestedAt);
    return { ok: false, dispatched: false, reason: "github-rejected", dispatch };
  } catch (error) {
    const dispatch = await failDispatch(env, claim.dispatch, `GitHub dispatch request failed: ${String(error?.message || error)}`, TRANSIENT_RETRY_MS, requestedAt);
    return { ok: false, dispatched: false, reason: "github-unreachable", dispatch };
  }
}

export async function recordRosterDispatchLifecycle(env, body = {}) {
  const dispatchId = String(body?.dispatchId || "").trim();
  const event = String(body?.event || "").trim().toLowerCase();
  if (!dispatchId || !["started", "completed", "failed"].includes(event)) {
    return { ok: false, reason: "invalid-lifecycle-event" };
  }
  const now = new Date();
  if (event === "started") {
    return updateRosterDispatch(env?.ROSTER_DB, dispatchId, {
      status: "running",
      githubRunId: String(body?.githubRunId || "").trim(),
      startedAt: now.toISOString(),
      retryAfter: addMilliseconds(now, ACCEPTED_LEASE_MS).toISOString(),
      lastError: "",
    });
  }
  return updateRosterDispatch(env?.ROSTER_DB, dispatchId, {
    status: event === "completed" ? "completed" : "failed",
    githubRunId: String(body?.githubRunId || "").trim(),
    completedAt: now.toISOString(),
    retryAfter: now.toISOString(),
    lastError: event === "failed" ? String(body?.message || "GitHub roster processor failed.").slice(0, 300) : "",
  });
}

async function failDispatch(env, dispatch, message, retryMs, now) {
  const updated = await updateRosterDispatch(env?.ROSTER_DB, dispatch.id, {
    status: "failed",
    completedAt: new Date().toISOString(),
    retryAfter: addMilliseconds(now, retryMs).toISOString(),
    lastError: String(message || "GitHub dispatch failed.").slice(0, 300),
  });
  return updated.dispatch;
}

function validDate(value) {
  const date = value instanceof Date ? value : new Date(value);
  return Number.isNaN(date.getTime()) ? new Date() : date;
}

function addMilliseconds(date, milliseconds) {
  return new Date(date.getTime() + milliseconds);
}

function sanitiseGitHubError(text, status) {
  let message = "";
  try {
    message = String(JSON.parse(text || "{}")?.message || "");
  } catch {
    message = String(text || "");
  }
  return `GitHub workflow dispatch returned HTTP ${status}${message ? `: ${message.slice(0, 180)}` : "."}`;
}
