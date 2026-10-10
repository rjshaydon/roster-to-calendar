import {
  claimRosterDispatch,
  updateRosterDispatch,
  listQueuedRosterSyncRuns,
} from "./d1-calendar.js";
import { guardedFetch, localOnlyEnabled } from "./outbound-network.js";
import { automatedRosterQueueEnabled, automatedRosterSourceEnabled } from "./roster-automation-guard.js";
import { refreshAccountMaintenanceBudget } from '../api/automation/account-budget.js';
import {markRosterDeliveryDeferred,clearRosterDeliveryWarning} from './roster-delivery-health.js';

const GITHUB_WORKFLOW = "monash-roster-sync.yml";
const GITHUB_REPOSITORY = "rjshaydon/roster-to-calendar";
const PRODUCTION_AUTOMATION_BASE_URL = "https://roster-to-calendar.pages.dev";
const REQUEST_LEASE_MS = 2 * 60 * 1000;
const ACCEPTED_LEASE_MS = 20 * 60 * 1000;
const TRANSIENT_RETRY_MS = 5 * 60 * 1000;
const AUTH_RETRY_MS = 60 * 60 * 1000;

export async function requestQueuedRosterProcessing(env, { sourceId = "", reason = "source-update", now = new Date() } = {}) {
  const normalizedSourceId = String(sourceId || "").trim();
  if (!automatedRosterQueueEnabled(env) || !automatedRosterSourceEnabled(env, normalizedSourceId)) {
    return { ok: false, dispatched: false, reason: "source-disabled", dispatch: null };
  }
  if (localOnlyEnabled(env)) {
    return { ok: false, dispatched: false, reason: "local-disabled", dispatch: null };
  }
  const requestedAt = validDate(now);
  if(env.ROSTER_ACCOUNT_BUDGET_ENABLED==='true') {
    let gate=await env.ROSTER_DB.prepare(`SELECT allocated_reads,allocated_writes,maximum_reads,maximum_writes,valid_until,stop_reason
      FROM roster_account_budget WHERE utc_day=?`).bind(requestedAt.toISOString().slice(0,10)).first();
    if(!gate || gate.valid_until<=requestedAt.toISOString()) {
      const refreshed=await (await refreshAccountMaintenanceBudget({env})).json();
      if(refreshed.deferred) return budgetDeferred(env,normalizedSourceId);
      gate=await env.ROSTER_DB.prepare('SELECT allocated_reads,allocated_writes,maximum_reads,maximum_writes,valid_until,stop_reason FROM roster_account_budget WHERE utc_day=?')
        .bind(requestedAt.toISOString().slice(0,10)).first();
    }
    if(!gate || gate.stop_reason || gate.valid_until<=requestedAt.toISOString() || gate.allocated_reads>=gate.maximum_reads || gate.allocated_writes>=gate.maximum_writes)
      return budgetDeferred(env,normalizedSourceId);
  }
  const claim = await claimRosterDispatch(env?.ROSTER_DB, {
    sourceId: normalizedSourceId,
    reason,
    now: requestedAt.toISOString(),
    retryAfter: addMilliseconds(requestedAt, REQUEST_LEASE_MS).toISOString(),
  });
  if (!claim.claimed) return { ok: true, dispatched: false, reason: claim.reason, dispatch: claim.dispatch || null };

  const token = String(env?.GITHUB_ACTIONS_TOKEN || "").trim();
  if (!token) {
    const dispatch = await failDispatch(env, claim.dispatch, "GitHub dispatch is not configured.", AUTH_RETRY_MS, requestedAt);
    return { ok: false, dispatched: false, reason: "github-token-missing", dispatch };
  }

  try {
    const automationBaseUrl = String(env?.ROSTER_AUTOMATION_BASE_URL || PRODUCTION_AUTOMATION_BASE_URL).replace(/\/$/, "");
    const workflowTarget = String(env?.ROSTER_AUTOMATION_WORKFLOW_TARGET || "production").trim().toLowerCase() === "preview" ? "preview" : "production";
    const workflowRef = String(env?.ROSTER_GITHUB_WORKFLOW_REF || "main").trim() || "main";
    const response = await guardedFetch(env, `https://api.github.com/repos/${GITHUB_REPOSITORY}/actions/workflows/${GITHUB_WORKFLOW}/dispatches`, {
      method: "POST",
      headers: {
        Accept: "application/vnd.github+json",
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "roster-to-calendar-dispatcher",
      },
      body: JSON.stringify({
        ref: workflowRef,
        inputs: {
          dispatch_id: claim.dispatch.id,
          source_id: normalizedSourceId,
          automation_base_url: automationBaseUrl,
          target: workflowTarget,
        },
      }),
    }, { label: "GitHub workflow dispatch" });
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
  const sourceId = String(body?.sourceId || "").trim();
  const event = String(body?.event || "").trim().toLowerCase();
  if (!automatedRosterQueueEnabled(env) || !automatedRosterSourceEnabled(env, sourceId)) {
    return { ok: false, reason: "source-disabled" };
  }
  if (!dispatchId || !dispatchId.startsWith(`dispatch:${sourceId}:`) || !["started", "completed", "failed", "deferred"].includes(event)) {
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
  const result=await updateRosterDispatch(env?.ROSTER_DB, dispatchId, {
    status: event === "completed" ? "completed" : event==='deferred'?'deferred':'failed',
    githubRunId: String(body?.githubRunId || "").trim(),
    completedAt: now.toISOString(),
    retryAfter: event==='deferred'?addMilliseconds(now,TRANSIENT_RETRY_MS).toISOString():now.toISOString(),
    lastError: event === "failed" ? String(body?.message || "GitHub roster processor failed.").slice(0, 300) : event==='deferred'?'Account maintenance allowance deferred; queued roster retained.':'',
  });
  if(event==='completed' && result.ok && !(await listQueuedRosterSyncRuns(env.ROSTER_DB,sourceId,1)).length) await clearRosterDeliveryWarning(env.ROSTER_FILES,sourceId);
  if(event==='deferred') await budgetDeferred(env,sourceId);
  return result;
}

async function budgetDeferred(env,sourceId) {
  const pending=(await listQueuedRosterSyncRuns(env.ROSTER_DB,sourceId,1))[0];
  if(pending) await markRosterDeliveryDeferred(env.ROSTER_FILES,sourceId,pending.startedAt);
  return {ok:true,dispatched:false,deferred:true,reason:'account-budget-deferred',dispatch:null};
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
