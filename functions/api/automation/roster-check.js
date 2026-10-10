import {loadPendingRosterPublication} from '../../_lib/roster-delivery-health.js';
import { automationSourceDefinition } from '../../_lib/automation-import.js';
import { findRosterSyncByProviderVersion, loadRawRosterFile, hasCalendarDb } from '../../_lib/d1-calendar.js';
import { automatedRosterSourceEnabled } from '../../_lib/roster-automation-guard.js';
import { refreshAccountMaintenanceBudget } from './account-budget.js';
import { requestQueuedRosterProcessing } from '../../_lib/automation-dispatch.js';
import { australianTermStartForDate, australianTermEndForStart } from '../../_lib/d1-calendar.js';
import { melbourneDateKey, rosterTermAvailableFrom } from '../../../public/static/roster-term-policy.js';

export function sharepointRosterWindows(now = new Date()) {
  const today = melbourneDateKey(now);
  const current = australianTermStartForDate(today);
  const end = australianTermEndForStart(current);
  const following = new Date(`${end}T12:00:00Z`); following.setUTCDate(following.getUTCDate() + 1);
  const next = australianTermStartForDate(following.toISOString().slice(0, 10));
  const terms = [current, ...(rosterTermAvailableFrom(next) <= today ? [next] : [])];
  const dataset = 'https://monashhealth.sharepoint.com/sites/MonashEDMMC-CLA-PFU-DEP-MedicalRoster';
  const libraryId = '43f6a549-2e31-486a-8ca2-ecdbe383986a';
  const windows = terms.flatMap(termStart => {
    const [year, month] = termStart.split('-').map(Number);
    const term = Math.floor(month / 3) + 1;
    return [{ sourceId: 'monash-adults', dataset, libraryId, fileName: `AdultTerm${term}.${year}.xlsx`, termStart },
      { sourceId: 'monash-paeds', dataset, libraryId, fileName: `Paeds - Term ${term} ${year}.xlsx`, termStart }]
      .map(item => ({ ...item, path: `/Shared Documents/Medical Roster/${item.fileName}` }));
  });
  windows.push({ sourceId: 'vhh-active-medical-roster', dataset: 'https://monashhealth.sharepoint.com/sites/VHHED-VHH-EMG-DEP',
    libraryId: 'dd1e780c-6901-4004-a615-73e8299158f4', path: '/Shared Documents/Medical/Rosters/Active Medical Roster.xlsx', fileName: 'Active Medical Roster.xlsx', termStart: current });
  return windows;
}

export async function onRequestPost(context, checkTime = new Date()) {
  const configured = String(context.env.ROSTER_AUTOMATION_TOKEN || '');
  if (!configured || context.request.headers.get('authorization') !== `Bearer ${configured}`) return new Response('Unauthorized', { status: 401 });
  if (context.env.ROSTER_METADATA_CHECK_ENABLED !== 'true') return new Response('Paused', { status: 503 });
  const body = await context.request.json().catch(() => ({}));
  if (!body || typeof body !== 'object' || Array.isArray(body)) return new Response('Invalid metadata', { status: 400 });
  if (body.mode === 'windows') return Response.json({ ok: true, windows: sharepointRosterWindows(checkTime).filter(window => automatedRosterSourceEnabled(context.env, window.sourceId)) });
  if (body.mode === 'reconcile') {
    const windows = sharepointRosterWindows(checkTime).filter(window => automatedRosterSourceEnabled(context.env, window.sourceId));
    const groups = body.libraries;
    if (!Array.isArray(groups) || groups.length !== 2 || groups.some(group => !group || !['monash','vhh'].includes(group.site) || !Array.isArray(group.files) || group.files.length > 1000) || new Set(groups.map(group => group.site)).size !== 2) return new Response('Invalid or incomplete metadata batch', { status: 400 });
    const downloads = [], checks = [];
    for (const window of windows) {
      const group = groups.find(group => group.site === (window.sourceId === 'vhh-active-medical-roster' ? 'vhh' : 'monash'));
      if (group.unavailable || group.nextLink) { checks.push({sourceId:window.sourceId,fileName:window.fileName,status:group.unavailable?'provider-unavailable':'incomplete-inventory'}); continue; }
      const matches = group.files.filter(file => file?.FileRef === window.path || file?.FileRef === new URL(window.dataset).pathname + window.path);
      if (!matches.length) { checks.push({sourceId:window.sourceId,fileName:window.fileName,status:'waiting-for-file'}); continue; }
      if (matches.length !== 1) return new Response('Ambiguous metadata', { status: 400 });
      const file = matches[0];
      const providerVersion = String(window.sourceId === 'vhh-active-medical-roster' ? file.File?.ETag || '' : file.OData__UIVersionString || '');
      if (typeof file.File?.ETag !== 'string' || !file.File.ETag || typeof file.Modified !== 'string' || !providerVersion || providerVersion.length > 200 || !Number.isFinite(Date.parse(file.Modified))) return new Response('Invalid provider metadata', { status: 400 });
      const response = await checkRosterMetadata(context,{sourceId:window.sourceId,fileName:window.fileName,providerVersion},checkTime);
      if (!response.ok) return response;
      const result = await response.json();
      checks.push({sourceId:window.sourceId,fileName:window.fileName,status:result.status});
      if (result.download) downloads.push({...window,providerVersion,providerModifiedAt:file.Modified,etag:String(file.File?.ETag || '')});
    }
    return Response.json({ok:true,downloads,checks});
  }
  return checkRosterMetadata(context,body,checkTime);
}

async function checkRosterMetadata(context,body,checkTime) {
  const sourceId = String(body.sourceId || '');
  const fileName = String(body.fileName || '');
  const providerVersion = String(body.providerVersion || '');
  if (automationSourceDefinition(sourceId)?.provider !== 'sharepoint' || !automatedRosterSourceEnabled(context.env, sourceId)) return new Response('Source unavailable', { status: 403 });
  if (!sharepointRosterWindows(checkTime).some(window => window.sourceId === sourceId && window.fileName.toLowerCase() === fileName.toLowerCase())) return Response.json({ ok: true, download: false, status: 'outside-active-windows' });
  if (!hasCalendarDb(context.env) || !fileName || fileName.length > 180 || !providerVersion || providerVersion.length > 200) return new Response('Invalid metadata', { status: 400 });
  const run = await findRosterSyncByProviderVersion(context.env.ROSTER_DB, sourceId, providerVersion, fileName);
  if (run?.status === 'success') {
    if(await loadPendingRosterPublication(context.env,sourceId)) {
      const dispatch=await requestQueuedRosterProcessing(context.env,{sourceId,reason:'publication-resume'});
      return Response.json({ok:true,download:false,status:dispatch.deferred?'deferred':'awaiting-publication'});
    }
    return Response.json({ ok: true, download: false, status: 'unchanged' });
  }
  if (run && ['queued', 'processing'].includes(run.status)) {
    const retained = await loadRawRosterFile(context.env.ROSTER_DB, run.sourceFileId || run.fileId);
    if (retained?.objectKey) {
      // Exact retained input can resume without retrieving the provider again.
      const dispatch=await requestQueuedRosterProcessing(context.env, { sourceId, reason: 'metadata-resume' });
      return Response.json({ ok: true, download: false, status: dispatch.deferred?'deferred':'resuming' });
    }
  }
  if (run?.status === 'failed') return Response.json({ ok: true, download: false, status: 'repair-required' });
  // Poll the latest snapshot rather than replaying autosaves. This exact-file
  // check is read-only and also protects manual reconciliation between ticks.
  const recent = await context.env.ROSTER_DB.prepare(`
    SELECT r.started_at FROM roster_sync_runs r
    INNER JOIN raw_roster_files f ON f.file_id = COALESCE(NULLIF(r.source_file_id, ''), r.file_id)
    WHERE r.source_id = ? AND r.started_at > ? AND LOWER(f.name) = LOWER(?)
    ORDER BY r.started_at DESC LIMIT 1
  `).bind(sourceId, new Date(checkTime.getTime() - 5 * 60 * 1000).toISOString(), fileName).first();
  if (recent) return Response.json({ ok: true, download: false, status: 'settling', retryAfter: new Date(Date.parse(recent.started_at) + 5 * 60 * 1000).toISOString() });
  if (context.env.ROSTER_ACCOUNT_BUDGET_ENABLED === 'true') {
    const now = new Date().toISOString();
    const grant = await context.env.ROSTER_DB.prepare('SELECT valid_until,stop_reason FROM roster_account_budget WHERE utc_day=?').bind(now.slice(0,10)).first();
    if (grant?.stop_reason?.startsWith('cost-overrun:')) return Response.json({ ok: true, download: false, status: 'deferred', reason: grant.stop_reason });
    if (!grant || grant.valid_until <= now || grant.stop_reason) {
      const admission = await (await refreshAccountMaintenanceBudget(context)).json();
      if (admission.deferred) return Response.json({ ok: true, download: false, status: 'deferred', reason: admission.reason || 'account-budget' });
    }
  }
  return Response.json({ ok: true, download: true, status: 'changed-or-new' });
}
