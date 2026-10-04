import { automationSourceDefinition } from '../../_lib/automation-import.js';
import { findRosterSyncByProviderVersion, loadRawRosterFile, hasCalendarDb } from '../../_lib/d1-calendar.js';
import { automatedRosterSourceEnabled } from '../../_lib/roster-automation-guard.js';
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
    libraryId: 'Documents', path: '/Shared Documents/Medical/Rosters/Active Medical Roster.xlsx', fileName: 'Active Medical Roster.xlsx', termStart: current });
  return windows;
}

export async function onRequestPost(context) {
  const configured = String(context.env.ROSTER_AUTOMATION_TOKEN || '');
  if (!configured || context.request.headers.get('authorization') !== `Bearer ${configured}`) return new Response('Unauthorized', { status: 401 });
  if (context.env.ROSTER_METADATA_CHECK_ENABLED !== 'true') return new Response('Paused', { status: 503 });
  const body = await context.request.json().catch(() => ({}));
  if (body.mode === 'windows') return Response.json({ ok: true, windows: sharepointRosterWindows().filter(window => automatedRosterSourceEnabled(context.env, window.sourceId)) });
  const sourceId = String(body.sourceId || '');
  const fileName = String(body.fileName || '');
  const providerVersion = String(body.providerVersion || '');
  if (automationSourceDefinition(sourceId)?.provider !== 'sharepoint' || !automatedRosterSourceEnabled(context.env, sourceId)) return new Response('Source unavailable', { status: 403 });
  if (!hasCalendarDb(context.env) || !fileName || fileName.length > 180 || !providerVersion || providerVersion.length > 200) return new Response('Invalid metadata', { status: 400 });
  const run = await findRosterSyncByProviderVersion(context.env.ROSTER_DB, sourceId, providerVersion, fileName);
  if (run?.status === 'success') return Response.json({ ok: true, download: false, status: 'unchanged' });
  if (run && ['queued', 'processing'].includes(run.status)) {
    const retained = await loadRawRosterFile(context.env.ROSTER_DB, run.sourceFileId || run.fileId);
    if (retained?.objectKey) {
      // Exact retained input can resume without retrieving the provider again.
      await requestQueuedRosterProcessing(context.env, { sourceId, reason: 'metadata-resume' });
      return Response.json({ ok: true, download: false, status: 'resuming' });
    }
  }
  if (run?.status === 'failed') return Response.json({ ok: true, download: false, status: 'repair-required' });
  return Response.json({ ok: true, download: true, status: 'changed-or-new' });
}
