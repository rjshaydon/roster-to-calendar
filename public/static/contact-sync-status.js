export const CONTACT_SYNC_DELAY_MS = 15 * 60 * 1000;
export function contactSyncHealth(lastSuccessAt, now = new Date()) {
  const time = Date.parse(lastSuccessAt || '');
  return !Number.isFinite(time) ? 'unknown' : now.getTime() - time >= CONTACT_SYNC_DELAY_MS ? 'delayed' : 'healthy';
}
export function contactSyncWarning(list, {live = false, admin = false} = {}) {
  if (!live) return '';
  if (list?.reason === 'legacy-workbook') return 'Live contact sync received a full Excel workbook instead of the doctors-only JSON extract. Check the contact flow in Power Automate.';
  if (list?.syncHealth?.status === 'delayed' || contactSyncHealth(list?.syncHealth?.lastSuccessAt) === 'delayed') return 'Contact syncing has not reported success for at least 15 minutes. Phone allocations may be out of date.' + (admin ? ' Check the contact flow and its workbook name or access in Power Automate.' : '');
  if (!list || ['unavailable','not-current'].includes(list.status)) return 'Live phone allocations are unavailable.' + (list?.lastSourceDate ? ` The last worksheet received was dated ${list.lastSourceDate}.` : '') + (admin ? ' Check the contact flow and worksheet date in Power Automate.' : '');
  return '';
}
