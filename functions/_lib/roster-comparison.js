import { loadPublishedDoctorSourceRange, doctorSourceKey } from './doctor-source-artifacts.js';
import { loadCachedSnapshot } from './d1-calendar.js';
import { normalizeRosterName } from './roster.js';
import { filterFacilityRowsBySegments, validFacilityDateRange, nextFacilityDate } from '../../public/static/facility-access-policy.js';
import { comparisonDayMask } from '../../public/static/comparison-policy.js';

export const COMPARISON_INDEX_VERSION = 'roster-comparison-index-v2';
const hash = async value => [...new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(JSON.stringify(value))))].map(v => v.toString(16).padStart(2, '0')).join('');
const uniqueKeys = values => [...new Set((values || []).map(normalizeRosterName).filter(Boolean))].sort();
export async function loadRosterComparison(r2, options, today, isWorking) {
  const { startDate, endDate, sourceTypes, authorisedSegments, scopeRevision, cachedRevision = '' } = options;
  const keys = uniqueKeys(options.doctorKeys);
  if (!validFacilityDateRange(startDate, endDate) || !keys.length || keys.length > 3) return { ok: false, invalid: true, events: [] };
  const identities = options.doctorIdentities;
  if (identities && (!Array.isArray(identities) || identities.length > 15 || identities.some(identity => !identity || !keys.includes(normalizeRosterName(identity.key)) || !['mmc', 'mch', 'ddh', 'vhh', 'casey'].includes(identity.sourceType)))) return { ok: false, invalid: true, events: [] };
  const markers = new Set(identities?.length ? identities.filter(identity => sourceTypes.includes(identity.sourceType)).map(identity => `${identity.sourceType}|${normalizeRosterName(identity.key)}`) : sourceTypes.flatMap(source => keys.map(key => `${source}|${key}`)));
  const prepared = await loadPublishedDoctorSourceRange(r2, sourceTypes, startDate, endDate, today, new Set([...markers].sort()), { includeCarryIn: true });
  const revision = await hash([COMPARISON_INDEX_VERSION, scopeRevision, startDate, endDate, keys, prepared.revision]);
  if (prepared.preparing) return { ok: false, unavailable: true, events: [], missing: prepared.missing };
  const personRevisions = Object.fromEntries(await Promise.all(keys.map(async key => [key, await hash([COMPARISON_INDEX_VERSION, scopeRevision, startDate, endDate, key, prepared.personFacts[key] || [], prepared.missing])])));
  const updatedDoctorKeys = keys.filter(key => options.cachedPersonRevisions?.[key] !== personRevisions[key]);
  if (cachedRevision === revision || !updatedDoctorKeys.length) return { ok:true, unchanged:true, revision, scopeRevision, personRevisions, missing:prepared.missing, incomplete:!!prepared.missing.length };
  const rows = filterFacilityRowsBySegments(await prepared.readEvents({doctorKeys:new Set(updatedDoctorKeys)}), authorisedSegments).filter(row => isWorking(row.event, row.sourceType));
  const events = [...new Map(rows.map(row => [`${row.sourceType}|${row.doctorKey}|${row.event.id || ''}|${row.event.start}|${row.event.end}`, row])).values()];
  return { ok: true, events, revision, scopeRevision, personRevisions, incomplete:!!prepared.missing.length, ...(updatedDoctorKeys.length < keys.length ? {delta:true,updatedDoctorKeys} : {}), missing: prepared.missing, source: 'prepared-doctor-comparison' };
}

export async function loadRosterComparisonDoctors(r2, options, today) {
  const { startDate, endDate, sourceTypes, authorisedSegments, scopeRevision, cachedRevision = '' } = options;
  const mine = new Set(uniqueKeys(options.overlapDoctorKeys)), exclude = new Set(uniqueKeys(options.excludeDoctorKeys));
  if (!validFacilityDateRange(startDate, endDate, 180) || mine.size > 64 || exclude.size > 64) return { ok: false, invalid: true, doctors: [] };
  if (!mine.size) return { ok: true, doctors: [], scopeRevision, revision: await hash([scopeRevision, startDate, endDate, []]) };
  const prepared = await loadPublishedDoctorSourceRange(r2, sourceTypes, startDate, endDate, today, new Set(), {includeCarryIn:true});
  if (prepared.preparing) return { ok: false, unavailable: true, doctors: [], missing: prepared.missing };
  const selected = [], missing = [...prepared.missing];
  for (const sourceType of sourceTypes) {
    const indexes = prepared.comparisonIndexes[sourceType];
    if (!indexes || prepared.comparisonMonths[sourceType].some(month => !indexes[month])) return { ok: false, unavailable: true, doctors: [] };
    const masks = new Map();
    for (const segment of authorisedSegments.filter(segment => segment.sourceType === sourceType)) {
      for(const term of prepared.visibleTerms.filter(term=>term.sourceType===sourceType)){
        const start=segment.startDate>term.termStart?segment.startDate:term.termStart;
        const last=nextFacilityDate(term.termEnd),end=segment.endDate<last?segment.endDate:last;
        if(start>end)continue;
        for (const month of Object.keys(indexes)) {
          if(month<start.slice(0,7)||month>end.slice(0,7))continue;
          if(!masks.has(month))masks.set(month,new Map());
          const byTerm=masks.get(month);byTerm.set(term.termStart,((byTerm.get(term.termStart)||0)|comparisonDayMask(month,start,end))>>>0);
        }
      }
    }
    for (const [month, byTerm] of masks) {
      const pointer = indexes[month];
      if (!pointer || pointer.key !== doctorSourceKey(sourceType, 'catalogues', pointer.digest)) return { ok: false, unavailable: true, doctors: [] };
      selected.push({ sourceType, month, byTerm:[...byTerm], pointer });
    }
  }
  const revision = await hash([COMPARISON_INDEX_VERSION, scopeRevision, startDate, endDate, [...mine].sort(), [...exclude].sort(), selected.map(({ sourceType, month, byTerm, pointer }) => [sourceType, month, byTerm, pointer.digest]), missing]);
  if (revision === cachedRevision) return { ok: true, unchanged: true, revision, scopeRevision, missing, incomplete:!!missing.length };
  const doctors = new Map();
  for (const { sourceType, month, byTerm, pointer } of selected) {
    const data = await loadCachedSnapshot(r2, pointer.key);
    if (data?.version !== COMPARISON_INDEX_VERSION || data.sourceType !== sourceType || data.month !== month || !Array.isArray(data.doctors) || data.doctors.length > 4096 || data.doctors.some(row=>!Array.isArray(row)||row.length!==4||typeof row[0]!=='string'||typeof row[1]!=='string'||!row[3]||typeof row[3]!=='object')) return { ok: false, unavailable: true, doctors: [] };
    const eligibleDates = row => byTerm.reduce((dates,[term,mask])=>(dates|((row[3]?.[term]||0)&mask))>>>0,0);
    let presence = 0;
    for (const row of data.doctors) if (mine.has(row[0])) presence = (presence | eligibleDates(row)) >>> 0;
    for (const row of data.doctors) {const [doctorKey, displayName]=row;if (!exclude.has(doctorKey) && (eligibleDates(row) & presence)) doctors.set(`${sourceType}|${doctorKey}`, { doctorKey, displayName, sourceType });}
  }
  return { ok: true, doctors: [...doctors.values()].sort((a, b) => a.displayName.localeCompare(b.displayName) || a.sourceType.localeCompare(b.sourceType)), revision, scopeRevision, missing, source: 'prepared-doctor-comparison' };
}
