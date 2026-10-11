import { buildFacilityOverviewFacts, collectDerivedEventAndIssueRows, sanitizeFileDoctors } from '../functions/_lib/d1-calendar.js';
import { rosterImportDigest } from '../functions/_lib/roster-import-batches.js';

export async function attachRosterCompactPlan(plan, payload) {
  const sourceType = payload.file.sourceType;
  const doctors = sanitizeFileDoctors(payload.doctors, sourceType);
  const {eventRows} = collectDerivedEventAndIssueRows(payload.file,sourceType,doctors,payload.eventsByDoctor,payload.issuesByDoctor);
  if(eventRows.length!==plan.manifest.eventCount) throw Error('Compact plan differs from normalized staged event counts.');
  const facts = await buildFacilityOverviewFacts(payload.file,
    doctors.map(d=>({doctor_key:d.key,display_name:d.displayName,seniority:d.seniority,membership_source:d.membershipSource,provider_staff_id:d.providerStaffId})),
    eventRows.map(r=>({id:r[0],doctor_key:r[3],display_name:r[4],start_date:r[5],end_date:r[6],seniority:r[11],event_json:r[16]})));
  const entries=[...facts.contributions.values()].map(fact=>({kind:'staff',fact}))
    .concat([...facts.catalog.values()].map(fact=>({kind:'catalog',fact})));
  if(entries.length>750) throw Error('Compact plan exceeds its retained metadata bound.');
  const compactBatches=[];
  for(let start=0;start<entries.length;start+=16) compactBatches.push({index:compactBatches.length,entries:entries.slice(start,start+16)});
  plan.manifest.compact={
    coverageStart:facts.coverageStart,coverageEnd:facts.coverageEnd,staffDigest:facts.staffDigest,dailyDigest:facts.dailyDigest,
    batches:await Promise.all(compactBatches.map(async batch=>({index:batch.index,hash:await rosterImportDigest(batch),count:batch.entries.length}))),
  };
  return {...plan,compactBatches,revision:await rosterImportDigest(plan.manifest)};
}
