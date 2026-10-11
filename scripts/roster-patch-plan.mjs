import { createHash } from 'node:crypto';
import { sanitizeFileDoctors,collectDerivedEventAndIssueRows,buildFacilityOverviewFacts,expandEventToDailyPresenceRows } from '../functions/_lib/d1-calendar.js';
import { rosterImportDigest } from '../functions/_lib/roster-import-batches.js';
const hash=value=>createHash('sha256').update(JSON.stringify(value)).digest('hex');
const wire=value=>JSON.parse(JSON.stringify(value));
const eventColumns='id,doctor_key,display_name,start_date,end_date,start_ts,end_ts,title,raw_value,seniority,provider_staff_id,location,all_day,time_label,event_json'.split(',');
const issueColumns='id,doctor_key,display_name,start_date,raw_value,seniority,status,message,resolution_type,suggested_title,time_label,issue_json'.split(',');
export async function planRosterPatch(payload,baseline,stored) {
  const file={...payload.file,id:baseline.fileId},sourceType=file.sourceType,now=baseline.baselineParsedAt;
  const doctors=sanitizeFileDoctors(payload.doctors,sourceType);
  const {eventRows,issueRows}=collectDerivedEventAndIssueRows(file,sourceType,doctors,payload.eventsByDoctor,payload.issuesByDoctor);
  const doctorRows=doctors.map(d=>[file.id,sourceType,d.key,d.displayName,d.seniority,d.membershipSource,d.providerStaffId]);
  const contentRevision=hash({sourceType,doctors:doctors.map(d=>[d.key,d.displayName,d.seniority,d.membershipSource,d.providerStaffId]),events:eventRows.map(r=>[r[0],...r.slice(3)]),issues:issueRows.map(r=>[r[0],...r.slice(3)])});
  const compact=await buildFacilityOverviewFacts(file,doctors.map(d=>({doctor_key:d.key,display_name:d.displayName,seniority:d.seniority,membership_source:d.membershipSource,provider_staff_id:d.providerStaffId})),
    eventRows.map(r=>({id:r[0],doctor_key:r[3],display_name:r[4],start_date:r[5],end_date:r[6],seniority:r[11],event_json:r[16]})),contentRevision);
  const entries=[],dates=new Set(),changedDoctors=new Set();let changedFacts=0;
  const span=(start,end)=>{
    let cursor=String(start).slice(0,10),last=String(end||start).slice(0,10),n=0;
    while(cursor<=last && n++<180) {dates.add(cursor);cursor=new Date(Date.parse(cursor+'T12:00:00Z')+86400000).toISOString().slice(0,10);}
    if(cursor<=last) throw Error('Patch date span exceeds its safety bound.');
  };
  const oldDoctors=new Map(stored.doctors.map(r=>[r.doctor_key,[file.id,sourceType,r.doctor_key,r.display_name,r.seniority||'',r.membership_source||'roster',r.provider_staff_id||'']]));
  const desiredDoctors=new Set(doctorRows.map(r=>r[2]));
  for(const row of doctorRows) if(JSON.stringify(row)!==JSON.stringify(oldDoctors.get(row[2]))) {entries.push({kind:'doctor',row});changedFacts++;changedDoctors.add(row[2]);}
  for(const key of oldDoctors.keys()) if(!desiredDoctors.has(key)) {entries.push({kind:'removeDoctor',row:[key]});changedFacts++;changedDoctors.add(key);}
  for(const [kind,desired,old,columns] of [['event',eventRows,stored.events,eventColumns],['issue',issueRows,stored.issues,issueColumns]]) {
    const oldMap=new Map(old.map(r=>[r.id,JSON.stringify(columns.map(k=>k==='provider_staff_id'?(r[k]||''):r[k]))]));
    const desiredIds=new Set(desired.map(r=>r[0])),affected=new Set();
    for(const row of desired) {
      if(JSON.stringify([row[0],...row.slice(3)])===oldMap.get(row[0])) continue;
      entries.push({kind,row});changedFacts++;affected.add(row[0]);
      if(kind==='event') {
        span(row[5],row[6]);
        const event=JSON.parse(row[16]);
        for(const presence of expandEventToDailyPresenceRows(event,{sourceType,doctorKey:row[3],displayName:row[4]})) entries.push({kind:'presence',row:[presence.date,presence.sourceType,presence.doctorKey,presence.displayName,row[0]]});
      }
    }
    for(const row of old) {
      if(!desiredIds.has(row.id)) {entries.push({kind:kind==='event'?'removeEvent':'removeIssue',row:[row.id]});changedFacts++;affected.add(row.id);}
      if(kind==='event' && (affected.has(row.id) || changedDoctors.has(row.doctor_key))) span(row.start_date,row.end_date);
    }
  }
  for(const [kind,facts,old,key] of [['staff',compact.contributions,stored.staff,'doctor_key'],['catalog',compact.catalog,stored.catalog,'catalog_key']]) {
    const previous=new Map(old.map(r=>[r.term_start+'|'+r[key],r.fact_digest])),desired=new Set();
    for(const fact of facts.values()) {
      const marker=fact.termStart+'|'+(kind==='staff'?fact.doctorKey:fact.catalogKey);desired.add(marker);
      if(previous.get(marker)===fact.factDigest) continue;
      const row=kind==='staff'?[fact.sourceType,fact.termStart,fact.doctorKey,fact.fileId,fact.displayName,fact.seniority,fact.membershipSource,fact.providerStaffId,fact.firstApplicableDate,fact.lastApplicableDate,fact.factDigest,now]
        :[fact.sourceType,fact.termStart,fact.fileId,fact.catalogKey,fact.seniority,fact.title,fact.rawValue,fact.location,fact.allDay,fact.timeLabel,fact.startTime,fact.endTime,fact.firstDate,fact.lastDate,fact.factDigest,now];
      entries.push({kind,row});
    }
    for(const marker of previous.keys()) if(!desired.has(marker)) entries.push({kind:kind==='staff'?'removeStaff':'removeCatalog',row:marker.split('|')});
  }
  if(changedFacts>1250 || entries.length>1500 || entries.length*32+1024+doctors.length*4>50000 || dates.size>180) {
    const error=Error('Roster patch exceeds its reviewed correction envelope.');error.code='ROSTER_INCREMENTAL_BUDGET';error.boundedReplacementRequired=true;throw error;
  }
  const batches=[];
  for(let i=0;i<entries.length;i+=16) {
    const batch={index:batches.length,entries:wire(entries.slice(i,i+16))};
    if(JSON.stringify(batch).length>24576) throw Error('One patch batch exceeds its byte budget.');
    batches.push(batch);
  }
  const manifest={fileId:file.id,baselineRevision:baseline.baselineRevision,baselineParsedAt:baseline.baselineParsedAt,contentRevision,
    doctorCount:doctors.length,eventCount:eventRows.length,issueCount:issueRows.length,changedFacts,dates:[...dates].sort(),reservedWrites:entries.length*32+1024+doctors.length*4,
    coverage:{start:compact.coverageStart,end:compact.coverageEnd,staffDigest:compact.staffDigest,dailyDigest:compact.dailyDigest},
    batches:await Promise.all(batches.map(async b=>({index:b.index,count:b.entries.length,hash:await rosterImportDigest(b)})))};
  return {manifest,batches,patchRevision:await rosterImportDigest(manifest)};
}
export async function executeRosterPatch(payload,baseline,callback) {
  if (!baseline.baselineRevision) {
    baseline = await callback({phase:'bounded-patch-start',targetFileId:baseline.fileId});
    if(baseline.deferred) return baseline;
  }
  if(baseline.activated) return callback({phase:"bounded-finalize",patchRevision:baseline.patchRevision});
  const stored={};
  for(const [section,limit] of [['doctors',512],['events',25000],['issues',5000],['staff',750],['catalog',750]]) {
    const rows=[];let after='';
    do {
      const page=await callback({phase:'bounded-patch-page',section,after});
      if(page.deferred) return page;
      rows.push(...page.rows);after=page.next;
      if(rows.length>limit) throw Error('Retained patch facts exceed the inspection limit.');
    } while(after);
    stored[section]=rows;
  }
  const plan=await planRosterPatch(payload,baseline,stored);
  const pinned=await callback({phase:'bounded-patch-plan',manifest:plan.manifest,patchRevision:plan.patchRevision});
  if(pinned.deferred) return pinned;
  for(const batch of plan.batches.slice(Number(pinned.nextBatch||0))) {
    const result=await callback({phase:'bounded-patch-stage',batch,patchRevision:plan.patchRevision});
    if(result.deferred) return result;
  }
  const activated=await callback({phase:'bounded-patch-activate',patchRevision:plan.patchRevision,deferCompletion:true});
  if(activated.deferred || activated.completed)return activated;
  return callback({phase:'bounded-finalize',patchRevision:plan.patchRevision});
}
