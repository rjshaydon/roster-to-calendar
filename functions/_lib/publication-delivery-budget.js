export const PUBLICATION_STEP_WRITES=112; // 64 execution/control + 48 receipt reservation
export const PUBLICATION_STEP_READS=125048;
// Holds are capacity commitments, not usage. A step's in-flight receipt and its
// capacity overlap conservatively until the cursor/remaining-tail checkpoint
// commits atomically. A killed/unknown request cannot erase future capacity.
export const publicationHeldWrites=`COALESCE((SELECT SUM(remaining_writes) FROM roster_publication_holds h WHERE ${activeHold()}),0)`;
export const publicationHeldReads=`COALESCE((SELECT SUM(remaining_reads) FROM roster_publication_holds h WHERE ${activeHold()}),0)`;
function activeHold(){return `h.source_type IN ('mmc','mch','ddh','vhh') AND ((h.owner_kind='publication' AND EXISTS(SELECT 1 FROM facility_refresh_jobs j WHERE j.source_type=h.source_type AND j.term_start=h.term_start AND j.request_revision=h.owner_id AND j.status='pending')) OR (h.owner_kind='staging' AND EXISTS(SELECT 1 FROM roster_sync_runs r WHERE r.id=h.owner_id AND r.status IN ('queued','processing'))))`;}


export function publicationTail(dates,plan=null,nextBatch=0,nextMonth=0){
 if(!Array.isArray(dates)||dates.length>180)throw Error('Invalid delivery dates.');
 const batches=plan?Math.max(0,plan.batchCount-nextBatch):Math.ceil(dates.length/7);
 const months=plan?Math.max(0,plan.months.length-nextMonth):new Set(dates.map(d=>d.slice(0,7))).size;
 // Planning (if needed), remaining daily/monthly steps, finalize and its two
 // external-artifact completion callbacks. R2 uploads themselves do not use D1.
 const steps=(plan?0:1)+batches+months+3;
 return {writes:steps*PUBLICATION_STEP_WRITES,reads:steps*PUBLICATION_STEP_READS};
}
export function publicationHoldStatement(db,source,term,owner,dates,{plan=null,nextBatch=0,nextMonth=0,kind='publication',guard='1=1',bindings=[],jobOwner=false}={}){
 const cost=publicationTail(dates,plan,nextBatch,nextMonth);
 if(kind==='staging')cost.reads=PUBLICATION_STEP_READS; // protect next read step; full tail pins with its plan
 const queuedSteps=`(SELECT 4+(json_array_length(dates_json)+6)/7+(SELECT COUNT(DISTINCT substr(value,1,7)) FROM json_each(dates_json)) FROM facility_refresh_jobs WHERE source_type=? AND term_start=?)`;
 return db.prepare(`INSERT INTO roster_publication_holds(source_type,term_start,owner_id,owner_kind,remaining_writes,remaining_reads)
 SELECT ?,?,${jobOwner?'(SELECT request_revision FROM facility_refresh_jobs WHERE source_type=? AND term_start=?)':'?'},?,${jobOwner?queuedSteps+'*112':'?'},${jobOwner?queuedSteps+'*125048':'?'} WHERE ${guard}
 ON CONFLICT(source_type,term_start) DO UPDATE SET owner_id=excluded.owner_id,owner_kind=excluded.owner_kind,
 remaining_writes=excluded.remaining_writes,remaining_reads=excluded.remaining_reads
 WHERE (roster_publication_holds.owner_kind='staging' OR roster_publication_holds.owner_id=excluded.owner_id OR excluded.owner_kind='publication')
 AND (roster_publication_holds.owner_id<>excluded.owner_id OR roster_publication_holds.owner_kind<>excluded.owner_kind
 OR roster_publication_holds.remaining_writes<>excluded.remaining_writes OR roster_publication_holds.remaining_reads<>excluded.remaining_reads)`)
 .bind(source,term,...(jobOwner?[source,term]:[owner]),kind,...(jobOwner?[source,term,source,term]:[cost.writes,cost.reads]),...bindings);
}
export function releasePublicationHoldStatement(db,source,term,owner){
 return db.prepare('DELETE FROM roster_publication_holds WHERE source_type=? AND term_start=? AND owner_id=? AND EXISTS(SELECT 1 FROM facility_refresh_jobs j WHERE j.source_type=roster_publication_holds.source_type AND j.term_start=roster_publication_holds.term_start AND j.request_revision=roster_publication_holds.owner_id AND j.status=\'complete\')').bind(source,term,owner);
}

// A retained staged import commits delivery capacity before its next fact batch.
// This is not recorded as actual spend. The hold transfers to the publication
// job atomically at activation; a superseding import replaces its staging owner.
export async function stageDeliveryHolds(db,source,owner,start,end,termForDate){
 const terms=new Map();let date=start,count=0;
 while(date&&date<=end){
  if(++count>180)throw Error('Delivery range exceeds its import bound.');
  const term=termForDate(date);if(!terms.has(term))terms.set(term,[]);terms.get(term).push(date);
  date=new Date(Date.parse(date+'T12:00:00Z')+86400000).toISOString().slice(0,10);
 }
 if(terms.size>3)throw Error('Delivery exceeds its term bound.');
 if(terms.size)await db.batch([...terms].map(([term,dates])=>publicationHoldStatement(db,source,term,owner,dates,{kind:'staging'})));
}
