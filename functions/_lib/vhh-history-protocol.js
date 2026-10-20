import { loadRosterSyncRun, australianTermStartForDate } from './d1-calendar.js';
import { beginBoundedRosterImport,stageBoundedRosterBatch,prepareBoundedRosterPresence,prepareBoundedRosterMetadata,activateBoundedRosterTerm } from './roster-import-staging.js';
import { activeVhhOwnership,historyJobId,historyFileName,historyWindow,VHH_HISTORY_SOURCE } from './vhh-history-preservation.js';
import { rosterImportDigest } from './roster-import-batches.js';
import { rosterDeliveryOrder } from './roster-delivery-order.js';
import { pinManualAuthority } from './roster-authority.js';
import { reserveRosterMaintenanceBudget } from './roster-maintenance-budget.js';
import { automaticFacilityPublicationEnabled } from './facility-refresh-queue.js';

// Runs behind /derived's token, source, pause, request-meter and account guards.
// No child sync deliveries: historical supplements cannot supersede their parent.
export async function handleVhhHistoryRequest(context,body) {
  const db=context.env.ROSTER_DB;
  try {
    if(body.sourceId!==VHH_HISTORY_SOURCE || context.env.ROSTER_AUTOMATION_BOUNDED_IMPORT_ENABLED!=='true') throw Error('VHH bounded history is unavailable.');
    if(!await reserveRosterMaintenanceBudget(db,0,32768)) return Response.json({ok:true,deferred:true});
    if(body.phase==='bounded-history-inspect') {
      let runId=body.runId;
      if(!runId) runId=(await db.prepare(`SELECT r.id FROM roster_sync_runs r JOIN raw_roster_files raw ON raw.file_id=COALESCE(NULLIF(r.source_file_id,''),r.file_id)
        WHERE r.source_id=? AND r.status='success' ORDER BY r.started_at DESC LIMIT 1`).bind(VHH_HISTORY_SOURCE).first())?.id;
      const run=await loadRosterSyncRun(db,runId||'');
      if(!run || run.sourceId!==VHH_HISTORY_SOURCE) throw Error('Retained VHH input is unavailable.');
      const raw=await db.prepare('SELECT name,type,size,last_modified FROM raw_roster_files WHERE file_id=?').bind(run.sourceFileId||run.fileId).first();
      if(!raw || Number(raw.size)>15000000) throw Error('Retained VHH input exceeds its bounds.');
      const active=await activeVhhOwnership(db);
      const ordered=[...active].sort((a,b)=>a.firstDate.localeCompare(b.firstDate));
      let last='',needsRecovery=false;
      for(const file of ordered) {
        if(last && Date.parse(file.firstDate)-Date.parse(last)>86400000 && last<new Date().toISOString().slice(0,10)) needsRecovery=true;
        if(file.lastDate>last)last=file.lastDate;
      }
      return Response.json({ok:true,active,needsRecovery,run:{id:run.id,sourceId:run.sourceId,fileId:run.sourceFileId||run.fileId,contentHash:run.contentHash,providerVersion:run.providerVersion,fileName:raw.name,contentType:raw.type,lastModified:raw.last_modified}});
    }
    const parent=await loadRosterSyncRun(db,String(body.runId||''));
    if(!parent || parent.sourceId!==VHH_HISTORY_SOURCE || !['success','queued','processing'].includes(parent.status)) throw Error('Historical input does not match a retained VHH delivery.');
    // A completed parent's authority fence predates its own activation. This
    // additive supplement takes no term authority: its freshly pinned child
    // fence guards the historical term, while this guard checks raw ordering.
    const delivery=await rosterDeliveryOrder(db,parent,[],false);
    if(delivery.superseded) throw Error('A newer VHH input superseded this historical preparation.');
    const {from,to}=historyWindow(body.vhhHistory?.from,body.vhhHistory?.to);
    if(to>=new Date().toISOString().slice(0,10) || australianTermStartForDate(from)!==australianTermStartForDate(to)) throw Error('History requires a past window within one medical term.');
    const id=historyJobId(parent.id,from,to), file={...body.file,sourceId:VHH_HISTORY_SOURCE,sourceType:'vhh',active:false};
    const fileId=`vhh-history:${await rosterImportDigest([parent.id,from,to])}`;
    if(file.id!==fileId || file.name!==historyFileName(from,to)) throw Error('Historical file identity differs from its parent/window.');
    let result;
    switch(body.phase) {
      case 'bounded-begin':
      case 'bounded-initialize': {
        if(body.manifest?.startDate!==from || body.manifest?.rosterEndDate!==to || body.manifest?.presenceMode!=='cached') throw Error('Historical plan differs from the requested dates.');
        await pinManualAuthority(db,id,'vhh');
        result=await beginBoundedRosterImport(db,id,file,{manifest:body.manifest,revision:body.revision},{holdOwnerId:parent.id});break;
      }
      case 'bounded-events': result=await stageBoundedRosterBatch(db,id,file,body.revision,body.batch);break;
      case 'bounded-presence': result=await prepareBoundedRosterPresence(db,id,file,body.revision,body.batch);break;
      case 'bounded-metadata': result=await prepareBoundedRosterMetadata(db,id,body.revision,body.batch);break;
      case 'bounded-activate':
        if(body.prepareOnly) {
          const job=await db.prepare('SELECT compact_ready,plan_revision FROM roster_import_jobs WHERE run_id=?').bind(id).first();
          if(!job?.compact_ready || job.plan_revision!==body.revision) throw Error('Historical preparation is incomplete.');
          result={prepared:true};
        } else result=await activateBoundedRosterTerm(db,id,body.revision,{publishFacility:automaticFacilityPublicationEnabled(context.env,'vhh'),deliveryGuard:delivery.guard,historySupplement:true});
        break;
      case 'bounded-finalize': result={completed:true};break;
      default: throw Error('Unsupported historical phase.');
    }
    return Response.json({ok:true,...result});
  } catch(error) {
    if(error.code==='ROSTER_MAINTENANCE_DEFERRED') return Response.json({ok:true,deferred:true});
    console.error('VHH historical preservation stopped',{phase:body.phase,error:error.message});
    return Response.json({error:error.message,resumable:true},{status:422});
  }
}
