import inventory from './d1-account-inventory.js';

// Exact, completed synthetic runs only. These databases were deleted before
// this settled report; no ordinary caller or production/provider binding ever
// used them. This is a forecast classification, never a usage refund.
// Evidence: docs/restoration-batch6-safeguard-review-2026-10-18.json.
const proofs=Object.freeze([
 {databaseId:'7b87d1be-2602-4a3e-8841-f6d525ad323d',utcDay:'2026-10-18',observedUntil:'2026-10-18T10:39:09.104Z',reads:25654,writes:18518},
 {databaseId:'a19cebcf-1ccd-43d5-84cd-cd105333cbdb',utcDay:'2026-10-18',observedUntil:'2026-10-18T10:39:09.104Z',reads:4220,writes:355}
]);
export function verifiedSyntheticTestCosts(analytics,day){
 const total={reads:0,writes:0},seen=new Set();
 for(const proof of proofs){
  if(proof.utcDay!==day||analytics.interval.start.slice(0,10)!==day||analytics.interval.observedUntil<proof.observedUntil)continue;
  if(seen.has(proof.databaseId)||inventory.databases.find(d=>d.id===proof.databaseId)?.environment!=='temporary-deleted')throw Error('Synthetic test evidence overlaps or is not retired.');
  seen.add(proof.databaseId);
  const measured=analytics.databases.find(d=>d.databaseId===proof.databaseId);
  if(!measured||Number(measured.rowsRead)<proof.reads||Number(measured.rowsWritten)<proof.writes)throw Error('Synthetic test costs have not reconciled with account analytics.');
  total.reads+=proof.reads;total.writes+=proof.writes;
 }
 return total;
}
