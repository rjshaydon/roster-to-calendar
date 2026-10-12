// Verified one-off CLI operations against the isolated identity test database.
// Native response/query hashes and settled analytics bound these costs. Neither
// SQL operation touched the maintenance receipt ledger. Historical facts only:
// outside their UTC day/cutoff they contribute nothing to the forecast.
export const verifiedOneOffMaintenanceProofs=Object.freeze([
  {
    "id": "identity-native-seed",
    "databaseId": "741ce534-eafe-4aa7-9305-7b7ed5661979",
    "utcDay": "2026-10-12",
    "observedAt": "2026-10-12T03:29:54.758946Z",
    "rowsRead": 0,
    "rowsWritten": 12000,
    "querySha256": "328db9b7e75d93608a083a12ed509737d50de9f8cb22f5a06e54b38ec4a7d8dc",
    "nativeResponseSha256": "57764ed84d608a66abe35a45b224f9d049d4cde14c86b7c727d2e2011e1e0d01",
    "transport": "native-d1-cli",
    "receiptOverlap": false
  },
  {
    "id": "identity-native-cleanup",
    "databaseId": "741ce534-eafe-4aa7-9305-7b7ed5661979",
    "utcDay": "2026-10-12",
    "observedAt": "2026-10-12T03:44:40.504575Z",
    "rowsRead": 9190,
    "rowsWritten": 5200,
    "querySha256": "164724f95605827ca86036b87a7c57fa235dde75fe48b9e817135d7bf2fa2aef",
    "nativeResponseSha256": "4b2046fdb58c7e3dc0038ddb7a7d76b35495bfa5fad9e17b7e61dfb308ab3419",
    "transport": "native-d1-cli",
    "receiptOverlap": false
  }
]);
export function verifiedOneOffMaintenanceCosts(analytics,databaseId,day){
 const costs={reads:0,writes:0};
 const seen=new Set();
 for(const proof of verifiedOneOffMaintenanceProofs){
  if(proof.databaseId!==databaseId||proof.utcDay!==day)continue;
  if(seen.has(proof.nativeResponseSha256)||proof.receiptOverlap!==false)throw Error('One-off maintenance evidence overlaps.');
  seen.add(proof.nativeResponseSha256);
  if(proof.observedAt>analytics.interval.observedUntil)continue;
  if(!proof.observedAt.startsWith(day+'T')||![proof.rowsRead,proof.rowsWritten].every(n=>Number.isSafeInteger(n)&&n>=0))throw Error('One-off maintenance evidence is invalid.');
  costs.reads+=proof.rowsRead;costs.writes+=proof.rowsWritten;
 }
 return costs;
}
