const sources={'monash-adults':'mmc','monash-paeds':'mch','vhh-active-medical-roster':'vhh','dandenong-findmyshift':'ddh'};
const key=source=>`roster-delivery/v1/${source}.json`;
export async function markRosterDeliveryDeferred(r2,sourceId,queuedAt) {
 const source=sources[sourceId];if(!source||!r2?.put||!queuedAt)return;
 try {
  const previous=await r2.get(key(source));
  if(previous&&(await previous.json()).queuedAt===queuedAt)return;
  await r2.put(key(source),JSON.stringify({sourceType:source,queuedAt,status:'deferred'}),{httpMetadata:{contentType:'application/json'}});
 } catch {console.warn('Roster delivery warning could not be published.');}
}
export async function clearRosterDeliveryWarning(r2,sourceId) {
 const source=sources[sourceId];try {if(source&&r2?.delete)await r2.delete(key(source));}catch{console.warn('Roster delivery warning could not be cleared.');}
}
export async function loadRosterDeliveryWarnings(r2,sourceTypes,now=Date.now()) {
 if(!r2?.get)return [];
 const warnings=await Promise.all([...new Set(sourceTypes)].filter(s=>Object.values(sources).includes(s)).map(async source=>{
  try {const object=await r2.get(key(source));if(!object)return null;
   const value=await object.json(),age=now-Date.parse(value.queuedAt);
   return value.status==='deferred'&&Number.isFinite(age)&&age>=15*60000?{sourceType:source,status:'delayed'}:null;
  }catch{return null;}
 }));
 return warnings.filter(Boolean);
}
