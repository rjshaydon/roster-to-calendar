// Runner-side scheduling only. The server still enforces every reservation,
// expiry, source fence and daily ceiling independently.
export function createMaintenanceAdmission({request,now=Date.now}) {
 let nextCheck=0;
 return async function ensure({force=false}={}) {
  if(!force&&now()<nextCheck)return {ok:true};
  const value=await request({reuseCurrentGrant:!force});
  if(value.deferred||value.paused){nextCheck=0;return value;}
  const expiry=Date.parse(value.validUntil||value.budget?.validUntil||'');
  // Older deployments omit expiry: check again soon rather than assuming
  // that a reused allocation has a new five-minute lifetime.
  nextCheck=Math.min(now()+300000,Number.isFinite(expiry)?expiry-60000:now()+30000);
  return value;
 };
}

export function isReadOnlyPatchPage(path,body) {
 return path==='/api/automation/derived'&&body?.phase==='bounded-patch-page';
}

export function retryablePatchReadFailure(status,result={}) {
 return [408,429,500,502,503,504].includes(status)
  ||status===422&&/D1_ERROR.*(?:Network connection lost|connection reset|temporarily unavailable)/i.test(String(result.error||''));
}

// Repeating comparison reads cannot apply roster changes. Bound retries and
// keep the same cursor; the endpoint checks the active baseline on every call.
export async function retryPatchRead(operation,{sleep=ms=>new Promise(resolve=>setTimeout(resolve,ms))}={}) {
 for(let attempt=0;attempt<3;attempt++) {
  try{return await operation();}catch(error){
   if(!error.retryablePatchRead||attempt===2)throw error;
   await sleep((attempt+1)*2000);
  }
 }
}
