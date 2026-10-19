// Read-only, aggregate acceptance inside the already-authorized runner. Never
// emits credentials, names, person identifiers or clinical calendar contents.
const sourceType={'monash-adults':'mmc','monash-paeds':'mch','vhh-active-medical-roster':'vhh','dandenong-findmyshift':'ddh'}[process.env.ROSTER_AUTOMATION_SOURCE_ID];
const base=process.env.ROSTER_AUTOMATION_BASE_URL||'https://roster-to-calendar.pages.dev';
const term=process.env.VERIFY_TERM_START||'';
if(!sourceType||!/^\d{4}-\d{2}-\d{2}$/.test(term))throw Error('A supported source and explicit term start are required.');
const response=await fetch(base+'/api/automation/doctor-calendars',{method:'POST',headers:{authorization:'Bearer '+process.env.ROSTER_AUTOMATION_TOKEN,'Content-Type':'application/json'},body:JSON.stringify({mode:'delivery-status',sourceType}),signal:AbortSignal.timeout(30000)});
if(!response.ok)throw Error('Delivery verification returned HTTP '+response.status);
const status=await response.json();
const termPresent=status.terms?.some(t=>t.termStart===term)===true;
console.log(JSON.stringify({sourceType,term,ready:status.ready,matchingParent:status.matchingParent,termPresent,monthCount:status.monthCount}));
if(!status.ready||!termPresent)throw Error('The requested term is not yet in a matching delivered personal-calendar generation.');
