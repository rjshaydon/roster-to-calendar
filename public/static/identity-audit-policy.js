export function identityAuditWindow(now=new Date()) {
 const parts=Object.fromEntries(new Intl.DateTimeFormat('en-CA',{timeZone:'Australia/Melbourne',weekday:'short',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',hourCycle:'h23'}).formatToParts(now).map(p=>[p.type,p.value]));
 return {eligible:parts.weekday==='Sun' && Number(parts.hour)>=2 && Number(parts.hour)<4,weekKey:`${parts.year}-${parts.month}-${parts.day}`};
}
