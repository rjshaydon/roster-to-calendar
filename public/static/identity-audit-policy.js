export function identityAuditWindow(now=new Date()) {
 const parts=Object.fromEntries(new Intl.DateTimeFormat('en-CA',{timeZone:'Australia/Melbourne',weekday:'short',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',hourCycle:'h23'}).formatToParts(now).map(p=>[p.type,p.value]));
 const minutes=Number(parts.hour)*60+Number(parts.minute);
 return {eligible:parts.weekday==='Sun' && minutes>=210 && minutes<270,weekKey:`${parts.year}-${parts.month}-${parts.day}`};
}
