// Initialise only when a caller needs this timezone; unrelated routes pay no ICU setup cost.
let auditFormatter;
export function identityAuditWindow(now=new Date()) {
 auditFormatter ||= new Intl.DateTimeFormat('en-CA',{timeZone:'Australia/Melbourne',weekday:'short',year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit',hourCycle:'h23'});
 const parts=Object.fromEntries(auditFormatter.formatToParts(now).map(p=>[p.type,p.value]));
 const minutes=Number(parts.hour)*60+Number(parts.minute);
 return {eligible:parts.weekday==='Sun' && minutes>=210 && minutes<270,
  catchupEligible:parts.weekday==='Sun' && minutes>=210,weekKey:`${parts.year}-${parts.month}-${parts.day}`};
}
