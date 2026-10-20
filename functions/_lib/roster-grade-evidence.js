// Only dated source/term evidence may override an unknown file-directory grade.
export function knownRosterGrade(value){
 const grade=String(value||'').trim();
 return grade&&!['unknown','not recorded'].includes(grade.toLowerCase())?grade:'';
}
export function resolveRosterGrades(values,fallback='Unknown'){
 const grades=[...new Set(values.map(knownRosterGrade).filter(Boolean))];
 return grades.length===1?grades[0]:grades.length>1?'Unknown':knownRosterGrade(fallback)||'Unknown';
}
