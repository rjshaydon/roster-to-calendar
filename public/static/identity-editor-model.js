// Names propose identifiers; they never establish that roster entries match.
export function interpretIdentityName(value) {
 const name=String(value||'').trim().replace(/\s+/g,' ');
 const comma=name.indexOf(',');
 if(comma>=0) return {given:name.slice(comma+1).trim(),surname:name.slice(0,comma).trim()};
 const words=name.split(' ').filter(Boolean);
 const upper=word=>/\p{L}/u.test(word)&&word===word.toUpperCase()&&word!==word.toLowerCase();
 const capitals=words.filter(upper);
 if(capitals.length && capitals.length<words.length) return {given:words.filter(w=>!upper(w)).join(' '),surname:capitals.join(' ')};
 return {given:words.slice(0,-1).join(' '),surname:words.at(-1)||''};
}
export function identityIdFromParts({given,surname}) {
 const slug=[surname,given].join(' ').normalize('NFKD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'');
 return slug?'person:'+slug:'';
}
export function identityIdFromEntry(value) {
 const raw=String(value||'').trim().replace(/^person:/,'');
 return 'person:'+raw.normalize('NFKD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'');
}
export const identityHistoryLabel=operation=>({merge:'Merged people',name:'Changed preferred name',id:'Changed person ID','alias-move':'Separated roster name','account-link':'Linked account',reverse:'Undid a change'}[operation.kind]||'Changed person');
