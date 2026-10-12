import assert from 'node:assert/strict';
const Original=Intl.DateTimeFormat;let constructors=0;
Intl.DateTimeFormat=class extends Original{constructor(...args){super(...args);constructors++;}};
try{
 await import('../functions/api/feed.js');await import('../functions/api/state.js');await import('../functions/api/automation/account-budget.js');await import('../functions/api/automation/identity-maintenance.js');
 assert.equal(constructors,0,'unrelated cold routes must not construct timezone formatters during module loading');
 const {identityAuditWindow}=await import('../public/static/identity-audit-policy.js');
 const {melbourneDateKey}=await import('../public/static/roster-term-policy.js');
 const {onShiftLaunchWindow}=await import('../public/static/shift-launch-policy.js');
 assert.equal(identityAuditWindow(new Date('2026-10-10T16:30:00Z')).eligible,true);
 identityAuditWindow(new Date('2026-10-10T17:30:00Z'));assert.equal(constructors,1);
 assert.equal(melbourneDateKey(new Date('2026-10-10T16:30:00Z')),'2026-10-11');melbourneDateKey(new Date());assert.equal(constructors,2);
 onShiftLaunchWindow([{id:'shift',source:'DDH',start:'2026-10-12T08:00:00',end:'2026-10-12T16:00:00'}],new Date('2026-10-11T22:00:00Z'));assert.equal(constructors,3);
 onShiftLaunchWindow([{id:'shift',source:'DDH',start:'2026-10-12T08:00:00',end:'2026-10-12T16:00:00'}],new Date('2026-10-11T22:00:00Z'));assert.equal(constructors,3,'each needed formatter is reused');
}finally{Intl.DateTimeFormat=Original;}
console.log('Cold route imports perform zero ICU constructions; requested Melbourne date/audit/shift conversions retain correct boundaries and reuse formatters.');
