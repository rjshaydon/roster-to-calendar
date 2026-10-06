import assert from 'node:assert/strict';
import {interpretIdentityName,identityIdFromParts,identityIdFromEntry,groupIdentityTerms} from '../public/static/identity-editor-model.js';
for(const [name,given,surname,id] of [
 ['Jayantha WEERASINGE','Jayantha','WEERASINGE','person:weerasinge-jayantha'],
 ['WEERASINGE, Jayantha','Jayantha','WEERASINGE','person:weerasinge-jayantha'],
 ['Mary Anne VAN DYKE','Mary Anne','VAN DYKE','person:van-dyke-mary-anne'],
 ['Van Dyke, Mary Anne','Mary Anne','Van Dyke','person:van-dyke-mary-anne'],
 ['Mary Anne Smith','Mary Anne','Smith','person:smith-mary-anne'],
 ['MARY ANNE SMITH','MARY ANNE','SMITH','person:smith-mary-anne'],
 ['  René   O’BRIEN ','René','O’BRIEN','person:o-brien-rene'],
]){const parts=interpretIdentityName(name);assert.deepEqual(parts,{given,surname});assert.equal(identityIdFromParts(parts),id);}
assert.equal(identityIdFromEntry('person:WEERASINGE Jayantha'),'person:weerasinge-jayantha');
console.log('Identity name interpretation and identifier normalisation passed.');

const grouped=groupIdentityTerms([{sourceType:'ddh',termStart:'2026-08-03',termEnd:'2026-11-01',grades:['HMO'],rosterNames:['Jun Lee']},{sourceType:'mmc',termStart:'2026-08-03',termEnd:'2026-11-01',grades:['Registrar'],rosterNames:['Jeremy Lee']},{sourceType:'ddh',termStart:'2026-05-04',termEnd:'2026-08-02',grades:['Intern']}]);
assert.equal(grouped.length,2);assert.deepEqual(grouped[0].sites.map(s=>[s.sourceType,s.grades]),[['ddh',['HMO']],['mmc',['Registrar']]],'concurrent term rows retain the grade at each site');
