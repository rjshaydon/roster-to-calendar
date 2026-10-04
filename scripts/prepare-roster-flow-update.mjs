// Prepare a recoverable update using the original export's existing connector
// and ingestion credential. Output is private and must never enter the repo.
import { readFile, writeFile, readdir, mkdir, chmod } from 'node:fs/promises';
import { resolve, join, relative } from 'node:path';
import { zipSync, strToU8 } from 'fflate';
const [input,output]=process.argv.slice(2);
if(!input||!output)throw Error('Usage: node scripts/prepare-roster-flow-update.mjs PRIVATE_EXPORT_DIRECTORY PRIVATE_OUTPUT.zip');
const root=resolve(input), destination=resolve(output);
if(destination.startsWith(resolve('.')+'/'))throw Error('Flow packages contain existing credentials and must remain outside the repository.');
async function files(dir){const found=[];for(const entry of await readdir(dir,{withFileTypes:true})){const path=join(dir,entry.name);found.push(...(entry.isDirectory()?await files(path):[path]));}return found;}
const paths=await files(root);const definitionPath=paths.find(path=>path.endsWith('/definition.json'));
if(!definitionPath)throw Error('Original export definition required.');
const resource=JSON.parse(await readFile(definitionPath,'utf8'));const definition=resource.properties.definition;
const original=definition.actions;
if(!original.HTTP?.inputs?.headers||!original.Get_file_content||!original.Get_file_metadata||!original.Condition||!original.Delay)throw Error('Export does not match the reviewed metadata/content ingestion layout. Inspect this Flow before updating it.');
const source=original.HTTP.inputs.body.sourceId;
const trigger=Object.values(definition.triggers)[0];
const oldFilter=JSON.stringify(trigger.conditions||[]);
const isPaeds=oldFilter.includes('Paeds');const isAdults=oldFilter.includes('Adult');const isVhh=oldFilter.includes('Active Medical Roster');
if(!isPaeds&&!isAdults&&!isVhh)throw Error('Exact reviewed site trigger required.');
trigger.conditions=[{expression:isVhh?"@equals(triggerOutputs()?['body/{FilenameWithExtension}'], 'Active Medical Roster.xlsx')":isPaeds?"@startsWith(triggerOutputs()?['body/{FilenameWithExtension}'], 'Paeds - Term ')":"@startsWith(triggerOutputs()?['body/{FilenameWithExtension}'], 'AdultTerm')"}];
trigger.runtimeConfiguration={...trigger.runtimeConfiguration,concurrency:{runs:1}};
const stopped=()=>({actions:{Stop_stale_version:{type:'Terminate',inputs:{runStatus:'Succeeded'},runAfter:{}}}});
const stable=(left,right,after)=>({type:'If',expression:{and:[{equals:[left,right]}]},actions:{},else:stopped(),runAfter:{[after]:['Succeeded']}});
const actions=structuredClone(original);
actions.Delay.inputs.interval={count:30,unit:'Second'};
actions.Get_file_metadata.runAfter={};
actions.Condition.runAfter={Get_file_metadata:['Succeeded']};
actions.Delay.runAfter={Condition:['Succeeded']};
actions.Get_latest_metadata=structuredClone(actions.Get_file_metadata);actions.Get_latest_metadata.runAfter={Delay:['Succeeded']};
actions.Still_latest=stable("@outputs('Get_file_metadata')?['body/LastModified']","@outputs('Get_latest_metadata')?['body/LastModified']",'Get_latest_metadata');
actions.Get_file_content.runAfter={Still_latest:['Succeeded']};
actions.Verify_metadata=structuredClone(actions.Get_file_metadata);actions.Verify_metadata.runAfter={Get_file_content:['Succeeded']};
actions.Stable_after_download=stable("@outputs('Get_latest_metadata')?['body/LastModified']","@outputs('Verify_metadata')?['body/LastModified']",'Verify_metadata');
actions.HTTP.runAfter={Stable_after_download:['Succeeded']};
const headers=structuredClone(original.HTTP.inputs.headers);
definition.actions={
 Check_roster_version:{type:'Http',inputs:{uri:'https://roster-to-calendar.pages.dev/api/automation/roster-check',method:'POST',headers,body:{sourceId:source,fileName:"@{triggerOutputs()?['body/{FilenameWithExtension}']}",providerVersion:"@{triggerOutputs()?['body/{VersionNumber}']}"},retryPolicy:{type:'none'}},runAfter:{}},
 Changed_roster:{type:'If',expression:{and:[{equals:["@body('Check_roster_version')?['download']",true]}]},actions,else:{actions:{}},runAfter:{Check_roster_version:['Succeeded']}},
};
// The HTTP headers are reused verbatim; do not print/export them separately.
const packageFiles={};for(const path of paths)packageFiles[relative(root,path).split('\\').join('/')]=path===definitionPath?strToU8(JSON.stringify(resource)):new Uint8Array(await readFile(path));
await mkdir(resolve(destination,'..'),{recursive:true,mode:0o700});await writeFile(destination,zipSync(packageFiles),{mode:0o600});await chmod(destination,0o600);
console.log(JSON.stringify({prepared:true,site:isPaeds?'MCH':isAdults?'MMC':'VHH',beforeDownloadVersionCheck:true,settleSeconds:30,existingConnectionPreserved:true}));
