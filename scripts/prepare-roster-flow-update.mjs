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
const contentName = original.Get_file_content ? 'Get_file_content' : 'Get_file_content_using_path';
const metadataName = original.Get_file_metadata ? 'Get_file_metadata' : 'Get_file_metadata_using_path';
const content = original[contentName];
if(original.HTTP?.inputs?.uri!=='https://roster-to-calendar.pages.dev/api/automation/ingest' || !original.HTTP?.inputs?.headers || !content || !['GetFileContent','GetFileContentByPath'].includes(content.inputs?.host?.operationId))throw Error('Export does not match a reviewed SharePoint ingestion layout.');
const trigger=Object.values(definition.triggers)[0];
const oldFilter=JSON.stringify(trigger.conditions||[]);
const isPaeds=oldFilter.includes('Paeds');const isAdults=oldFilter.includes('Adult');const isVhh=oldFilter.includes('Active Medical Roster');
if(!isPaeds&&!isAdults&&!isVhh)throw Error('Exact reviewed site trigger required.');
const sourceId=isPaeds?'monash-paeds':isAdults?'monash-adults':'vhh-active-medical-roster';
trigger.conditions=[{expression:isVhh?"@equals(triggerOutputs()?['body/{FilenameWithExtension}'], 'Active Medical Roster.xlsx')":isPaeds?"@startsWith(triggerOutputs()?['body/{FilenameWithExtension}'], 'Paeds - Term ')":"@startsWith(triggerOutputs()?['body/{FilenameWithExtension}'], 'AdultTerm')"}];
trigger.runtimeConfiguration={...trigger.runtimeConfiguration,concurrency:{runs:1}};
const stopped=name=>({actions:{[name]:{type:'Terminate',inputs:{runStatus:'Succeeded'},runAfter:{}}}});
const stable=(left,right,after,name)=>({type:'If',expression:{and:[{equals:[left,right]}]},actions:{},else:stopped(name),runAfter:{[after]:['Succeeded']}});
const metadata=structuredClone(original[metadataName] || content);
metadata.runAfter={};
metadata.inputs.host.operationId=content.inputs.host.operationId==='GetFileContentByPath'?'GetFileMetadataByPath':'GetFileMetadata';
delete metadata.inputs.parameters.inferContentType;
const latest=structuredClone(metadata);latest.runAfter={Delay:['Succeeded']};
const verify=structuredClone(metadata);verify.runAfter={[contentName]:['Succeeded']};
const download=structuredClone(content);download.runAfter={Still_latest:['Succeeded']};
const ingest=structuredClone(original.HTTP);ingest.runAfter={Stable_after_download:['Succeeded']};
const headers=structuredClone(original.HTTP.inputs.headers);
const actions={
 Delay:{type:'Wait',inputs:{interval:{count:30,unit:'Second'}},runAfter:{}},
 Get_latest_metadata:latest,
 Still_latest:stable(`@outputs('${metadataName}')?['body/ETag']`,"@outputs('Get_latest_metadata')?['body/ETag']",'Get_latest_metadata','Stop_changed_during_settle'),
 [contentName]:download,
 Verify_metadata:verify,
 Stable_after_download:stable("@outputs('Get_latest_metadata')?['body/ETag']","@outputs('Verify_metadata')?['body/ETag']",'Verify_metadata','Stop_changed_during_download'),
 HTTP:ingest,
};
const check={type:'Http',inputs:{uri:'https://roster-to-calendar.pages.dev/api/automation/roster-check',method:'POST',headers,body:{sourceId,fileName:"@{triggerOutputs()?['body/{FilenameWithExtension}']}",providerVersion:isVhh?`@{outputs('${metadataName}')?['body/ETag']}`:"@{triggerOutputs()?['body/{VersionNumber}']}"},retryPolicy:{type:'none'}},runAfter:{},runtimeConfiguration:{secureData:{properties:['inputs','outputs']}}};
const current=stable("@triggerBody()?['Modified']",`@outputs('${metadataName}')?['body/LastModified']`,metadataName,'Stop_superseded_trigger');
current.actions={Check_roster_version:check,Changed_roster:{type:'If',expression:{and:[{equals:["@body('Check_roster_version')?['download']",true]}]},actions,else:{actions:{}},runAfter:{Check_roster_version:['Succeeded']}}};
definition.actions={[metadataName]:metadata,Current_trigger:current};
// The HTTP headers are reused verbatim; do not print/export them separately.
const packageFiles={};for(const path of paths)packageFiles[relative(root,path).split('\\').join('/')]=path===definitionPath?strToU8(JSON.stringify(resource)):new Uint8Array(await readFile(path));
await mkdir(resolve(destination,'..'),{recursive:true,mode:0o700});await writeFile(destination,zipSync(packageFiles),{mode:0o600});await chmod(destination,0o600);
console.log(JSON.stringify({prepared:true,site:isPaeds?'MCH':isAdults?'MMC':'VHH',beforeDownloadVersionCheck:true,settleSeconds:30,existingConnectionPreserved:true}));
