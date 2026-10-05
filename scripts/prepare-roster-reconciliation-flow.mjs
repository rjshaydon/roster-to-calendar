// Build a private package using the reviewed MMC export's existing connection
// and ingestion credential. Import as a new Flow; inspect allowance before enabling.
import {readFile,writeFile,readdir,mkdir,chmod} from 'node:fs/promises';
import {resolve,join,relative} from 'node:path';
import {zipSync,strToU8} from 'fflate';
const [input,output,minutes='5']=process.argv.slice(2);
if(!input||!output||!['2','3','5','15'].includes(minutes))throw Error('Usage: PRIVATE_MMC_EXPORT PRIVATE_OUTPUT.zip [2|3|5|15 minutes]');
const root=resolve(input),destination=resolve(output);
if(destination.startsWith(resolve('.')+'/'))throw Error('Credential-bearing output must remain outside the repository.');
async function files(dir){const result=[];for(const entry of await readdir(dir,{withFileTypes:true})){const path=join(dir,entry.name);result.push(...(entry.isDirectory()?await files(path):[path]));}return result;}
const paths=await files(root),definitionPath=paths.find(path=>path.endsWith('/definition.json'));
const resource=JSON.parse(await readFile(definitionPath,'utf8'));
const definition=resource.properties.definition,original=definition.actions;
if(original.HTTP?.inputs?.uri!=='https://roster-to-calendar.pages.dev/api/automation/ingest'||original.Get_file_content?.inputs?.host?.operationId!=='GetFileContent')throw Error('Reviewed original MMC export required.');
const name='Reconcile current and next-term roster metadata';
resource.properties.displayName=name;
const headers=structuredClone(original.HTTP.inputs.headers);
const connector=(operationId,parameters,runAfter={})=>({type:'OpenApiConnection',inputs:{parameters,host:{...original.Get_file_content.inputs.host,operationId},authentication:original.Get_file_content.inputs.authentication},runAfter});
const rest=(dataset,table,filter)=>connector('HttpRequest',{dataset,'parameters/method':'GET','parameters/uri':`_api/web/lists(guid'${table}')/items?$select=FileRef,Modified,OData__UIVersionString,File/ETag&$expand=File&$filter=${filter}&$top=1000`,'parameters/headers':{Accept:'application/json;odata=nometadata'}});
const stop=name=>({actions:{[name]:{type:'Terminate',inputs:{runStatus:'Succeeded'},runAfter:{}}}});
const stable=(left,right,after,name)=>({type:'If',expression:{and:[{equals:[left,right]}]},actions:{},else:stop(name),runAfter:{[after]:['Succeeded']}});
const metadata=connector('GetFileMetadataByPath',{dataset:"@items('Changed_files')?['dataset']",path:"@items('Changed_files')?['path']"});
const content=connector('GetFileContentByPath',{dataset:"@items('Changed_files')?['dataset']",path:"@items('Changed_files')?['path']",inferContentType:true},{Version_still_current:['Succeeded']});
const verify=structuredClone(metadata);verify.runAfter={Get_file_content_using_path:['Succeeded']};
const ingest=structuredClone(original.HTTP);ingest.runAfter={Stable_download:['Succeeded']};
ingest.inputs.body={sourceId:"@items('Changed_files')?['sourceId']",fileName:"@items('Changed_files')?['fileName']",contentType:'application/octet-stream',contentBase64:"@body('Get_file_content_using_path')?['$content']",providerModifiedAt:"@items('Changed_files')?['providerModifiedAt']",providerVersion:"@items('Changed_files')?['providerVersion']"};
ingest.inputs.retryPolicy={type:'none'};
// Start in the future so import can be inspected and explicitly turned off before
// the chosen schedule is enabled. No capacity purchase/assignment is requested.
definition.triggers={Recurrence:{type:'Recurrence',recurrence:{frequency:'Minute',interval:Number(minutes),startTime:'2030-01-01T00:00:00Z'},runtimeConfiguration:{concurrency:{runs:1}}}};
definition.actions={
 Monash_metadata:rest('https://monashhealth.sharepoint.com/sites/MonashEDMMC-CLA-PFU-DEP-MedicalRoster','43f6a549-2e31-486a-8ca2-ecdbe383986a',"startswith(FileLeafRef,'AdultTerm') or startswith(FileLeafRef,'Paeds - Term ')") ,
 VHH_metadata:rest('https://monashhealth.sharepoint.com/sites/VHHED-VHH-EMG-DEP','dd1e780c-6901-4004-a615-73e8299158f4',"FileLeafRef eq 'Active Medical Roster.xlsx'"),
 Check_versions:{type:'Http',inputs:{uri:'https://roster-to-calendar.pages.dev/api/automation/roster-check',method:'POST',headers,body:{mode:'reconcile',libraries:[{site:'monash',files:"@coalesce(body('Monash_metadata')?['value'],json('[]'))",unavailable:"@not(equals(actions('Monash_metadata')?['status'],'Succeeded'))",nextLink:"@coalesce(body('Monash_metadata')?['odata.nextLink'],body('Monash_metadata')?['@odata.nextLink'])"},{site:'vhh',files:"@coalesce(body('VHH_metadata')?['value'],json('[]'))",unavailable:"@not(equals(actions('VHH_metadata')?['status'],'Succeeded'))",nextLink:"@coalesce(body('VHH_metadata')?['odata.nextLink'],body('VHH_metadata')?['@odata.nextLink'])"}]},retryPolicy:{type:'none'}},runAfter:{Monash_metadata:['Succeeded','Failed','TimedOut'],VHH_metadata:['Succeeded','Failed','TimedOut']},runtimeConfiguration:{secureData:{properties:['inputs','outputs']}}},
 Changed_files:{type:'Foreach',foreach:"@body('Check_versions')?['downloads']",runtimeConfiguration:{concurrency:{repetitions:1}},runAfter:{Check_versions:['Succeeded']},actions:{
 Get_metadata:metadata,
 Version_still_current:stable("@items('Changed_files')?['etag']","@body('Get_metadata')?['ETag']",'Get_metadata','Stop_superseded_candidate'),
 Get_file_content_using_path:content,
 Verify_metadata:verify,
 Stable_download:stable("@body('Get_metadata')?['ETag']","@body('Verify_metadata')?['ETag']",'Verify_metadata','Stop_changed_download'),
 HTTP:ingest,
 }},
};
// Terminate is not permitted inside a foreach. Scope each download under its
// metadata conditions instead, so one changed file never stops other sites.
const loop=definition.actions.Changed_files.actions;
loop.Version_still_current.else={actions:{}};
loop.Stable_download.else={actions:{}};
loop.HTTP.runAfter={};loop.Stable_download.actions={HTTP:loop.HTTP};
loop.Stable_download.runAfter={Verify_metadata:['Succeeded']};
loop.Get_file_content_using_path.runAfter={};
loop.Version_still_current.actions={Get_file_content_using_path:loop.Get_file_content_using_path,Verify_metadata:loop.Verify_metadata,Stable_download:loop.Stable_download};
definition.actions.Changed_files.actions={Get_metadata:loop.Get_metadata,Version_still_current:loop.Version_still_current};
const packageFiles={};
for(const path of paths){let bytes=new Uint8Array(await readFile(path));if(path===definitionPath)bytes=strToU8(JSON.stringify(resource));if(path===join(root,'manifest.json')){const manifest=JSON.parse(new TextDecoder().decode(bytes));manifest.details.displayName=name;for(const value of Object.values(manifest.resources))if(value.type==='Microsoft.Flow/flows'){value.suggestedCreationType='New';value.details.displayName=name;}bytes=strToU8(JSON.stringify(manifest));}packageFiles[relative(root,path).split('\\').join('/')]=bytes;}
await mkdir(resolve(destination,'..'),{recursive:true,mode:0o700});await writeFile(destination,zipSync(packageFiles),{mode:0o600});await chmod(destination,0o600);
console.log(JSON.stringify({prepared:true,unchangedActions:5,intervalMinutes:Number(minutes),unchangedDailyActions:5*1440/Number(minutes),startHeldUntil:'2030-01-01',existingConnectionPreserved:true}));
