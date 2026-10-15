// Builds an isolated native Workers benchmark. Does not deploy or read app data.
// Usage: node scripts/security/prepare-password-feasibility.mjs <hash-wasm package dir> <output dir>
// Obtain hash-wasm@4.12.0 using npm pack --ignore-scripts and unpack outside the app.
import fs from 'node:fs';
import path from 'node:path';
import { randomBytes, pbkdf2Sync, createHash } from 'node:crypto';
import { createRequire } from 'node:module';

const [packageArgument, outputArgument] = process.argv.slice(2);
if (!packageArgument || !outputArgument) throw new Error('Provide the unpacked hash-wasm package directory and a separate temporary output directory.');
const packageRoot = path.resolve(packageArgument), output = path.resolve(outputArgument);
if (output === packageRoot) throw new Error('Output must be separate from the package directory.');
const metadata = JSON.parse(fs.readFileSync(path.join(packageRoot, 'package.json'), 'utf8'));
if (metadata.name !== 'hash-wasm' || metadata.version !== '4.12.0') throw new Error('This reviewed adapter expects hash-wasm@4.12.0.');
fs.mkdirSync(output, { recursive: true });
const require = createRequire(import.meta.url);
const { argon2id } = require(path.join(packageRoot, 'dist/index.umd.js'));
let library = fs.readFileSync(path.join(packageRoot, 'dist/index.esm.js'), 'utf8');
const librarySha256 = createHash('sha256').update(library).digest('hex');
for (const name of ['argon2', 'blake2b']) {
  const match = library.match(new RegExp('var name(\\$[a-z]+)? = "' + name + '";'));
  if (!match) throw new Error('Expected embedded WASM module is absent: ' + name);
  const suffix = (match[1] || '').replace('$', '\\$');
  const data = library.match(new RegExp('var data' + suffix + ' = "([^"]+)";'))?.[1];
  if (!data) throw new Error('Expected embedded WASM data is absent: ' + name);
  fs.writeFileSync(path.join(output, name + '.wasm'), Buffer.from(data, 'base64'));
}
// Workers prohibit runtime WASM compilation. Replace only module loading with
// deployment-time imports. Hashing code, parameters and memory work are intact.
const compilation = 'const promise = WebAssembly.compile(asm);';
if (library.split(compilation).length !== 2) throw new Error('Unexpected library compilation adapter shape.');
library = "import argonModule from './argon2.wasm';\nimport blakeModule from './blake2b.wasm';\n" + library.replace(compilation,
  `const promise = Promise.resolve(binary.name === 'argon2' ? argonModule : binary.name === 'blake2b' ? blakeModule : (() => {throw new Error('Unexpected WASM module')})());`);
fs.writeFileSync(path.join(output, 'hash-wasm-workers.mjs'), library);
fs.copyFileSync(path.join(packageRoot, 'LICENSE'), path.join(output, 'hash-wasm-LICENSE'));
const access = randomBytes(32).toString('hex');
fs.writeFileSync(path.join(output, 'access'), access, { mode: 0o600 });
const password = 'Synthetic benchmark only - not a user password', salt = 'Synthetic salt!!';
const cases = {
  'pbkdf2-sha256': { hash: 'SHA-256', iterations: 600000 },
  'pbkdf2-sha512': { hash: 'SHA-512', iterations: 220000 },
};
const expected = {};
for (const [name, options] of Object.entries(cases)) {
  expected[name] = pbkdf2Sync(password, salt, options.iterations, 32, options.hash.toLowerCase().replace('-', '')).toString('hex');
}
expected.argon2id = await argon2id({ password, salt, parallelism: 1, iterations: 2, memorySize: 19456, hashLength: 32, outputType: 'hex' });
fs.writeFileSync(path.join(output, 'expected.json'), JSON.stringify(expected, null, 2));
fs.writeFileSync(path.join(output, 'provenance.json'), JSON.stringify({ package: 'hash-wasm', version: metadata.version, librarySha256, hostingAdapter: 'static WASM imports only' }, null, 2));
fs.writeFileSync(path.join(output, 'worker.mjs'), `import {argon2id} from './hash-wasm-workers.mjs';
const encoder=new TextEncoder(),password=${JSON.stringify(password)},salt=${JSON.stringify(salt)};
const cases=${JSON.stringify(cases)},expected=${JSON.stringify(expected)};
const hex=bytes=>Array.from(new Uint8Array(bytes),b=>b.toString(16).padStart(2,'0')).join('');
export default {async fetch(request){
 if(request.method!=='POST'||request.headers.get('authorization')!=='Bearer ${access}')return new Response('Not found',{status:404});
 const test=new URL(request.url).pathname.slice(1);let result;
 try{
  let actual;
  if(test==='argon2id')actual=await argon2id({password,salt,parallelism:1,iterations:2,memorySize:19456,hashLength:32,outputType:'hex'});
  else if(cases[test]){
   const key=await crypto.subtle.importKey('raw',encoder.encode(password),'PBKDF2',false,['deriveBits']);
   actual=hex(await crypto.subtle.deriveBits({name:'PBKDF2',...cases[test],salt:encoder.encode(salt)},key,256));
  }else if(test==='session-control'){
   actual=hex(await crypto.subtle.digest('SHA-256',crypto.getRandomValues(new Uint8Array(32))));
  }else return new Response('Not found',{status:404});
  result={test,ok:true,parity:test==='session-control'?actual.length===64:actual===expected[test]};
 }catch(error){result={test,ok:false,error:String(error.message)};}
 console.log(JSON.stringify({event:'password-feasibility',...result}));
 return Response.json(result,{headers:{'Cache-Control':'no-store'}});
}};`);
fs.writeFileSync(path.join(output, 'wrangler.toml'), `name = "roster-password-benchmark-oct15"
main = "worker.mjs"
compatibility_date = "2026-10-15"
workers_dev = true
preview_urls = false
[observability]
enabled = true
head_sampling_rate = 1
`);
fs.writeFileSync(path.join(output, 'run.mjs'), `import fs from 'node:fs';
const base=process.argv[2],url=new URL(base);if(url.protocol!=='https:'||url.origin!==base||!/^[a-z0-9-]+\\.[a-z0-9-]+\\.workers\\.dev$/.test(url.hostname))throw Error('Provide the isolated workers.dev origin, without a trailing slash.');
const key=fs.readFileSync(new URL('./access',import.meta.url),'utf8').trim(),results=[];
for(let repeat=0;repeat<3;repeat++)for(const test of ['session-control','pbkdf2-sha256','pbkdf2-sha512','argon2id']){
 const response=await fetch(base+'/'+test,{method:'POST',headers:{authorization:'Bearer '+key},signal:AbortSignal.timeout(30000)});
 const raw=await response.text();let data;try{data=JSON.parse(raw);}catch{data={error:raw.slice(0,250)};}
 const result={repeat,test,status:response.status,...data};results.push(result);console.log(JSON.stringify(result));
}
fs.writeFileSync(new URL('./results.json',import.meta.url),JSON.stringify(results,null,2));
`);
console.log('Prepared dummy-only native benchmark. No deployment, production bindings or application changes.');
