// Isolated browser fixture only: no app/provider requests, staff or saved browser
// profile. Use an installed Playwright runtime; never install a browser here.
import assert from 'node:assert/strict';
import {readFile,mkdir} from 'node:fs/promises';
import {createRequire} from 'node:module';
const require=createRequire(import.meta.url);
const {chromium}=require(process.env.ROSTER_PLAYWRIGHT||'playwright');
const browser=await chromium.launch({headless:true,...(process.env.ROSTER_TEST_BROWSER?{executablePath:process.env.ROSTER_TEST_BROWSER}:{})});
const html=await readFile('public/index.html','utf8');
const dialog=html.slice(html.indexOf('    <dialog id="passkeyDialog"'),html.indexOf('    <div id="appDialog"'));
const fixture=`<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><link rel="stylesheet" href="/static/styles.css"><body data-skin="console"><button id="passkeySignIn" class="hidden">Sign in with a passkey</button><input id="stayLoggedIn" type="checkbox"><details data-passkey-settings hidden><summary>Passkeys</summary><div data-passkey-settings-body></div></details>${dialog}<script type="module">
import {installPasskeyUI} from '/static/passkey-ui.js';
let who='fixture@example.test',reminders=true,keys=[],calls=[],fail=false;
window.fixture={calls,get keys(){return keys;},get reminders(){return reminders;},get who(){return who;},set who(x){who=x;},set fail(x){fail=x;}};
Object.defineProperty(navigator,'credentials',{value:{async create(){return {toJSON(){return {id:'synthetic-only',type:'public-key',response:{clientDataJSON:'e30',attestationObject:'YWJj'}};}};},async get(){throw new DOMException('Cancelled','NotAllowedError');}}});
window.ui=installPasskeyUI({transport:{validated:()=>!!who,options:async()=>({passkeys:true}),async passkeyAccount(action,fields){calls.push({action,fields});if(fail)throw Error('Synthetic provider unavailable');if(action==='status')return Response.json({ok:true,passkeys:keys,remind:reminders&&!keys.length,remindersEnabled:reminders});if(action==='reminders')reminders=fields.enabled;if(action==='startRegistration')return Response.json({ok:true,challengeId:'fixture',options:{challenge:'YWJj',rp:{name:'Fixture',id:'rtc.curiousmind.app'},user:{id:'YWJj',name:who,displayName:'Fixture'},pubKeyCredParams:[{type:'public-key',alg:-7}]}});if(action==='finishRegistration')keys=[{id:'synthetic-only',name:'Fixture key'}];return Response.json({ok:true});}},email:()=>who,readResponse:r=>r.json(),escapeHtml:x=>String(x).replaceAll('&','&amp;').replaceAll('<','&lt;'),confirm:async()=>true,onSignIn(){},status(){}});
await window.ui.mount();await window.ui.loginAvailable();window.ready=true;
</script>`;
try{
 const context=await browser.newContext({viewport:{width:390,height:844},isMobile:true});
 await context.route('**/*',async route=>{
  const url=new URL(route.request().url());
  if(url.origin!=='https://rtc.curiousmind.app')return route.abort();
  if(url.pathname==='/fixture')return route.fulfill({contentType:'text/html',body:fixture});
  if(!['/static/styles.css','/static/passkey-ui.js','/static/passkeys.js'].includes(url.pathname))return route.abort();
  await route.fulfill({contentType:url.pathname.endsWith('.css')?'text/css':'text/javascript',body:await readFile('public'+url.pathname,'utf8')});
 });
 const page=await context.newPage();const errors=[];page.on('pageerror',e=>errors.push(e.message));await page.goto('https://rtc.curiousmind.app/fixture');await page.waitForFunction(()=>window.ready);
 assert.equal(await page.locator('[data-passkey-settings]').isVisible(),true);
 assert.equal(await page.evaluate(()=>fixture.calls.length),0,'mounting a collapsed setting does no provider work');
 await page.evaluate(()=>ui.afterLogin());await page.locator('#passkeyDialog').waitFor({state:'visible'});
 assert.equal(await page.locator('#passkeySetupPassword').isVisible(),false,'immediate setup does not demand a password');
 assert.match(await page.locator('#passkeyDialog').innerText(),/add a passkey later from Account settings/);
 await page.locator('#passkeySuppressReminder').check();await page.locator('#passkeyNotNow').click();await page.locator('#passkeyDialog').waitFor({state:'hidden'});
 assert.equal(await page.evaluate(()=>fixture.reminders),false);
 await page.evaluate(()=>{ui.reset();ui.afterLogin();});await page.waitForTimeout(1600);assert.equal(await page.locator('#passkeyDialog').isVisible(),false,'dismissal survives another login');
 await page.locator('[data-passkey-settings] summary').click();await page.locator('[data-passkey-reminders]').waitFor({state:'visible'});
 assert.equal(await page.locator('[data-passkey-reminders]').isChecked(),false);
 await page.locator('[data-passkey-reminders]').check();assert.equal(await page.evaluate(()=>fixture.reminders),true);
 await page.locator('[data-add-passkey]').click();await page.locator('#passkeyDialog').waitFor({state:'visible'});
 const overflow=await page.locator('#passkeyDialog').evaluate(node=>node.scrollWidth>node.clientWidth);assert.equal(overflow,false,'mobile setup fits one column');
 await mkdir('/tmp/roster-passkey-ui-20261018',{recursive:true});await page.screenshot({path:'/tmp/roster-passkey-ui-20261018/setup-mobile.png',fullPage:true});
 await page.locator('#passkeySetupForm button[type=submit]').click();await page.locator('#passkeyDialog').waitFor({state:'hidden'});
 await page.getByText('Fixture key',{exact:true}).waitFor();assert.equal(await page.evaluate(()=>fixture.keys.length),1);
 await page.evaluate(()=>{ui.reset();ui.afterLogin();});await page.waitForTimeout(1600);assert.equal(await page.locator('#passkeyDialog').isVisible(),false,'a registered passkey prevents reminders');
 await page.evaluate(()=>{ui.reset();fixture.fail=true;ui.afterLogin();});await page.waitForTimeout(1600);assert.equal(await page.locator('#passkeyDialog').isVisible(),false,'provider failure must not show a misleading setup prompt');
 await page.evaluate(()=>{fixture.fail=false;ui.reset();ui.afterLogin();fixture.who='';ui.reset();});await page.waitForTimeout(1600);assert.equal(await page.locator('#passkeyDialog').isVisible(),false,'logout cancels a queued prompt');
 assert.deepEqual(errors,[]);await context.close();
 console.log('Isolated mobile browser passed: optional setup, account-wide dismissal, later Account setup, reminder re-enable, successful add, no reminder with a key, provider failure and logout suppression, no collapsed-panel fetch, no horizontal overflow. Screenshot in /tmp/roster-passkey-ui-20261018. No external requests or real credentials.');
}finally{await browser.close();}
