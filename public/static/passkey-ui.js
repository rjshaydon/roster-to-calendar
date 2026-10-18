import {createPasskeyCeremony,passkeySupported,passkeyError} from './passkeys.js';

// Account-only UI; no timers poll Supabase and no roster render is refreshed.
export function installPasskeyUI({transport,email,readResponse,escapeHtml,confirm,onSignIn,status,document:doc=globalThis.document}) {
  const ceremony=createPasskeyCeremony();
  const dialog=doc.querySelector('#passkeyDialog'),form=doc.querySelector('#passkeySetupForm');
  const message=doc.querySelector('#passkeySetupStatus'),password=doc.querySelector('#passkeySetupPassword');
  const suppress=doc.querySelector('#passkeySuppressReminder');
  let actor='',serial=0,prompted='',cache=null,request=null,enabled=false,busy=false,returnFocus=null;
  const current=(value)=>value&&value===email()&&transport.validated();
  async function api(action,fields={}) {
    return readResponse(await transport.passkeyAccount(action,fields),'Could not update passkeys.');
  }
  async function allowed(){const options=await transport.options();enabled=options.passkeys===true&&passkeySupported();return enabled;}
  function feedback(text){message.textContent=text;}
  function close(){serial++;ceremony.cancel();password.value='';password.closest('label').hidden=true;dialog.close();actor='';busy=false;form.querySelector('[type=submit]').disabled=false;returnFocus?.focus?.();}
  function render(root,data){
    if(!root?.isConnected)return;
    root.innerHTML=`<p>Use Face ID, a fingerprint, a device PIN or a security key to sign in. Password sign-in stays available.</p>
      ${data.passkeys.length?`<ul class="passkey-list">${data.passkeys.map(item=>`<li><span>${escapeHtml(item.name)}</span><button type="button" class="button button-secondary" data-remove-passkey="${escapeHtml(item.id)}">Remove</button></li>`).join('')}</ul>`:'<p>No passkeys added yet.</p>'}
      <button type="button" class="button button-secondary" data-add-passkey>Add a passkey</button>
      <label class="toggle"><input type="checkbox" data-passkey-reminders ${data.remindersEnabled?'checked':''}> Remind me to set up a passkey after signing in</label>
      <p data-passkey-settings-status role="status" aria-live="polite"></p>`;
  }
  async function load(root,force=false){
    const who=email();if(!current(who))return;
    if(!force&&cache?.email===who&&Date.now()-cache.at<60000){render(root,cache.data);return;}
    if(!request||request.email!==who){const work=api('status').then(data=>{if(current(who))cache={email:who,at:Date.now(),data};return data;});const own={email:who,work};request=own;void work.finally(()=>{if(request===own)request=null;}).catch(()=>{});}
    try{const data=await request.work;if(current(who))render(root,data);}catch(error){if(current(who)&&root.isConnected)root.textContent=passkeyError(error);}
  }
  async function open(){
    const who=email();if(!current(who)||!await allowed()||!current(who))return;
    actor=who;serial++;suppress.checked=false;password.value='';password.closest('label').hidden=true;
    feedback('');returnFocus=doc.activeElement;dialog.showModal();
    doc.querySelector('#passkeyNotNow').focus();
  }
  async function dismiss(){
    if(busy)return;
    const who=actor,lease=serial;
    if(suppress.checked){busy=true;try{await api('reminders',{enabled:false});if(!current(who)||lease!==serial)return;cache=null;}
      catch(error){feedback(passkeyError(error));return;}finally{busy=false;}}
    close();
  }
  dialog.addEventListener('cancel',event=>{event.preventDefault();void dismiss();});
  doc.querySelector('#passkeyNotNow').addEventListener('click',()=>{void dismiss();});
  form.addEventListener('submit',async event=>{
    event.preventDefault();if(busy||!current(actor))return;
    busy=true;const lease=serial,who=actor,button=form.querySelector('[type=submit]');button.disabled=true;feedback('Preparing passkey…');
    try{
      const start=await api('startRegistration',{password:password.value});password.value='';
      if(lease!==serial||!current(who))return;
      feedback('Follow your device’s passkey prompt.');
      const credential=await ceremony.run(start.options,true);
      if(lease!==serial||!current(who))return;
      await api('finishRegistration',{challengeId:start.challengeId,credential});
      if(lease!==serial||!current(who))return;
      if(suppress.checked)await api('reminders',{enabled:false});
      if(lease!==serial||!current(who))return;
      cache=null;close();
      const root=doc.querySelector('[data-passkey-settings-body]');if(root)void load(root,true);
    }catch(error){
      if(lease!==serial||!current(who))return;
      if(error.code==='auth-passkey-password-required'){password.closest('label').hidden=false;password.focus();}
      feedback(passkeyError(error));
    }finally{if(lease===serial){busy=false;button.disabled=false;}}
  });
  doc.addEventListener('toggle',event=>{
    const section=event.target;if(section.matches?.('[data-passkey-settings]')&&section.open)void load(section.querySelector('[data-passkey-settings-body]'));
  },true);
  doc.addEventListener('change',async event=>{
    const input=event.target;if(!input.matches?.('[data-passkey-reminders]'))return;
    const who=email();input.disabled=true;const desired=input.checked;
    try{await api('reminders',{enabled:desired});if(current(who))cache=null;}
    catch(error){if(current(who)&&input.isConnected){input.checked=!desired;input.closest('[data-passkey-settings-body]').querySelector('[data-passkey-settings-status]').textContent=passkeyError(error);}}
    finally{input.disabled=false;}
  });
  doc.addEventListener('click',async event=>{
    if(event.target.closest?.('[data-add-passkey]')){void open();return;}
    const remove=event.target.closest?.('[data-remove-passkey]');if(!remove)return;
    const who=email(),root=remove.closest('[data-passkey-settings-body]');remove.disabled=true;
    try{
      if(!await confirm({title:'Remove passkey?',message:'This passkey will no longer sign you in. Your password will still work.',confirmLabel:'Remove passkey'})||!current(who))return;
      try{await api('remove',{id:remove.dataset.removePasskey});}
      catch(error){if(error.code!=='auth-reauthentication-required')throw error;
        const proof=await confirm({title:'Confirm it’s you',message:'Confirm your password to remove this sign-in method.',confirmLabel:'Continue',requirePassword:true,email:who});
        if(!proof||!current(who))return;
        await readResponse(await transport.reauthenticate(proof),'Could not confirm your password.');
        if(!current(who))return;await api('remove',{id:remove.dataset.removePasskey});
      }
      if(current(who)){cache=null;await load(root,true);}
    }catch(error){if(current(who)&&root.isConnected)root.querySelector('[data-passkey-settings-status]').textContent=passkeyError(error);}
    finally{remove.disabled=false;}
  });
  const signIn=doc.querySelector('#passkeySignIn');
  signIn.addEventListener('click',async()=>{
    signIn.disabled=true;
    try{
      const remember=doc.querySelector('#stayLoggedIn').checked;
      const response=await transport.passkeySignIn(options=>ceremony.run(options),remember);
      const data=await readResponse(response,'Could not sign in with a passkey.');
      await onSignIn(data.email,remember);
    }catch(error){status(passkeyError(error),true);}finally{signIn.disabled=false;}
  });
  return {
    async mount(){
      const who=email(),root=doc.querySelector('[data-passkey-settings]');
      if(root){const permitted=await allowed();if(root.isConnected&&current(who))root.hidden=!permitted;}
    },
    async loginAvailable(){signIn.classList.toggle('hidden',!await allowed());},
    afterLogin(){
      const who=email();if(!current(who)||prompted===who)return;prompted=who;
      // One deferred check per explicit login. Calendar paint wins; no await
      // in login/hydration, no request on a normal session resume or refresh.
      setTimeout(()=>{void(async()=>{
        if(!current(who)||!await allowed()||!current(who))return;
        const data=await api('status',{prompt:true});
        if(current(who)&&data.remind&&!doc.hidden&&!doc.querySelector('.modal:not(.hidden)'))await open();
      })().catch(()=>{});},1500);
    },
    reset(){prompted='';cache=null;request=null;close();},
  };
}
