// React only to the browser's autofill state, never to ordinary password typing.
// No polling, credential copies, provider calls or background activity.
export function installPasswordAutofillSubmit({form,password,eligible,pending,schedule=setTimeout,cancel=clearTimeout}) {
  let timer=null,consumed=false;
  function autofilled(){
    try{return password.matches(':autofill')||password.matches(':-webkit-autofill');}
    catch{try{return password.matches(':-webkit-autofill');}catch{return false;}}
  }
  function stop(){if(timer!==null)cancel(timer);timer=null;}
  function notice(event){
    if(event.type==='animationstart'&&event.animationName!=='login-password-autofilled')return;
    if(event.type==='input'&&['insertText','deleteContentBackward','deleteContentForward'].includes(event.inputType))return;
    if(consumed||pending()||!eligible()||!password.value||!autofilled())return;
    stop();
    // Let paired username filling and Safari's native submission finish first.
    timer=schedule(()=>{
      timer=null;
      if(consumed||pending()||!eligible()||!password.value||!autofilled()||!form.checkValidity())return;
      consumed=true;
      form.requestSubmit();
    },120);
  }
  for(const type of ['input','change','animationstart'])password.addEventListener(type,notice);
  form.addEventListener('submit',()=>{consumed=true;stop();});
  return {reset(){stop();consumed=false;}};
}
