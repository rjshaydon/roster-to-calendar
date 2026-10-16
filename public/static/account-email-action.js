// Email credentials live only in memory, never queries, storage or HTTP referrers.
export function takeAccountEmailAction(location,history) {
  const fragment=new URLSearchParams(String(location.hash||'').replace(/^#/,''));
  if(!fragment.has('account'))return null;
  history.replaceState({},'',location.pathname+location.search);
  const purpose=fragment.get('account'),id=fragment.get('challenge'),token=fragment.get('token');
  if(!['confirm','recovery'].includes(purpose)||!/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i.test(id||'')||!/^[A-Za-z0-9_-]{32,256}$/.test(token||''))return{invalid:true};
  return{purpose,id,token};
}

export function takeInvitationToken(location,history) {
  const fragment=new URLSearchParams(String(location.hash||'').replace(/^#/,'')),query=new URLSearchParams(location.search||'');
  const token=fragment.get('invite')||query.get('invite');if(!token)return null;
  fragment.delete('invite');query.delete('invite');
  history.replaceState({},'',location.pathname+(query.toString()?'?'+query:'')+(fragment.toString()?'#'+fragment:''));
  return /^[A-Za-z0-9_-]{24,256}$/.test(token)?token:null;
}
