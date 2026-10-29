(function(){
const CLOUD_API='https://provimi-energy-cloud-api.onrender.com';
let cloudLoading=false,saveTimer=null,lastCloudUpdated=null;

function cloudBadge(text,kind){
  let el=document.getElementById('cloudStateBadge');
  if(!el){
    el=document.createElement('span');el.id='cloudStateBadge';el.className='status';el.style.marginLeft='8px';
    const top=document.querySelector('.top');if(top)top.prepend(el);
  }
  el.textContent=text;
  el.style.borderColor=kind==='ok'?'#76c893':kind==='err'?'#e89a9a':'#cbd5df';
  el.style.color=kind==='ok'?'#177245':kind==='err'?'#a33434':'#607080';
}

async function putCloudState(){
  if(cloudLoading)return;
  try{
    cloudBadge('☁ Saving…','wait');
    const r=await fetch(CLOUD_API+'/api/state',{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({state:S})});
    if(!r.ok)throw new Error('HTTP '+r.status);
    const j=await r.json();lastCloudUpdated=j.updated_at||new Date().toISOString();
    cloudBadge('☁ Cloud saved','ok');
  }catch(e){cloudBadge('☁ Cloud offline','err');}
}

const localSave=save;
save=function(){
  localSave();
  if(cloudLoading)return;
  clearTimeout(saveTimer);saveTimer=setTimeout(putCloudState,250);
};

async function loadCloudState(){
  cloudBadge('☁ Connecting…','wait');
  try{
    const r=await fetch(CLOUD_API+'/api/state',{cache:'no-store'});
    if(!r.ok)throw new Error('HTTP '+r.status);
    const j=await r.json();
    if(j.state&&typeof j.state==='object'){
      cloudLoading=true;
      S=j.state;
      localStorage.setItem('provimiEnergyV5',JSON.stringify(S));
      cloudLoading=false;
      lastCloudUpdated=j.updated_at||null;
      if(document.getElementById('modeSelect'))document.getElementById('modeSelect').value=S.mode||'week';
      if(typeof buildImport==='function')buildImport();
      if(typeof renderValidation==='function')renderValidation();
      if(typeof renderAll==='function')renderAll();
      cloudBadge('☁ Cloud synced','ok');
    }else{
      const hasLocal=!!(S.pmeRaw&&Object.keys(S.pmeRaw).length)||!!Object.keys(S.production?.d||{}).length;
      if(hasLocal){await putCloudState();cloudBadge('☁ Cloud initialized','ok');}
      else cloudBadge('☁ Cloud ready','ok');
    }
  }catch(e){
    cloudBadge('☁ Local fallback','err');
  }
}

window.addEventListener('focus',()=>{loadCloudState();});
loadCloudState();
})();
