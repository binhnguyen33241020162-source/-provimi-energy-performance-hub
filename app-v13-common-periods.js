(function(){
const baseCurrentKey=window.currentKey;
const basePeriodKeys=window.periodKeys;
const baseRenderValidation=window.renderValidation;
const baseRenderAll=window.renderAll;

function uniqSorted(arr){return [...new Set(arr)].sort();}

function energyWeekKeys(){
  const raw=S.pmeRaw?.plant||{};
  if(!raw||!Object.keys(raw).length)return[];
  const keys=uniqSorted(Object.keys(raw).map(Number).filter(Number.isFinite).map(t=>wKey(new Date(t))));
  return keys.filter(k=>typeof commonPeriodEndpoints==='function'&&commonPeriodEndpoints(k));
}

function productionWeekKeys(){
  if(S.production?.weekly&&Object.keys(S.production.weekly).length){
    return uniqSorted(Object.keys(S.production.weekly));
  }
  const keys=[];
  for(const ds of Object.keys(S.production?.d||{})){
    const d=new Date(ds+'T00:00:00');
    if(!isNaN(d))keys.push(wKey(d));
  }
  return uniqSorted(keys);
}

function commonWeekKeys(){
  const energy=new Set(energyWeekKeys());
  return productionWeekKeys().filter(k=>energy.has(k)).sort();
}

window.energyWeekKeys=energyWeekKeys;
window.productionWeekKeys=productionWeekKeys;
window.commonWeekKeys=commonWeekKeys;

window.currentKey=function(){
  if(S.mode!=='week')return baseCurrentKey();
  const common=commonWeekKeys();
  return common.length?common.at(-1):baseCurrentKey();
};

window.periodKeys=function(){
  if(S.mode!=='week')return basePeriodKeys();
  const common=commonWeekKeys();
  if(!common.length)return basePeriodKeys();
  return common.slice(-6);
};

function rangeText(keys){
  if(!keys.length)return'—';
  return keys.length===1?keyLabel(keys[0]):keyLabel(keys[0])+' → '+keyLabel(keys.at(-1));
}

window.renderValidation=function(){
  baseRenderValidation();
  if(S.mode!=='week')return;
  const e=energyWeekKeys(),p=productionWeekKeys(),c=commonWeekKeys();
  const body=document.getElementById('validationRows');if(!body)return;
  const tr=document.createElement('tr');
  tr.innerHTML='<td><b>COMMON WEEK RANGE</b><div style="font-size:10px;color:#7a8792;margin-top:3px">Energy '+rangeText(e)+' · Production '+rangeText(p)+'</div></td><td>'+(c.length?rangeText(c):'—')+'</td><td><span class="badge '+(c.length?'ok':'wait')+'">'+(c.length?c.length+' overlapping week'+(c.length>1?'s':''):'No overlapping week')+'</span></td>';
  body.appendChild(tr);
};

window.renderAll=function(){
  baseRenderAll();
  if(S.mode!=='week')return;
  const keys=periodKeys(),c=currentKey();
  const title=document.getElementById('overviewTrendTitle');
  const meta=document.getElementById('overviewTrendMeta');
  const period=document.getElementById('periodText');
  if(title)title.textContent=(keys.length||6)+'-Week Plant SEC Trend';
  if(meta)meta.textContent=keys.length?keys.map(keyLabel).join(' · '):'No common weeks';
  if(period&&keys.length)period.textContent='Weekly · '+keyLabel(c)+' · common range '+rangeText(keys);
};

renderValidation();
renderAll();
})();