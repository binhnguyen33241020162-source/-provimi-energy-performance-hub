(function(){
const baseCurrentKey=window.currentKey;
const basePeriodKeys=window.periodKeys;
const baseRenderValidation=window.renderValidation;
const baseRenderAll=window.renderAll;
const baseRenderDevice=window.renderDevice;

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

// Legacy dashboard code still reads keys[5] as current and keys[4] as previous.
// Return a compatibility array for direct-index reads, while map()/slice()/at()
// expose only the real overlapping weeks so charts/tables do not invent periods.
function compatibilityPeriodArray(display){
  if(display.length===6)return display;
  const compat=new Array(6);
  const first=display[0],last=display.at(-1),prev=display.length>1?display.at(-2):last;
  compat[0]=first;
  for(let i=1;i<4;i++)compat[i]=display[Math.min(i,display.length-1)]||last;
  compat[4]=prev;
  compat[5]=last;
  compat.map=function(cb,thisArg){return display.map(cb,thisArg);};
  compat.forEach=function(cb,thisArg){return display.forEach(cb,thisArg);};
  compat.slice=function(...args){return display.slice(...args);};
  compat.at=function(i){return display.at(i);};
  Object.defineProperty(compat,'displayLength',{value:display.length,enumerable:false});
  return compat;
}

window.periodKeys=function(){
  if(S.mode!=='week')return basePeriodKeys();
  const common=commonWeekKeys();
  if(!common.length)return basePeriodKeys();
  return compatibilityPeriodArray(common.slice(-6));
};

function visibleWeekKeys(){
  return commonWeekKeys().slice(-6);
}

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
  const keys=visibleWeekKeys(),c=currentKey();
  const title=document.getElementById('overviewTrendTitle');
  const meta=document.getElementById('overviewTrendMeta');
  const period=document.getElementById('periodText');
  if(title)title.textContent=(keys.length||6)+'-Week Plant SEC Trend';
  if(meta)meta.textContent=keys.length?keys.map(keyLabel).join(' · '):'No common weeks';
  if(period&&keys.length)period.textContent='Weekly · '+keyLabel(c)+' · common range '+rangeText(keys);
};

window.renderDevice=function(k){
  baseRenderDevice(k);
  if(S.mode!=='week')return;
  const keys=visibleWeekKeys(),view=document.getElementById(k);if(!view||!keys.length)return;
  const heroSpan=view.querySelector('.deviceHero > span');
  if(heroSpan)heroSpan.textContent=keys.length+'-period view';
  const headings=[...view.querySelectorAll('.panel h3')];
  const dataHeading=headings.find(h=>/Period Data$/i.test(h.textContent||''));
  if(dataHeading)dataHeading.textContent=keys.length+'-Period Data';
};

renderValidation();
renderAll();
})();