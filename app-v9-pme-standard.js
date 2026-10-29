(function(){
const INPUT_MODEL_V9='pme-standard-v3';
const SOURCES={
  hm:{label:'HM310A',aliases:['processq14hm310a','hm310a']},
  ma:{label:'Mixer A',aliases:['processq15mixer525a','mixer525a']},
  mb:{label:'Mixer B',aliases:['processq16mixer525b','mixer525b']},
  comp:{label:'Air Compressor',aliases:['msbdbac']},
  plant:{label:'Whole Plant',aliases:['msbq12500a']}
};
const KEYS=['plant','hm','ma','mb','comp'];
function norm(v){return String(v??'').toLowerCase().replace(/[^a-z0-9]/g,'')}
function bucket15(d){let ms=d.getTime(),q=15*60*1000;return new Date(Math.round(ms/q)*q)}
function findHeader(rows){let best=null;for(let r=0;r<Math.min(rows.length,50);r++){let row=rows[r]||[],ts=row.findIndex(v=>norm(v).includes('timestamp'));if(ts<0)continue;let score=0;for(let def of Object.values(SOURCES))if(row.some(v=>def.aliases.some(a=>norm(v).includes(a))))score++;if(!best||score>best.score)best={row:r,ts,score}}return best}
function meterCol(headers,def){let hits=[];headers.forEach((v,i)=>{let h=norm(v);if(!h)return;let src=def.aliases.some(a=>h.includes(a));let metric=h.includes('realenergyintotheload');let bad=h.includes('interval')||h.includes('outoftheload')||h.includes('reactive')||h.includes('apparent')||h.includes('kvar')||h.includes('kva');if(src&&metric&&!bad)hits.push(i)});return hits.length?hits[0]:-1}
function num(v){if(typeof v==='number'&&isFinite(v))return v;let n=Number(String(v??'').replace(/,/g,''));return isFinite(n)?n:null}
function dayKey(d){return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0')}
function plusDays(ds,n){let d=new Date(ds+'T00:00:00');d.setDate(d.getDate()+n);return dayKey(d)}
function parsePME(wb){
  let best=null;
  for(let sn of wb.SheetNames){let rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:null});let h=findHeader(rows);if(h&&(!best||h.score>best.h.score))best={sn,rows,h}}
  if(!best||best.h.score<5)throw new Error('Không tìm thấy đủ 5 meter chuẩn trong PME consolidated report.');
  let headers=best.rows[best.h.row],cols={},missing=[];for(let [k,def] of Object.entries(SOURCES)){cols[k]=meterCol(headers,def);if(cols[k]<0)missing.push(def.label)}
  if(missing.length)throw new Error('PME thiếu hoặc không nhận diện được: '+missing.join(', '));
  let raw={plant:{},hm:{},ma:{},mb:{},comp:{}},duplicates=0;
  for(let r=best.h.row+1;r<best.rows.length;r++){
    let row=best.rows[r]||[],dt=dateFromAny(row[best.h.ts]);if(!dt)continue;let b=bucket15(dt),t=String(b.getTime());
    for(let k of KEYS){let v=num(row[cols[k]]);if(v==null)continue;if(Object.prototype.hasOwnProperty.call(raw[k],t))duplicates++;raw[k][t]=v}
  }
  let common=null;for(let k of KEYS){let s=new Set(Object.keys(raw[k]));common=common==null?s:new Set([...common].filter(x=>s.has(x)))}
  let commonTimes=[...(common||[])].map(Number).sort((a,b)=>a-b);if(commonTimes.length<2)throw new Error('Không có đủ timestamp chung giữa 5 meter.');
  let commonSet=new Set(commonTimes.map(String));
  let allDays=[...new Set(commonTimes.map(t=>dayKey(new Date(t))))].sort(),energy={plant:{},hm:{},ma:{},mb:{},comp:{}},completeDays=[],excludedDays=[],resets={plant:0,hm:0,ma:0,mb:0,comp:0};
  for(let ds of allDays){let t0=new Date(ds+'T00:00:00').getTime(),t1=new Date(plusDays(ds,1)+'T00:00:00').getTime();if(!commonSet.has(String(t0))||!commonSet.has(String(t1))){excludedDays.push(ds);continue}let ok=true,vals={};for(let k of KEYS){let dv=raw[k][String(t1)]-raw[k][String(t0)];if(!isFinite(dv)||dv<0){if(dv<0)resets[k]++;ok=false;break}vals[k]=dv}if(!ok){excludedDays.push(ds);continue}completeDays.push(ds);for(let k of KEYS)energy[k][ds]=vals[k]}
  let meta={sheet:best.sn,headerRow:best.h.row+1,columns:{},duplicates,commonPoints:commonTimes.length,completeDays:completeDays.length,excludedDays,resets,firstCommon:new Date(commonTimes[0]),lastCommon:new Date(commonTimes.at(-1))};
  for(let [k,def] of Object.entries(SOURCES))meta.columns[k]={index:cols[k],header:String(headers[cols[k]]||'')};
  return{energy,raw,meta};
}
function exactValidDays(p){let sets=KEYS.map(k=>new Set(Object.keys(S.energy[k]?.d||{}).filter(ds=>periodOfDateKey(ds)===p)));sets.push(new Set(Object.keys(S.production.d||{}).filter(ds=>periodOfDateKey(ds)===p)));if(sets.some(s=>!s.size))return[];return [...sets[0]].filter(ds=>sets.every(s=>s.has(ds))).sort()}
energy=function(k,p){let days=exactValidDays(p);if(!days.length)return null;return days.reduce((a,ds)=>a+(+S.energy[k].d[ds]||0),0)};
prod=function(p){let days=exactValidDays(p);if(!days.length)return null;let out={s47:0,s50:0,ma:0,mb:0};for(let ds of days){let r=S.production.d[ds]||{};out.s47+=+r.s47||0;out.s50+=+r.s50||0;out.ma+=+r.ma||0;out.mb+=+r.mb||0}return out};
sec=function(k,p){let e=energy(k,p),pr=prod(p),d=den(pr,k);return e!=null&&d>0?e/d:null};
function hierarchy(p){let plant=energy('plant',p);if(!(plant>0))return{ok:false,reason:'No Whole Plant data'};let vals={},tracked=0;for(let k of ['hm','ma','mb','comp']){let e=energy(k,p);vals[k]=e;if(e!=null)tracked+=e;if(e!=null&&e>plant*1.001)return{ok:false,reason:LABELS[k]+' > Whole Plant',plant,vals,tracked}}if(tracked>plant*1.001)return{ok:false,reason:'Tracked loads > Whole Plant',plant,vals,tracked};return{ok:true,plant,vals,tracked,other:plant-tracked}}
const baseRenderAll=renderAll;
renderAll=function(){baseRenderAll();let p=currentKey(),h=hierarchy(p);if(!h.ok&&h.plant){document.getElementById('trackedShare').textContent='Check';document.getElementById('shares').innerHTML='<div class="decision bad"><strong>Energy hierarchy warning</strong><div class="subtitle">'+h.reason+'. Kiểm tra mapping/coverage trước khi dùng tỷ trọng.</div></div>'}else if(h.ok){let html='';for(let k of ['hm','ma','mb','comp']){let e=h.vals[k];if(e!=null)html+=shareRow(LABELS[k],e/h.plant*100,COLORS[k])}html+=shareRow('Other Loads',h.other/h.plant*100,COLORS.other);document.getElementById('shares').innerHTML=html;document.getElementById('trackedShare').textContent=fmt(h.tracked/h.plant*100,1)+'%'}};
const baseRenderDevice=renderDevice;
renderDevice=function(k){baseRenderDevice(k);let p=currentKey(),h=hierarchy(p);if(!h.ok){let view=document.getElementById(k);let shareCard=[...view.querySelectorAll('.kpi')].find(x=>x.querySelector('.lab')?.textContent==='Plant share');if(shareCard){shareCard.querySelector('.val').textContent='Check';shareCard.querySelector('.delta').textContent=h.reason}}};
window.buildImport=function(){document.getElementById('fileCards').innerHTML=fileCard(1,'pme','PME Consolidated Energy','Whole Plant · HM310A · Mixer A · Mixer B · Air Compressor',false)+fileCard(2,'production','Production Volume','Daily production · Soy47 · Soy50 · Mixer A · Mixer B',true);document.querySelectorAll('#fileCards input[type=file]').forEach(inp=>inp.onchange=()=>{if(!inp.files.length)return;document.getElementById('n'+inp.id.slice(1)).textContent=inp.files[0].name;readWorkbook(inp.files[0],wb=>{if(inp.dataset.prod==='1'){S.production.d=productionFromWorkbook(wb).d}else{let res=parsePME(wb);for(let k of KEYS)S.energy[k]={d:res.energy[k]};S.pmeRaw=res.raw;S.pmeMeta=res.meta}S.inputModel=INPUT_MODEL_V9;save();renderValidation();renderAll()})});bindOnlineProduction()};
window.renderValidation=function(){let p=currentKey(),rows='',mapped=S.pmeMeta?.columns?Object.keys(S.pmeMeta.columns).length:0,days=exactValidDays(p),h=hierarchy(p);rows+='<tr><td><b>PME Combined</b></td><td>'+mapped+'/5 meters</td><td><span class="badge '+(mapped===5?'ok':'wait')+'">'+(mapped===5?'Mapped':'Awaiting')+'</span></td></tr>';for(let k of KEYS){let v=energy(k,p);rows+='<tr><td>'+LABELS[k]+'</td><td>'+(v!=null?fmt(v)+' kWh':'—')+'</td><td><span class="badge '+(v!=null?'ok':'wait')+'">'+(v!=null?'Aligned':'Awaiting')+'</span></td></tr>'}let pp=prod(p);rows+='<tr><td>Production Volume</td><td>'+(pp?fmt(den(pp,'comp'))+' t':'—')+'</td><td><span class="badge '+(pp?'ok':'wait')+'">'+(pp?'Aligned':'Awaiting')+'</span></td></tr>';rows+='<tr><td><b>VALID FULL DAYS</b></td><td>'+days.length+(days.length?' · '+days[0]+' → '+days.at(-1):'')+'</td><td><span class="badge '+(days.length?'ok':'wait')+'">'+(days.length?'Same days for 5 meters + production':'Needs aligned days')+'</span></td></tr>';if(S.pmeMeta){rows+='<tr><td>PME Data Quality</td><td>'+S.pmeMeta.commonPoints+' common points · '+S.pmeMeta.completeDays+' complete days</td><td><span class="badge '+(S.pmeMeta.excludedDays?.length?'wait':'ok')+'">'+(S.pmeMeta.excludedDays?.length?S.pmeMeta.excludedDays.length+' partial/excluded day(s)':'Clean boundaries')+'</span></td></tr>'}rows+='<tr><td><b>ENERGY SHARE CHECK</b></td><td>'+(h.ok?'Tracked '+fmt(h.tracked/h.plant*100,1)+'% · Other '+fmt(h.other/h.plant*100,1)+'%':'—')+'</td><td><span class="badge '+(h.ok?'ok':'wait')+'">'+(h.ok?'Hierarchy OK':h.reason)+'</span></td></tr>';document.getElementById('validationRows').innerHTML=rows};
if(S.inputModel!==INPUT_MODEL_V9){let keep={mode:S.mode||'week',kpi:S.kpi||{plant:null,hm:null,ma:null,mb:null,comp:null},productionUrl:S.productionUrl||''};S=fresh();S.mode=keep.mode;S.kpi=keep.kpi;S.productionUrl=keep.productionUrl;S.inputModel=INPUT_MODEL_V9;save()}
buildImport();renderValidation();renderAll();
})();