(function(){
const PME_KEYS=['plant','hm','ma','mb','comp'];
function periodBounds(key){
  if(key.includes('-W')){
    const [y,w]=key.split('-W').map(Number);
    const jan4=new Date(y,0,4,0,0,0,0);
    const day=jan4.getDay()||7;
    const start=new Date(y,0,4-(day-1)+(w-1)*7,0,0,0,0);
    const end=new Date(start);end.setDate(end.getDate()+7);
    return{start,end};
  }
  const [y,m]=key.split('-').map(Number);
  return{start:new Date(y,m-1,1,0,0,0,0),end:new Date(y,m,1,0,0,0,0)};
}
function commonPeriodEndpoints(p){
  const b=periodBounds(p), sets=[];
  for(const k of PME_KEYS){
    const raw=S.pmeRaw?.[k]||{};
    const arr=Object.keys(raw).map(Number).filter(t=>t>=b.start.getTime()&&t<=b.end.getTime()).sort((a,b)=>a-b);
    if(!arr.length)return null;
    sets.push(new Set(arr.map(String)));
  }
  let common=[...sets[0]].filter(t=>sets.every(s=>s.has(t))).map(Number).sort((a,b)=>a-b);
  if(common.length<2)return null;
  const first=common[0],last=common.at(-1);
  if(last<=first)return null;
  const tol=20*60*1000;
  const fullStart=Math.abs(first-b.start.getTime())<=tol;
  const fullEnd=Math.abs(last-b.end.getTime())<=tol;
  return{first,last,start:b.start,end:b.end,fullStart,fullEnd,full:fullStart&&fullEnd};
}
function periodEnergyDetail(k,p){
  const ep=commonPeriodEndpoints(p);if(!ep)return null;
  const raw=S.pmeRaw?.[k]||{};
  const sv=Number(raw[String(ep.first)]),ev=Number(raw[String(ep.last)]);
  if(!isFinite(sv)||!isFinite(ev))return null;
  const value=ev-sv;if(!isFinite(value)||value<0)return null;
  return{value,startValue:sv,endValue:ev,first:ep.first,last:ep.last,full:ep.full,fullStart:ep.fullStart,fullEnd:ep.fullEnd};
}
function dsLocal(t){let d=new Date(t);return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0')}
function periodProductionDetail(p){
  const ep=commonPeriodEndpoints(p);if(!ep)return null;
  const firstDay=dsLocal(ep.first),lastDay=dsLocal(ep.last),out={s47:0,s50:0,ma:0,mb:0};let found=false;
  for(const [ds,row] of Object.entries(S.production.d||{})){
    if(periodOfDateKey(ds)!==p||ds<firstDay||ds>lastDay)continue;
    found=true;out.s47+=+row.s47||0;out.s50+=+row.s50||0;out.ma+=+row.ma||0;out.mb+=+row.mb||0;
  }
  return found?{value:out,firstDay,lastDay,full:ep.full}:null;
}
energy=function(k,p){let d=periodEnergyDetail(k,p);return d?d.value:null};
prod=function(p){let d=periodProductionDetail(p);return d?d.value:null};
sec=function(k,p){let e=energy(k,p),pr=prod(p),d=den(pr,k);return e!=null&&d>0?e/d:null};
window.periodEnergyDetail=periodEnergyDetail;
window.commonPeriodEndpoints=commonPeriodEndpoints;

const previousValidation=window.renderValidation;
window.renderValidation=function(){
  const p=currentKey(),ep=commonPeriodEndpoints(p);let rows='';
  const mapped=S.pmeMeta?.columns?Object.keys(S.pmeMeta.columns).length:0;
  rows+='<tr><td><b>PME Combined</b></td><td>'+mapped+'/5 meters</td><td><span class="badge '+(mapped===5?'ok':'wait')+'">'+(mapped===5?'Mapped':'Awaiting')+'</span></td></tr>';
  for(const k of PME_KEYS){
    const d=periodEnergyDetail(k,p);
    const detail=d?fmt(d.startValue)+' → '+fmt(d.endValue):'—';
    rows+='<tr><td>'+LABELS[k]+'</td><td>'+(d?fmt(d.value)+' kWh':'—')+'<div style="font-size:10px;color:#7a8792;margin-top:3px">'+detail+'</div></td><td><span class="badge '+(d?'ok':'wait')+'">'+(d?'End − Start':'Awaiting')+'</span></td></tr>';
  }
  const pd=periodProductionDetail(p),pp=pd?.value||null;
  rows+='<tr><td>Production Volume</td><td>'+(pp?fmt(den(pp,'comp'))+' t':'—')+'</td><td><span class="badge '+(pp?'ok':'wait')+'">'+(pp?'Period sum':'Awaiting')+'</span></td></tr>';
  if(ep){
    const f=new Date(ep.first),l=new Date(ep.last),fmtDt=d=>d.toLocaleString('vi-VN',{day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});
    rows+='<tr><td><b>PERIOD COVERAGE</b></td><td>'+fmtDt(f)+' → '+fmtDt(l)+'</td><td><span class="badge '+(ep.full?'ok':'wait')+'">'+(ep.full?'Full period':'Partial period')+'</span></td></tr>';
  }else rows+='<tr><td><b>PERIOD COVERAGE</b></td><td>—</td><td><span class="badge wait">No common endpoints</span></td></tr>';
  const h=typeof hierarchy==='function'?hierarchy(p):null;
  if(h)rows+='<tr><td><b>ENERGY SHARE CHECK</b></td><td>'+(h.ok?'Tracked '+fmt(h.tracked/h.plant*100,1)+'% · Other '+fmt(h.other/h.plant*100,1)+'%':'—')+'</td><td><span class="badge '+(h.ok?'ok':'wait')+'">'+(h.ok?'Hierarchy OK':h.reason)+'</span></td></tr>';
  document.getElementById('validationRows').innerHTML=rows;
};

const oldProcess=window.processAll;
window.processAll=function(){
  const p=currentKey(),ep=commonPeriodEndpoints(p);
  if(!ep)return alert('Không tìm thấy cùng timestamp đầu/cuối kỳ cho 5 meter PME trong '+keyLabel(p)+'.');
  const pp=periodProductionDetail(p);
  if(!pp)return alert('Không tìm thấy Production Volume cho khoảng ngày của '+keyLabel(p)+'.');
  save();renderAll();renderValidation();showView('dashboard');
};

renderValidation();renderAll();
})();