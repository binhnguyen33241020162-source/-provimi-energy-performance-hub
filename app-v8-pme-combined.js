(function(){
const INPUT_MODEL='pme-combined-v1';
const PME_SOURCES={
  hm:{label:'HM310A',aliases:['processq14hm310a','hm310a']},
  ma:{label:'Mixer A',aliases:['processq15mixer525a','mixer525a']},
  mb:{label:'Mixer B',aliases:['processq16mixer525b','mixer525b']},
  comp:{label:'Air Compressor',aliases:['msbdbac']},
  plant:{label:'Whole Plant',aliases:['msbq12500a']}
};
function nrm(v){return String(v??'').toLowerCase().replace(/[^a-z0-9]/g,'')}
function med(a){if(!a.length)return null;let b=a.slice().sort((x,y)=>x-y);let m=Math.floor(b.length/2);return b.length%2?b[m]:(b[m-1]+b[m])/2}
function findPmeHeader(rows){
  let best=null;
  for(let r=0;r<Math.min(rows.length,40);r++){
    let row=rows[r]||[],ts=row.findIndex(v=>nrm(v).includes('timestamp'));
    if(ts<0)continue;
    let score=0;
    for(let def of Object.values(PME_SOURCES)) if(row.some(v=>def.aliases.some(a=>nrm(v).includes(a)))) score++;
    if(!best||score>best.score)best={row:r,ts,score};
  }
  return best;
}
function sourceColumn(headers,def){
  let exact=[];
  headers.forEach((v,i)=>{
    let h=nrm(v);
    if(!h)return;
    let source=def.aliases.some(a=>h.includes(a));
    let into=h.includes('realenergyintotheload');
    let cumulative=into&&!h.includes('interval')&&!h.includes('outoftheload');
    if(source&&cumulative) exact.push(i);
  });
  if(exact.length)return exact[0];
  return -1;
}
function pointsToDaily(points){
  points.sort((a,b)=>a.d-b.d);
  if(points.length<2)return{daily:{},meta:{first:points[0]?.d||null,last:points[0]?.d||null,gaps:0,resets:0,interval:null,points:points.length}};
  let ds=[];
  for(let i=1;i<points.length;i++){
    let m=(points[i].d-points[i-1].d)/60000;
    if(m>0&&m<=120)ds.push(m);
  }
  let interval=med(ds)||15,gapThreshold=Math.max(22.5,interval*1.55),daily={},gaps=0,resets=0;
  for(let i=1;i<points.length;i++){
    let p0=points[i-1],p1=points[i],mins=(p1.d-p0.d)/60000,dv=p1.v-p0.v;
    if(mins>gapThreshold){gaps++;continue}
    if(dv<0){resets++;continue}
    if(!isFinite(dv))continue;
    let key=normalizeDateKey(p1.d);daily[key]=(daily[key]||0)+dv;
  }
  return{daily,meta:{first:points[0].d,last:points[points.length-1].d,gaps,resets,interval,points:points.length}};
}
window.pmeCombinedFromWorkbook=function(wb){
  let bestSheet=null,bestHeader=null,bestRows=null;
  for(let sn of wb.SheetNames){
    let rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:null});
    let h=findPmeHeader(rows);
    if(h&&(!bestHeader||h.score>bestHeader.score)){bestSheet=sn;bestHeader=h;bestRows=rows}
  }
  if(!bestHeader||bestHeader.score<3)throw new Error('Không tìm thấy header PME tổng hợp có Timestamp và các meter chuẩn.');
  let headers=bestRows[bestHeader.row],cols={},missing=[];
  for(let [k,def] of Object.entries(PME_SOURCES)){
    cols[k]=sourceColumn(headers,def);
    if(cols[k]<0)missing.push(def.label);
  }
  if(missing.length)throw new Error('PME thiếu hoặc không nhận diện được: '+missing.join(', '));
  let pts={plant:[],hm:[],ma:[],mb:[],comp:[]};
  for(let r=bestHeader.row+1;r<bestRows.length;r++){
    let row=bestRows[r]||[],dt=dateFromAny(row[bestHeader.ts]);
    if(!dt)continue;
    for(let k of Object.keys(pts)){
      let v=row[cols[k]],num=typeof v==='number'?v:Number(String(v??'').replace(/,/g,''));
      if(isFinite(num))pts[k].push({d:dt,v:num});
    }
  }
  let energy={},meta={sheet:bestSheet,headerRow:bestHeader.row+1,columns:{}};
  for(let [k,def] of Object.entries(PME_SOURCES)){
    let res=pointsToDaily(pts[k]);energy[k]=res.daily;meta[k]=res.meta;meta.columns[k]={index:cols[k],header:String(headers[cols[k]]||'')};
  }
  return{energy,meta};
};
function pmeQualityText(k){
  let m=S.pmeMeta?.[k];
  if(!m)return'';
  let parts=[];
  if(m.gaps)parts.push(m.gaps+' gap');
  if(m.resets)parts.push(m.resets+' reset');
  return parts.length?parts.join(' · '):'Clean';
}
window.buildImport=function(){
  document.getElementById('fileCards').innerHTML=
    fileCard(1,'pme','PME Consolidated Energy','Whole Plant · HM310A · Mixer A · Mixer B · Air Compressor',false)+
    fileCard(2,'production','Production Volume','Daily production · Soy47 · Soy50 · Mixer A · Mixer B',true);
  document.querySelectorAll('#fileCards input[type=file]').forEach(inp=>inp.onchange=()=>{
    if(!inp.files.length)return;
    document.getElementById('n'+inp.id.slice(1)).textContent=inp.files[0].name;
    readWorkbook(inp.files[0],wb=>{
      if(inp.dataset.prod==='1'){
        S.production.d=productionFromWorkbook(wb);
      }else{
        let res=pmeCombinedFromWorkbook(wb);
        for(let k of ['plant','hm','ma','mb','comp'])S.energy[k]={d:res.energy[k]};
        S.pmeMeta=res.meta;
      }
      S.inputModel=INPUT_MODEL;save();renderValidation();renderAll();
    });
  });
  bindOnlineProduction();
};
window.renderValidation=function(){
  let c=currentKey(),rows='';
  let mapped=S.pmeMeta?.columns?Object.keys(S.pmeMeta.columns).length:0;
  rows+='<tr><td><b>PME Combined</b></td><td>'+mapped+'/5 meters</td><td><span class="badge '+(mapped===5?'ok':'wait')+'">'+(mapped===5?'Mapped':'Awaiting')+'</span></td></tr>';
  for(let k of ['plant','hm','ma','mb','comp']){
    let v=energy(k,c),ok=Object.keys(S.energy[k]?.d||{}).length>0,q=pmeQualityText(k);
    rows+='<tr><td>'+LABELS[k]+'</td><td>'+(v!=null?fmt(v)+' kWh':'—')+'</td><td><span class="badge '+(ok?'ok':'wait')+'">'+(ok?(q||'Ready'):'Awaiting')+'</span></td></tr>';
  }
  let pk=prod(c);rows+='<tr><td>Production Volume</td><td>'+(pk?fmt(den(pk,'comp'))+' t':'—')+'</td><td><span class="badge '+(pk?'ok':'wait')+'">'+(pk?'Ready':'Awaiting')+'</span></td></tr>';
  let sw=typeof syncWindow==='function'?syncWindow():null;
  rows+='<tr><td><b>SYNC WINDOW</b></td><td>'+(sw?sw.start+' → '+sw.end:'—')+'</td><td><span class="badge '+(sw?'ok':'wait')+'">'+(sw?'Aligned':'Needs both files')+'</span></td></tr>';
  document.getElementById('validationRows').innerHTML=rows;
};
function applyTwoFileLabels(){
  let btn=document.querySelector('.importBtn');if(btn)btn.textContent='+ Import 2 files';
  let h=document.querySelector('#import h1');if(h)h.textContent='Import Center · 2 Standard Files';
  let sub=document.querySelector('#import .titleRow .subtitle');if(sub)sub.textContent='1 PME consolidated energy + 1 Production Volume. Meter columns may move; mapping follows source names.';
  let st=document.querySelector('#import .titleRow .status');if(st)st.textContent='2 required inputs';
}
if(S.inputModel!==INPUT_MODEL){
  S.energy={plant:{d:{}},hm:{d:{}},ma:{d:{}},mb:{d:{}},comp:{d:{}}};
  S.production={d:{}};S.pmeMeta=null;S.inputModel=INPUT_MODEL;save();
}
applyTwoFileLabels();buildImport();renderValidation();renderAll();
})();