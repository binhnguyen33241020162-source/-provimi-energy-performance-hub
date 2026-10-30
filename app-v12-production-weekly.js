(function(){
const baseProductionFromWorkbook=window.productionFromWorkbook;
function n(v){return String(v??'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'')}
function val(v){
  if(v===null||v===undefined)return 0;
  if(typeof v==='number')return isFinite(v)?v:0;
  let s=String(v).trim();
  if(!s||s==='-'||s==='–'||s==='—')return 0;
  s=s.replace(/,/g,'');
  let x=Number(s);return isFinite(x)?x:0;
}
function weekParts(v){
  let s=String(v??'').trim();
  let m=s.match(/(20\d{2})\D+(\d{1,2})/);
  if(!m)return null;
  let y=+m[1],w=+m[2];if(w<1||w>53)return null;
  return{y,w,key:y+'-W'+String(w).padStart(2,'0')};
}
function isoMonday(y,w){
  const jan4=new Date(y,0,4,0,0,0,0),day=jan4.getDay()||7;
  const d=new Date(y,0,4-(day-1)+(w-1)*7,0,0,0,0);
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
}
function findHeader(rows,need){
  for(let r=0;r<Math.min(rows.length,40);r++){
    let ns=(rows[r]||[]).map(n),week=ns.findIndex(x=>x==='weekofyear'||x.includes('weekofyear'));
    if(week<0)continue;
    let cols={week};for(const [k,aliases] of Object.entries(need)){let c=ns.findIndex(x=>aliases.some(a=>x===a||x.includes(a)));if(c>=0)cols[k]=c;}
    if(Object.keys(need).some(k=>cols[k]>=0))return{row:r,cols};
  }
  return null;
}
function parseWeekly(wb){
  let mix=null,htm=null;
  for(const sn of wb.SheetNames){
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:null});
    const snn=n(sn);
    let hMix=findHeader(rows,{ma:['a'],mb:['b']});
    if(hMix&&(snn.includes('mixerenergydata')||(!mix&&hMix.cols.ma>=0&&hMix.cols.mb>=0)))mix={sn,rows,h:hMix};
    let hHtm=findHeader(rows,{s47:['htmsoy47'],s50:['htmsoy50'],premix:['premix']});
    if(hHtm&&(snn.includes('productionvolumedata')||(!htm&&(hHtm.cols.s47>=0||hHtm.cols.s50>=0))))htm={sn,rows,h:hHtm};
  }
  if(!mix&&!htm)return null;
  const byWeek={};
  function ensure(w){return byWeek[w.key]??=( {week:w.key,y:w.y,w:w.w,s47:0,s50:0,ma:0,mb:0} )}
  let rawNumbers=[];
  if(mix){for(let r=mix.h.row+1;r<mix.rows.length;r++){let row=mix.rows[r]||[],w=weekParts(row[mix.h.cols.week]);if(!w)continue;let o=ensure(w);if(mix.h.cols.ma>=0){o.ma=val(row[mix.h.cols.ma]);rawNumbers.push(o.ma)}if(mix.h.cols.mb>=0){o.mb=val(row[mix.h.cols.mb]);rawNumbers.push(o.mb)}}}
  if(htm){for(let r=htm.h.row+1;r<htm.rows.length;r++){let row=htm.rows[r]||[],w=weekParts(row[htm.h.cols.week]);if(!w)continue;let o=ensure(w);if(htm.h.cols.s47>=0){o.s47=val(row[htm.h.cols.s47]);rawNumbers.push(o.s47)}if(htm.h.cols.s50>=0){o.s50=val(row[htm.h.cols.s50]);rawNumbers.push(o.s50)}}}
  let pos=rawNumbers.filter(x=>x>0).sort((a,b)=>a-b),med=pos.length?pos[Math.floor(pos.length/2)]:0;
  const factor=med>1000?0.001:1;
  const d={},weekly={};
  for(const o of Object.values(byWeek)){
    const v={s47:o.s47*factor,s50:o.s50*factor,ma:o.ma*factor,mb:o.mb*factor};
    d[isoMonday(o.y,o.w)]=v;weekly[o.week]=v;
  }
  const weeks=Object.keys(weekly).sort();
  return{d,weekly,meta:{status:weeks.length?'ready':'needs_mapping',format:'weekly_summary',unit:factor===0.001?'kg→t':'t',weeks,sheets:{mixer:mix?.sn||null,production:htm?.sn||null},ignored:['PREMIX'],fields:['A','B','HTMSOY47','HTMSOY50']}};
}
window.productionFromWorkbook=function(wb){
  const weekly=parseWeekly(wb);
  if(weekly&&Object.keys(weekly.d).length)return weekly;
  return baseProductionFromWorkbook(wb);
};

const baseBuildImport=window.buildImport;
window.buildImport=function(){
  baseBuildImport();
  const card=document.querySelector('#fileCards .upload:nth-child(2) .drop');
  if(card){
    const text=card.childNodes[0];
    if(text&&text.nodeType===3)text.textContent='Weekly production · A · B · HTMSOY47 · HTMSOY50 · PREMIX ignored ';
  }
};

if(typeof buildImport==='function')buildImport();
})();