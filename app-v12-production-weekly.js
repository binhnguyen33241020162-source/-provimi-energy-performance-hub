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
  const jan4=new Date(Date.UTC(y,0,4)),day=jan4.getUTCDay()||7;
  const d=new Date(jan4);d.setUTCDate(jan4.getUTCDate()-(day-1)+(w-1)*7);
  return d.getUTCFullYear()+'-'+String(d.getUTCMonth()+1).padStart(2,'0')+'-'+String(d.getUTCDate()).padStart(2,'0');
}
function findHeader(rows,need){
  for(let r=0;r<Math.min(rows.length,40);r++){
    let ns=(rows[r]||[]).map(n),week=ns.findIndex(x=>x==='weekofyear'||x.includes('weekofyear'));
    if(week<0)continue;
    let cols={week},ok=true;
    for(const [k,aliases] of Object.entries(need)){
      let c=ns.findIndex(x=>aliases.some(a=>x===a));
      if(c<0){ok=false;break}cols[k]=c;
    }
    if(ok)return{row:r,cols};
  }
  return null;
}
function getSheet(wb,target){
  const exact=wb.SheetNames.find(sn=>n(sn)===n(target));
  if(exact)return exact;
  return wb.SheetNames.find(sn=>n(sn).includes(n(target)))||null;
}
function parseWeekly(wb){
  const mixName=getSheet(wb,'Mixer_Energy_Data');
  const grinderName=getSheet(wb,'Grinder_Energy_Data');
  if(!mixName||!grinderName)return null;

  const mixRows=XLSX.utils.sheet_to_json(wb.Sheets[mixName],{header:1,raw:true,defval:null});
  const grinderRows=XLSX.utils.sheet_to_json(wb.Sheets[grinderName],{header:1,raw:true,defval:null});
  const mixH=findHeader(mixRows,{ma:['a'],mb:['b']});
  const grinderH=findHeader(grinderRows,{s47:['htmsoy47'],s50:['htmsoy50']});
  if(!mixH||!grinderH)return null;

  const byWeek={};
  function ensure(w){return byWeek[w.key]??={week:w.key,y:w.y,w:w.w,s47:0,s50:0,ma:0,mb:0}}

  for(let r=mixH.row+1;r<mixRows.length;r++){
    const row=mixRows[r]||[],w=weekParts(row[mixH.cols.week]);if(!w)continue;
    const o=ensure(w);o.ma=val(row[mixH.cols.ma]);o.mb=val(row[mixH.cols.mb]);
  }
  for(let r=grinderH.row+1;r<grinderRows.length;r++){
    const row=grinderRows[r]||[],w=weekParts(row[grinderH.cols.week]);if(!w)continue;
    const o=ensure(w);o.s47=val(row[grinderH.cols.s47]);o.s50=val(row[grinderH.cols.s50]);
  }

  const d={},weekly={};
  for(const o of Object.values(byWeek)){
    const v={s47:o.s47/1000,s50:o.s50/1000,ma:o.ma/1000,mb:o.mb/1000};
    d[isoMonday(o.y,o.w)]=v;weekly[o.week]=v;
  }
  const weeks=Object.keys(weekly).sort();
  return{d,weekly,meta:{
    status:weeks.length?'ready':'needs_mapping',format:'weekly_summary',unit:'t',sourceUnit:'kg',conversion:'kg/1000',weeks,
    sheets:{mixer:mixName,grinder:grinderName},ignored:['PREMIX','Production_Volume_Data'],fields:['A','B','HTMSOY47','HTMSOY50']
  }};
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
    if(text&&text.nodeType===3)text.textContent='Weekly production · Mixer_Energy_Data: A/B · Grinder_Energy_Data: HTMSOY47/HTMSOY50 · PREMIX ignored ';
  }
};

if(typeof buildImport==='function')buildImport();
})();