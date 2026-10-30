(function(){
const baseProductionFromWorkbook=window.productionFromWorkbook;
const baseRenderValidation=window.renderValidation;
function n(v){return String(v??'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'')}
function val(v){
  if(v===null||v===undefined)return 0;
  if(typeof v==='number')return isFinite(v)?v:0;
  let s=String(v).trim();
  if(!s||s==='-'||s==='–'||s==='—')return 0;
  s=s.replace(/,/g,'').replace(/\s/g,'');
  let x=Number(s);return isFinite(x)?x:0;
}
function weekParts(v){
  if(v instanceof Date&&!isNaN(v)){
    const i=isoInfo(v);return{y:i.year,w:i.week,key:i.year+'-W'+String(i.week).padStart(2,'0')};
  }
  if(typeof v==='number'&&isFinite(v)){
    const s=String(Math.trunc(v));
    if(/^20\d{4}$/.test(s)){
      const y=+s.slice(0,4),w=+s.slice(4);if(w>=1&&w<=53)return{y,w,key:y+'-W'+String(w).padStart(2,'0')};
    }
  }
  let s=String(v??'').trim();
  let m=s.match(/(20\d{2})\D*W?\D*(\d{1,2})/i);
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
  for(let r=0;r<Math.min(rows.length,80);r++){
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
function findRoleSheet(wb,role){
  const need=role==='mixer'?{ma:['a'],mb:['b']}:{s47:['htmsoy47'],s50:['htmsoy50']};
  const candidates=[];
  for(const sn of wb.SheetNames){
    // raw:false intentionally preserves Excel display text such as 2024-09 even when source cell is numeric/date-formatted.
    const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:false,defval:null});
    const h=findHeader(rows,need);if(!h)continue;
    const name=n(sn);let score=0;
    if(role==='mixer'&&name.includes('mixer'))score+=20;
    if(role==='grinder'&&name.includes('grinder'))score+=20;
    if(name.includes('energydata')||name.includes('enerydata'))score+=5;
    if(name.includes('productionvolumedata'))score-=30;
    candidates.push({sn,rows,h,score});
  }
  candidates.sort((a,b)=>b.score-a.score);
  return candidates[0]||null;
}
function parseWeekly(wb){
  const mix=findRoleSheet(wb,'mixer');
  const grinder=findRoleSheet(wb,'grinder');
  if(!mix||!grinder){
    return{d:{},weekly:{},meta:{status:'needs_mapping',format:'weekly_summary',reason:!mix&&!grinder?'Không tìm thấy sheet A/B và HTMSOY47/HTMSOY50':!mix?'Không tìm thấy sheet Week_Of_Year + A + B':'Không tìm thấy sheet Week_Of_Year + HTMSOY47 + HTMSOY50',sheets:{mixer:mix?.sn||null,grinder:grinder?.sn||null}}};
  }

  const byWeek={};
  function ensure(w){return byWeek[w.key]??={week:w.key,y:w.y,w:w.w,s47:0,s50:0,ma:0,mb:0}}

  for(let r=mix.h.row+1;r<mix.rows.length;r++){
    const row=mix.rows[r]||[],w=weekParts(row[mix.h.cols.week]);if(!w)continue;
    const o=ensure(w);o.ma=val(row[mix.h.cols.ma]);o.mb=val(row[mix.h.cols.mb]);
  }
  for(let r=grinder.h.row+1;r<grinder.rows.length;r++){
    const row=grinder.rows[r]||[],w=weekParts(row[grinder.h.cols.week]);if(!w)continue;
    const o=ensure(w);o.s47=val(row[grinder.h.cols.s47]);o.s50=val(row[grinder.h.cols.s50]);
  }

  const d={},weekly={};
  for(const o of Object.values(byWeek)){
    const v={s47:o.s47/1000,s50:o.s50/1000,ma:o.ma/1000,mb:o.mb/1000};
    d[isoMonday(o.y,o.w)]=v;weekly[o.week]=v;
  }
  const weeks=Object.keys(weekly).sort();
  return{d,weekly,meta:{
    status:weeks.length?'ready':'needs_mapping',format:'weekly_summary',unit:'t',sourceUnit:'kg',conversion:'kg/1000',weeks,
    startWeek:weeks[0]||null,endWeek:weeks.at(-1)||null,
    sheets:{mixer:mix.sn,grinder:grinder.sn},ignored:['PREMIX','Production_Volume_Data'],fields:['A','B','HTMSOY47','HTMSOY50'],
    reason:weeks.length?'':'Không đọc được giá trị Week_Of_Year hợp lệ'
  }};
}
window.productionFromWorkbook=function(wb){
  const weekly=parseWeekly(wb);
  if(weekly&&Object.keys(weekly.d).length)return weekly;
  return weekly?.meta?.format==='weekly_summary'?weekly:baseProductionFromWorkbook(wb);
};

function installProductionHandler(){
  const inp=document.querySelector('#fileCards input[data-prod="1"]');
  if(!inp)return;
  inp.onchange=()=>{
    if(!inp.files?.length)return;
    const file=inp.files[0];
    const nameEl=document.getElementById('n'+inp.id.slice(1));if(nameEl)nameEl.textContent=file.name;
    readWorkbook(file,wb=>{
      const parsed=parseWeekly(wb);
      S.production={d:parsed.d||{},weekly:parsed.weekly||{},meta:parsed.meta||{}};
      S.importFiles=S.importFiles||{};
      S.importFiles.production={name:file.name,size:file.size,lastModified:file.lastModified,importedAt:new Date().toISOString()};
      save();renderValidation();renderAll();
    });
  };
}

const baseBuildImport=window.buildImport;
window.buildImport=function(){
  baseBuildImport();
  const card=document.querySelector('#fileCards .upload:nth-child(2) .drop');
  if(card){
    const text=card.childNodes[0];
    if(text&&text.nodeType===3)text.textContent='Weekly production · Mixer sheet: A/B · Grinder sheet: HTMSOY47/HTMSOY50 · PREMIX ignored ';
  }
  installProductionHandler();
};

window.renderValidation=function(){
  baseRenderValidation();
  const m=S.production?.meta;
  const rows=[...document.querySelectorAll('#validationRows tr')];
  const tr=rows.find(r=>/^Production/i.test((r.cells?.[0]?.textContent||'').trim()));
  if(!tr)return;
  if(!m||m.format!=='weekly_summary')return;
  const p=currentKey(),current=S.production.weekly?.[p]||null;
  const total=current?((+current.s47||0)+(+current.s50||0)+(+current.ma||0)+(+current.mb||0)):null;
  if(total!=null&&tr.cells[1])tr.cells[1].innerHTML=fmt(total)+' t<div style="font-size:10px;color:#7a8792;margin-top:3px">A '+fmt(current.ma)+' · B '+fmt(current.mb)+' · HTMSOY47 '+fmt(current.s47)+' · HTMSOY50 '+fmt(current.s50)+'</div>';
  if(tr.cells[2]){
    let label='';
    if(m.status==='needs_mapping')label='Needs mapping · '+(m.reason||'Check workbook');
    else if(current)label='Weekly mapped';
    else label='Parsed '+(m.startWeek||'—')+' → '+(m.endWeek||'—')+' · No '+p;
    tr.cells[2].innerHTML='<span class="badge '+(current?'ok':'wait')+'" title="Mixer: '+(m.sheets?.mixer||'—')+' | Grinder: '+(m.sheets?.grinder||'—')+'">'+label+'</span>';
  }
};

if(typeof buildImport==='function')buildImport();
if(typeof renderValidation==='function')renderValidation();
})();