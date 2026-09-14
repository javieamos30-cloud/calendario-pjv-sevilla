import fs from 'node:fs/promises';
import ExcelJS from 'exceljs';

const OUTPUT = process.argv[2] || '_site/calendar-data.json';
const EXCEL_URL = process.env.EXCEL_URL?.trim();
const MONTHS = new Set(['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']);
const CATEGORY_COLS = [3,4,5,6,7,8];

function text(cell){
  if(cell == null) return '';
  if(typeof cell === 'object' && cell.richText) return cell.richText.map(x=>x.text).join('');
  if(typeof cell === 'object' && cell.text) return cell.text;
  if(typeof cell === 'object' && cell.result != null) return text(cell.result);
  return String(cell);
}
function addParam(url,key,val){
  try{const u=new URL(url);if(!u.searchParams.has(key))u.searchParams.set(key,val);return u.toString()}catch{return url}
}
async function fetchWorkbook(url){
  const candidates=[addParam(url,'download','1'),addParam(url,'action','download'),url];
  const notes=[];
  for(const candidate of [...new Set(candidates)]){
    try{
      const res=await fetch(candidate,{redirect:'follow',headers:{'User-Agent':'Mozilla/5.0 PJV-Calendar-Sync/2.0'}});
      const ab=await res.arrayBuffer();const b=Buffer.from(ab);const ct=res.headers.get('content-type')||'';
      if(res.ok && b.length>4 && b[0]===0x50 && b[1]===0x4b) return b;
      notes.push(`${res.status} ${ct} ${candidate}`);
    }catch(e){notes.push(`${e.message} ${candidate}`)}
  }
  throw new Error(`No he podido descargar el .xlsx desde EXCEL_URL. Comprueba que el vínculo sea público de solo lectura y que "Bloquear descarga" esté desactivado. Intentos: ${notes.join(' | ')}`);
}
function parseYear(ws){
  const title=text(ws.getCell(1,3).value).toUpperCase();
  const m=title.match(/(?:-|\s)(\d{2,4})\b/);
  if(m){const n=Number(m[1]);return n<100?2000+n:n}
  const name=ws.name.toUpperCase();
  return ['SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'].includes(name)?2026:2027;
}
function monthNumber(name){
  return ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'].indexOf(name.toUpperCase())+1;
}
function normalizeTime(raw){
  const range=raw.match(/\b([01]?\d|2[0-3])[.,:]([0-5]\d)\s*(?:h)?\s*[-–]\s*([01]?\d|2[0-3])[.,:]([0-5]\d)\b/i);
  if(range)return `${String(range[1]).padStart(2,'0')}:${range[2]}-${String(range[3]).padStart(2,'0')}:${range[4]}`;
  const m=raw.match(/\b([01]?\d|2[0-3])[.,:]([0-5]\d)\s*h?\b/i);
  return m?`${String(m[1]).padStart(2,'0')}:${m[2]}`:'';
}
function cleanTitle(raw){
  let s=raw.replace(/^\s*(\[WEB\+?\]|#WEB|🌐)\s*/i,'').trim();
  s=s.replace(/^([01]?\d|2[0-3])[.,:]([0-5]\d)\s*h?\s*[:.\-–]?\s*/i,'').trim();
  return s.replace(/\s*\r?\n+\s*/g,' · ').trim();
}
function extractPublic(raw){
  const whole=raw.match(/^\s*\[WEB\+\]\s*([\s\S]+)$/i);if(whole)return [whole[1].trim()];
  return raw.split(/\r?\n/).map(s=>s.trim()).filter(s=>/^\s*(\[WEB\]|#WEB|🌐)/i.test(s)).map(s=>s.replace(/^\s*(\[WEB\]|#WEB|🌐)\s*/i,'').trim()).filter(Boolean);
}
function isYes(v){
  return ['SI','SÍ','YES','TRUE','1'].includes(text(v).trim().toUpperCase());
}
function dayFromExcelValue(v){
  if(v instanceof Date) return v.getUTCDate();
  if(typeof v === 'number'){
    const d=new Date(Math.round((v-25569)*86400*1000));
    return d.getUTCDate();
  }
  const m=text(v).match(/(?:^|\D)(\d{1,2})(?:\D|$)/);
  return m?Number(m[1]):NaN;
}
function sourceCellFor(ws, day, category){
  let col=null;
  for(const c of CATEGORY_COLS){
    if(text(ws.getCell(2,c).value).trim().toUpperCase()===category.trim().toUpperCase()){col=c;break;}
  }
  if(!col) return null;
  for(let r=3;r<=ws.rowCount;r++){
    if(Number(ws.getCell(r,1).value)===day) return ws.getCell(r,col);
  }
  return null;
}
function buildFromControl(wb){
  const control=wb.getWorksheet('CONTROL WEB');
  if(!control) return null;
  const events=[];
  for(let r=5;r<=control.rowCount;r++){
    if(!isYes(control.getCell(r,1).value)) continue;
    const monthName=text(control.getCell(r,3).value).trim().toUpperCase();
    const category=text(control.getCell(r,4).value).trim().toUpperCase();
    if(!MONTHS.has(monthName) || !category) continue;
    const ws=wb.getWorksheet(monthName); if(!ws) continue;
    const day=dayFromExcelValue(control.getCell(r,2).value); if(!Number.isInteger(day)||day<1||day>31) continue;
    const sourceCell=sourceCellFor(ws,day,category);
    const rawSource=sourceCell?text(sourceCell.value).trim():'';
    const override=text(control.getCell(r,6).value).trim();
    const raw=override||rawSource;
    if(!raw) continue;
    const title=cleanTitle(raw); if(!title) continue;
    const year=parseYear(ws); const month=monthNumber(monthName);
    events.push({
      date:`${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`,
      category,
      title,
      time:normalizeTime(raw)
    });
  }
  return events;
}
function buildFromLegacyMarkers(wb){
  const events=[];
  for(const ws of wb.worksheets){
    const sheetName=ws.name.toUpperCase().trim();if(!MONTHS.has(sheetName))continue;
    const month=monthNumber(sheetName);const year=parseYear(ws);const categories={};
    for(const c of CATEGORY_COLS)categories[c]=text(ws.getCell(2,c).value).trim().toUpperCase();
    for(let r=3;r<=ws.rowCount;r++){
      const day=Number(ws.getCell(r,1).value);if(!Number.isInteger(day)||day<1||day>31)continue;
      for(const c of CATEGORY_COLS){
        const raw=text(ws.getCell(r,c).value).trim();if(!raw)continue;
        for(const item of extractPublic(raw)){
          const title=cleanTitle(item);if(!title)continue;
          events.push({date:`${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`,category:categories[c]||'GENERAL',title,time:normalizeTime(item)});
        }
      }
    }
  }
  return events;
}
async function build(){
  if(!EXCEL_URL) throw new Error('Falta EXCEL_URL en los Secrets de GitHub.');
  const buffer=await fetchWorkbook(EXCEL_URL);
  const wb=new ExcelJS.Workbook();await wb.xlsx.load(buffer);
  let events=buildFromControl(wb);
  const mode=events===null?'legacy':'control-web';
  if(events===null) events=buildFromLegacyMarkers(wb);
  events.sort((a,b)=>`${a.date} ${a.time||'99:99'}`.localeCompare(`${b.date} ${b.time||'99:99'}`));
  const out={source:'onedrive',mode,generatedAt:new Date().toISOString(),eventCount:events.length,events};
  await fs.mkdir(OUTPUT.split('/').slice(0,-1).join('/')||'.',{recursive:true});
  await fs.writeFile(OUTPUT,JSON.stringify(out,null,2));
  console.log(`Calendario generado (${mode}): ${events.length} eventos públicos.`);
}
build().catch(e=>{console.error(e);process.exit(1)});
