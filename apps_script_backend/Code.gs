/**
 * BACKEND INDEPENDIENTE — CHECKLIST VAPOR Y CONDENSADO
 * Google Sheet exclusiva: Base_Maestra_Checklist_Vapor_Condensado
 * REQUIERE Script Properties: WRITE_KEY (cargas)
 * DESPLIEGUE: Web app, ejecutar como propietario; conceder acceso solo conforme
 * a la política de acceso de la organización.
 *
 * IMPORTANTE: las consultas de registros están abiertas mediante POST.
 * Cuando el Web App es público, cualquier persona que conozca el endpoint
 * puede consultar el historial. No publicar información confidencial.
 */
const VC_VERSION = 'VC-1.3-FECHAS-2026-11-03';
const VC_SHEET_ID = '1wk8lgVDfJdKgTKo0ZGznBBhVw_pmPmu30r41G5aA9FY';
const VC_TAB = 'LISTA_VERIFICADOS';
const VC_HEADERS = [
  'Registro ID','Fecha','Turno','Supervisor','Operador','Área','Punto / Equipo',
  'Condicion Fuga','Condicion by pass','Condicion Manometros','Condicion Aislamiento',
  'Condicion Seguridad','#Aviso B1','Observacion Operador','Archivo origen','Fecha de carga'
];
function vcSheet_(){
  const ss=SpreadsheetApp.openById(VC_SHEET_ID);
  const sh=ss.getSheetByName(VC_TAB);
  if(!sh)throw Error('No existe la pestaña '+VC_TAB+' en la base independiente.');
  const first=sh.getRange(1,1,1,VC_HEADERS.length).getDisplayValues()[0];
  if(first.some((x,i)=>String(x).trim()!==VC_HEADERS[i]))throw Error('Los encabezados de LISTA_VERIFICADOS no coinciden con el esquema previsto.');
  return sh;
}
function vcDate_(value,sheetTimeZone){
  // Las fechas procedentes de Sheets se interpretan con la zona horaria de la hoja,
  // no con la del proyecto Apps Script. Evita que 02/11 aparezca como 01/11.
  if(value instanceof Date&&!isNaN(value.getTime()))return Utilities.formatDate(value,sheetTimeZone||Session.getScriptTimeZone()||'America/Caracas','yyyy-MM-dd');
  const s=String(value||'').trim();
  const m=s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  return m?m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2):s;
}
/** Lee el día que se visualiza en la hoja, sin cambiarlo por zona horaria.
 * La columna B de LISTA_VERIFICADOS usa el formato dd/mm/yyyy.
 */
function vcDateDisplayed_(value){
  const s=String(value||'').trim();
  let m=s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:\s|$)/);
  if(m)return m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2);
  m=s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})(?:\s|$)/);
  if(m)return m[3]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[1]).slice(-2);
  return '';
}
function vcRead_(){
  const sh=vcSheet_(),last=sh.getLastRow();
  if(last<2)return [];
  const sheetTimeZone=sh.getParent().getSpreadsheetTimeZone();
  const values=sh.getRange(2,1,last-1,VC_HEADERS.length).getValues();
  // Obtener exactamente el texto visible de la columna Fecha (B).
  // De este modo, la fecha 02/11/2026 nunca se interpreta como 01/11.
  const displayDates=sh.getRange(2,2,last-1,1).getDisplayValues();
  return values.map((row,rowIndex)=>{
    if(!row.some(v=>v!==''&&v!==null))return null;
    const item={};
    VC_HEADERS.forEach((h,i)=>{
      let value=row[i];
      if(h==='Fecha')value=vcDateDisplayed_(displayDates[rowIndex]?.[0])||vcDate_(value,sheetTimeZone);
      else if(value instanceof Date)value=Utilities.formatDate(value,Session.getScriptTimeZone()||'America/Caracas','yyyy-MM-dd HH:mm:ss');
      item[h]=value===undefined||value===null?'':String(value);
    });
    return item;
  }).filter(Boolean);
}
function vcData_(extra){return Object.assign({ok:true,version:VC_VERSION,rows:vcRead_(),updatedAt:new Date().toISOString()},extra||{});}
function vcJson_(result){return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);}
function vcFrame_(result,nonce,origin){
  const target=/^https?:\/\/[a-z0-9.-]+(?::\d+)?$/i.test(String(origin||''))?origin:'*';
  const envelope=JSON.stringify({type:'VC_SHEET_RESPONSE',nonce:String(nonce||''),payload:result})
    .replace(/</g,'\\u003c').replace(/\u2028/g,'\\u2028').replace(/\u2029/g,'\\u2029');
  const js='try{parent.postMessage('+envelope+','+JSON.stringify(target)+');}catch(e){}try{top.postMessage('+envelope+','+JSON.stringify(target)+');}catch(e){}';
  const html='<!doctype html><html><head><meta charset="utf-8"></head><body><script>'+js+'</'+'script></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function doGet(e){
  // GET solo informa del estado; los registros se consultan mediante POST.
  return vcJson_({ok:true,service:'Checklist Vapor y Condensado',version:VC_VERSION,readProtected:false});
}
function vcAuthorized_(key,propertyName){
  const expected=PropertiesService.getScriptProperties().getProperty(propertyName);
  if(!expected||String(expected).length<8)throw Error('La propiedad '+propertyName+' no está configurada (mínimo 8 caracteres).');
  if(!key||String(key)!==expected)throw Error(propertyName==='READ_KEY'?'Clave de consulta incorrecta.':'Clave de actualización incorrecta.');
}
function vcSafeValue_(v){
  if(v===undefined||v===null)return '';
  const s=String(v);
  // Evitar que un valor procedente de Excel se interprete como fórmula de Google Sheets.
  return /^[=+@]/.test(s.trimStart())?"'"+s:s;
}
function vcUpsert_(incoming){
  if(!Array.isArray(incoming)||!incoming.length)throw Error('El checklist no contiene verificaciones.');
  if(incoming.length>5000)throw Error('Demasiados registros en una sola carga. Divida el archivo.');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const sh=vcSheet_(),last=sh.getLastRow(),existing=last>1?sh.getRange(2,1,last-1,VC_HEADERS.length).getValues():[];
    const byId=new Map();existing.forEach((r,i)=>{if(String(r[0]||'').trim())byId.set(String(r[0]).trim(),i);});
    const appended=[],updates=[];let added=0,updated=0;const now=new Date();
    incoming.forEach(item=>{
      if(!item||typeof item!=='object')return;
      const o=Object.assign({},item);o.Fecha=vcDate_(o.Fecha);o['Área']=String(o['Área']||'').trim().toUpperCase();
      o.Turno=String(o.Turno||'').trim();
      if(!o['Condicion Fuga']&&o.Fugas)o['Condicion Fuga']=o.Fugas;
      if(!o.Fecha||!o['Área']||!String(o['Punto / Equipo']||'').trim())return;
      const id=String(o['Registro ID']||[o.Fecha,o.Turno,o['Área'],o['Punto / Equipo']].join('|')).trim();
      o['Registro ID']=id;o['Fecha de carga']=now;
      const row=VC_HEADERS.map(h=>h==='Fecha de carga'?o[h]:vcSafeValue_(o[h]));
      if(byId.has(id)){const pos=byId.get(id);if(pos<existing.length){updates.push({pos,row});existing[pos]=row;updated++;}else{appended[pos-existing.length]=row;updated++;}}
      else{byId.set(id,existing.length+appended.length);appended.push(row);added++;}
    });
    if(updates.length){updates.forEach(x=>sh.getRange(x.pos+2,1,1,VC_HEADERS.length).setValues([x.row]));}
    if(appended.length)sh.getRange(sh.getLastRow()+1,1,appended.length,VC_HEADERS.length).setValues(appended);
    SpreadsheetApp.flush();return {added,updated,total:sh.getLastRow()-1};
  }finally{lock.releaseLock();}
}
function doPost(e){
  const p=(e&&e.parameter)||{};let result;
  try{
    const body=JSON.parse(p.payload||(e&&e.postData&&e.postData.contents)||'{}');
    if(body.action==='read_verified'){
      result=vcData_();
    }else if(body.action==='upsert_verified'){
      vcAuthorized_(body.key,'WRITE_KEY');
      result=Object.assign({ok:true,version:VC_VERSION},vcUpsert_(body.rows));
    }else throw Error('Acción POST no reconocida.');
  }catch(err){result={ok:false,error:String(err&&err.message||err)};}
  return p.nonce?vcFrame_(result,p.nonce,p.origin):vcJson_(result);
}

/** Ejecutar desde el editor: valida la hoja y la clave de escritura sin escribir registros. */
function validarConfiguracion(){
  const write=PropertiesService.getScriptProperties().getProperty('WRITE_KEY')||'';
  if(write.length<8)throw Error('Configure WRITE_KEY con al menos 8 caracteres.');
  const sheet=vcSheet_();
  const result={ok:true,version:VC_VERSION,tab:sheet.getName(),registros:Math.max(0,sheet.getLastRow()-1),claveEscrituraConfigurada:true,lecturaSinClave:true};
  Logger.log(JSON.stringify(result));
  return result;
}
