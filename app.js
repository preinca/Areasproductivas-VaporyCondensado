'use strict';
(() => {
  const AREAS=['QUESO','TOMATE','MAYONESA','MANIFOLD','MARGARINA'];
  const FIELDS=['Condicion Fuga','Condicion by pass','Condicion Manometros','Condicion Aislamiento','Condicion Seguridad'];
  const HEADERS=['Registro ID','Fecha','Turno','Supervisor','Operador','Área','Punto / Equipo',...FIELDS,'#Aviso B1','Observacion Operador','Archivo origen','Fecha de carga'];
  const ORDER={
    QUESO:['Linea Alimentacion','Linea Cocina 1','Linea Vapor Alimentacion','Linea de condensado #1 entrada vapor','Linea de condensado #2 Filtro','Linea de condensado #3 Cocina','Linea Cocina 2','Intercambiador CIP','Fase Oleosa','Tina de fundicion','Tapadora'],
    TOMATE:['Intercambiador de Calor','Intercambiador Tubular Salsa Base'],
    MAYONESA:['CIP','Manifold'],
    MANIFOLD:['Vapor Principal','Vapor Pick Heater','Manifold Vapor Investigacion y Desarrollo','Manifold Vapor Margarina'],
    MARGARINA:['Area de Pasteurizacion','Area de Fundicion de Emulsificante']
  };
  const PHOTOS=[
    ['QUESO','Linea Alimentacion','queso_linea_alimentacion.jpg','Queso · Línea Alimentación'],
    ['QUESO','Linea Cocina 1','queso_cocina_1.jpg','Queso · Cocina 1'],
    ['QUESO','Linea Vapor Alimentacion','queso_cocina_1.jpg','Queso · Cocina 1 · Vapor y condensado'],
    ['QUESO','Linea Vapor Alimentacion','queso_cocina_2.jpg','Queso · Cocina 2 · Vapor y condensado'],
    ['QUESO','Linea de condensado #1 entrada vapor','queso_cocina_1.jpg','Queso · Cocina 1 · Condensado #1'],
    ['QUESO','Linea de condensado #1 entrada vapor','queso_cocina_2.jpg','Queso · Cocina 2 · Condensado #1'],
    ['QUESO','Linea de condensado #2 Filtro','queso_cocina_1.jpg','Queso · Cocina 1 · Condensado #2'],
    ['QUESO','Linea de condensado #2 Filtro','queso_cocina_2.jpg','Queso · Cocina 2 · Condensado #2'],
    ['QUESO','Linea de condensado #3 Cocina','queso_cocina_1.jpg','Queso · Cocina 1 · Condensado #3'],
    ['QUESO','Linea de condensado #3 Cocina','queso_cocina_2.jpg','Queso · Cocina 2 · Condensado #3'],
    ['QUESO','Linea Cocina 2','queso_cocina_2.jpg','Queso · Cocina 2'],
    ['QUESO','Intercambiador CIP','queso_intercambiador_cip.jpg','Queso · Intercambiador CIP'],
    ['QUESO','Fase Oleosa','queso_fase_oleosa.jpg','Queso · Fase Oleosa'],
    ['QUESO','Tina de fundicion','queso_tina_fundicion.jpg','Queso · Tina de fundición'],
    ['QUESO','Tapadora','queso_tapadora.jpg','Queso · Tapadora'],
    ['TOMATE','Intercambiador de Calor','tomate_intercambiador_calor.jpg','Tomate · Intercambiador de Calor'],
    ['TOMATE','Intercambiador Tubular Salsa Base','tomate_intercambiador_tubular_salsa_base.jpg','Tomate · Intercambiador Tubular'],
    ['MAYONESA','CIP','mayonesa_cip.jpg','Mayonesa · CIP'],
    ['MAYONESA','Manifold','mayonesa_manifold.jpg','Mayonesa · Manifold'],
    ['MANIFOLD','Vapor Principal','manifold_vapor_principal.jpg','Manifold · Vapor Principal'],
    ['MANIFOLD','Manifold Vapor Margarina','manifold_vapor_margarina.jpg','Manifold · Vapor Margarina'],
    ['MANIFOLD','Vapor Pick Heater','manifold_pick_heater.jpg','Manifold · Pick Heater'],
    ['MANIFOLD','Manifold Vapor Investigacion y Desarrollo','manifold_investigacion_desarrollo.jpg','Manifold · Investigación y Desarrollo'],
    ['MARGARINA','Area de Pasteurizacion','margarina_pasteurizacion_1.jpg','Margarina · Pasteurización · 1'],
    ['MARGARINA','Area de Pasteurizacion','margarina_pasteurizacion_2.jpg','Margarina · Pasteurización · 2'],
    ['MARGARINA','Area de Pasteurizacion','margarina_pasteurizacion_3.jpg','Margarina · Pasteurización · 3'],
    ['MARGARINA','Area de Pasteurizacion','margarina_pasteurizacion_4.jpg','Margarina · Pasteurización · 4'],
    ['MARGARINA','Area de Fundicion de Emulsificante','margarina_fundicion_emulsificante_1.jpg','Margarina · Fundición · 1'],
    ['MARGARINA','Area de Fundicion de Emulsificante','margarina_fundicion_emulsificante_2.jpg','Margarina · Fundición · 2'],
    ['MARGARINA','Area de Fundicion de Emulsificante','margarina_fundicion_emulsificante_3.jpg','Margarina · Fundición · 3'],
    ['MARGARINA','Area de Fundicion de Emulsificante','margarina_fundicion_emulsificante_4.jpg','Margarina · Fundición · 4']
  ].map(([area,point,file,label])=>({area,point,file,label}));
  const $=id=>document.getElementById(id);
  const esc=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const tx=x=>String(x??'').trim();
  const norm=x=>tx(x).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/\s+/g,' ');
  const dateValue=value=>{
    if(value===undefined||value===null||value==='')return '';
    if(typeof value==='number'){
      const parsed=window.XLSX?.SSF?.parse_date_code(value);
      return parsed?`${parsed.y}-${String(parsed.m).padStart(2,'0')}-${String(parsed.d).padStart(2,'0')}`:'';
    }
    const s=tx(value); let m;
    if(m=s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/))return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
    if(m=s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/)){const year=Number(m[3])<100?Number(m[3])+2000:Number(m[3]);return `${year}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;}
    return '';
  };
  const fmtDate=d=>dateValue(d)?.split('-').reverse().join('/')||'—';
  const stateCell=value=>{
    const s=norm(value);if(!s)return 'pending';
    if(['NO APLICA','N/A','NA','NO APLICA.'].includes(s))return 'na';
    if(['OK','BIEN','NORMAL','SI','SÍ','CONFORME'].includes(s))return 'ok';
    return 'issue';
  };
  const stateRow=r=>{const states=FIELDS.map(k=>stateCell(r[k]));return states.includes('issue')?'issue':states.includes('ok')||states.includes('na')?'ok':'pending';};
  const labelState=value=>`<span class="state ${stateCell(value)==='ok'?'':stateCell(value)}">${esc(tx(value)||'—')}</span>`;
  let data=[],page=1,pageSize=10,selectedArea='*',toastTimer=null,lastSync='';
  function toast(text){let e=$('toast');e.textContent=text;e.hidden=false;clearTimeout(toastTimer);toastTimer=setTimeout(()=>e.hidden=true,6500);}
  function status(msg,mode){let e=$('sync-status');e.textContent=msg;e.className='status-pill'+(mode?' '+mode:'');}
  function normalizeRow(r){let o={};for(let h of HEADERS)o[h]=tx(r?.[h]);if(!o['Condicion Fuga']&&r?.Fugas)o['Condicion Fuga']=tx(r.Fugas);o.Fecha=dateValue(o.Fecha);o['Área']=norm(o['Área']);return o;}
  function choice(id,arr,placeholder){let el=$(id),old=el.value;el.replaceChildren();let first=new Option(placeholder,'*');el.add(first);for(let v of arr)el.add(new Option(v,v));el.value=[...el.options].some(x=>x.value===old)?old:'*';}
  function selected(){return {area:$('filter-area').value,date:$('filter-date').value,turn:$('filter-turn').value,state:$('filter-status').value,q:norm($('filter-search').value)};}
  function filtered(){const f=selected();return data.filter(r=>{
    if(f.area!=='*'&&r['Área']!==f.area)return false;
    if(f.date!=='*'&&r.Fecha!==f.date)return false;
    if(f.turn!=='*'&&r.Turno!==f.turn)return false;
    if(f.state!=='*'&&stateRow(r)!==f.state)return false;
    return !f.q||[r['Punto / Equipo'],r['#Aviso B1'],r['Observacion Operador'],r.Operador,r.Supervisor,r['Área']].some(x=>norm(x).includes(f.q));
  }).sort((a,b)=>b.Fecha.localeCompare(a.Fecha)||b.Turno.localeCompare(a.Turno)||a['Área'].localeCompare(b['Área'])||a['Punto / Equipo'].localeCompare(b['Punto / Equipo']));}
  function render(){
    choice('filter-area',[...new Set(data.map(r=>r['Área']).filter(Boolean))].sort(),'Todas las áreas');
    choice('filter-date',[...new Set(data.map(r=>r.Fecha).filter(Boolean))].sort().reverse(),'Todas las fechas');
    choice('filter-turn',[...new Set(data.map(r=>r.Turno).filter(Boolean))].sort(),'Todos los turnos');
    const rows=filtered();$('kpi-total').textContent=rows.length;$('kpi-issues').textContent=rows.filter(r=>stateRow(r)==='issue').length;
    $('kpi-b1').textContent=rows.filter(r=>tx(r['#Aviso B1'])).length;
    const latest=rows.map(r=>r.Fecha).sort().at(-1)||'';$('kpi-latest').textContent=fmtDate(latest);
    const countLatest=rows.filter(r=>r.Fecha===latest).length;$('kpi-latest-detail').textContent=latest?`${countLatest} registro(s) de esa fecha`:'Sin información';
    const areaSelected=selected().area;
    $('areas').innerHTML=AREAS.map(area=>{
      const a=data.filter(r=>r['Área']===area),last=a.map(r=>r.Fecha).sort().at(-1)||'';
      const day=a.filter(r=>r.Fecha===last),turn=day.map(r=>r.Turno).sort().at(-1)||'',latestRows=turn?day.filter(r=>r.Turno===turn):day;
      const issues=latestRows.filter(r=>stateRow(r)==='issue').length,pending=latestRows.filter(r=>stateRow(r)==='pending').length,ok=latestRows.filter(r=>stateRow(r)==='ok').length;
      const code=issues?'issue':latestRows.length?'ok':'',n=issues||pending||ok;
      const caption=issues?'Desviaciones':pending?'Pendientes':latestRows.length?'Verificados':'Sin registros';
      return `<button class="area-card ${areaSelected===area?'active':''}" data-area="${area}"><strong>${area}<i class="${code}"></i></strong><div class="count">${n}</div><div class="caption">${caption}</div><div class="extra">${last?`Último: ${esc(fmtDate(last))}${turn?' · '+esc(turn):''}<br>${ok} OK · ${pending} pendientes`:'Aún no hay verificaciones'}</div></button>`;
    }).join('');
    const pages=Math.max(1,Math.ceil(rows.length/pageSize));page=Math.max(1,Math.min(page,pages));
    let start=(page-1)*pageSize,end=Math.min(start+pageSize,rows.length);
    $('history-count').textContent=`${rows.length} registro${rows.length===1?'':'s'}`;
    $('page-meta').textContent=`${rows.length?start+1:0}–${end} de ${rows.length}`;
    $('prev-page').disabled=page===1;$('next-page').disabled=page===pages;
    $('page-select').innerHTML=Array.from({length:pages},(_,i)=>`<option value="${i+1}" ${page===i+1?'selected':''}>${i+1} / ${pages}</option>`).join('');
    $('history-body').innerHTML=!rows.length?'<tr><td colspan="13" class="empty">Sin registros que coincidan con los filtros.</td></tr>':rows.slice(start,end).map(r=>{
      const photo=PHOTOS.some(p=>p.area===r['Área']&&norm(p.point)===norm(r['Punto / Equipo']));
      return `<tr><td>${esc(fmtDate(r.Fecha))}</td><td>${esc(r.Turno||'—')}</td><td><strong>${esc(r['Área'])}</strong></td><td>${esc(r['Punto / Equipo'])}${photo?`<button class="photo-link" data-photo-area="${esc(r['Área'])}" data-photo-point="${esc(r['Punto / Equipo'])}">Ver foto</button>`:''}</td>${FIELDS.map(f=>`<td>${labelState(r[f])}</td>`).join('')}<td>${esc(r['#Aviso B1']||'—')}</td><td title="${esc(r['Observacion Operador'])}">${esc(r['Observacion Operador']||'—')}</td><td>${esc(r.Supervisor||'—')}</td><td>${esc(r.Operador||'—')}</td></tr>`;
    }).join('');
    if(lastSync)$('last-sync').textContent=`Última sincronización: ${lastSync}`;
  }
  function areaRefs(area,point,mode){if(mode==='area'){let seen=new Set();return PHOTOS.filter(p=>p.area===area&&!seen.has(p.file)&&seen.add(p.file));}return PHOTOS.filter(p=>p.area===area&&norm(p.point)===norm(point));}
  function renderRefChoices(reset=false){let area=$('ref-area').value||AREAS[0],el=$('ref-point'),old=el.value;
    el.replaceChildren();for(let p of ORDER[area]||[])el.add(new Option(p,p));el.value=!reset&&[...el.options].some(o=>o.value===old)?old:ORDER[area]?.[0]||'';renderRefs();}
  function imageSrc(file){return `referencias_condensado/${encodeURIComponent(file)}`;}
  function remoteSrc(file){return `https://raw.githubusercontent.com/preinca/dashboard/main/referencias_condensado/${encodeURIComponent(file)}`;}
  function renderRefs(){const area=$('ref-area').value,point=$('ref-point').value,mode=$('ref-mode').value,view=$('reference-content');$('ref-point').disabled=mode!=='point';
    if(mode==='hide'){view.hidden=true;view.replaceChildren();return;}
    view.hidden=false;const refs=areaRefs(area,point,mode);const head=`<div class="ref-heading"><h3>${esc(area)} · ${esc(mode==='area'?'Todas las referencias':point)}</h3><span>${refs.length} fotografía(s)</span></div>`;
    view.innerHTML=head+(refs.length?`<div class="photo-grid ${refs.length===1?'single':''}">${refs.map(p=>`<button class="photo-card" data-image-file="${esc(p.file)}" data-image-label="${esc(p.label)}"><img src="${imageSrc(p.file)}" data-remote="${remoteSrc(p.file)}" alt="${esc(p.label)}" loading="lazy"><strong>${esc(p.label)}</strong></button>`).join('')}</div>`:'<div class="empty">No hay fotografía asociada a este punto.</div>');
    view.querySelectorAll('img[data-remote]').forEach(img=>img.addEventListener('error',()=>{const remote=img.dataset.remote;if(remote&&img.src!==remote){img.removeAttribute('data-remote');img.src=remote;}}));
  }
  function openPhoto(file,label){const modal=$('modal'),img=$('modal-img');img.src=imageSrc(file);img.onerror=()=>{img.onerror=null;img.src=remoteSrc(file)};img.alt=label;$('modal-caption').textContent=label;modal.hidden=false;}
  function closePhoto(){$('modal').hidden=true;$('modal-img').removeAttribute('src');}
  // El navegador conserva SOLO la URL pública del Web App. Nunca guarda contraseñas.
  let readKey='';
  const API_STORAGE_KEY='vc_checklist_api_url';
  function readStoredURL(){try{return localStorage.getItem(API_STORAGE_KEY)||'';}catch(_){return '';}}
  function saveStoredURL(url){try{localStorage.setItem(API_STORAGE_KEY,url);}catch(_){/* La URL permanece en memoria */}window.CHECKLIST_API_URL=url;}
  function apiURL(){return tx(window.CHECKLIST_API_URL||readStoredURL());}
  function validateEndpoint(endpoint){
    try{const u=new URL(endpoint);return u.protocol==='https:'&&u.hostname==='script.google.com'&&/^\/macros\/s\/[^/]+\/exec\/?$/.test(u.pathname);}
    catch(_){return false;}
  }
  function ensureURL(){
    let endpoint=apiURL();
    if(!endpoint){const supplied=prompt('Pegue la URL /exec de la nueva implementación de Google Apps Script:');
      if(supplied===null)return '';
      endpoint=tx(supplied);
      if(!validateEndpoint(endpoint))throw new Error('Indique una URL oficial de Apps Script que termine en /exec.');
      saveStoredURL(endpoint);
    }
    return endpoint;
  }
  function askSecret(title){return new Promise(resolve=>{
    const d=document.createElement('dialog');d.className='secret-dialog';
    const heading=document.createElement('h3');heading.textContent=title;
    const label=document.createElement('label');label.textContent='Clave privada';
    const input=document.createElement('input');input.type='password';input.required=true;input.autocomplete='off';input.placeholder='Introducir clave';
    label.appendChild(input);
    const actions=document.createElement('div');actions.className='secret-actions';
    const cancel=document.createElement('button');cancel.className='btn muted';cancel.type='button';cancel.textContent='Cancelar';
    const submit=document.createElement('button');submit.className='btn primary';submit.type='button';submit.textContent='Confirmar';
    actions.append(cancel,submit);d.append(heading,label,actions);document.body.appendChild(d);
    let done=false;
    const finish=value=>{if(done)return;done=true;resolve(value);d.close();d.remove();};
    cancel.addEventListener('click',()=>finish(''));
    submit.addEventListener('click',()=>finish(input.value));
    input.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();finish(input.value)}});
    d.addEventListener('cancel',e=>{e.preventDefault();finish('')});
    d.showModal();input.focus();
  });}
  async function connectionKey(){
    if(readKey)return readKey;
    const supplied=await askSecret('Clave de consulta (READ_KEY)');
    if(!tx(supplied))return '';
    readKey=supplied;
    return readKey;
  }
  // Lectura y escritura por POST de formulario: la clave no se incluye en la URL.
  function sendWithIframe(payload){return new Promise((resolve,reject)=>{
    const endpoint=apiURL();if(!endpoint){reject(new Error('No se configuró Apps Script. Pulse Conectar.'));return;}
    const nonce='vc_post_'+Date.now()+'_'+Math.random().toString(36).slice(2),frame=document.createElement('iframe'),form=document.createElement('form');let done=false;
    frame.name=nonce;frame.hidden=true;frame.setAttribute('aria-hidden','true');
    form.action=endpoint;form.method='POST';form.target=nonce;form.style.display='none';
    const values={payload:JSON.stringify(payload),nonce,origin:location.origin||'*'};
    for(const [name,value] of Object.entries(values)){let input=document.createElement('input');input.type='hidden';input.name=name;input.value=value;form.appendChild(input);}
    const finish=(err,res)=>{if(done)return;done=true;clearTimeout(timer);window.removeEventListener('message',receive);form.remove();frame.remove();if(err)reject(err);else resolve(res)};
    const receive=e=>{const message=e.data;if(!message||message.type!=='VC_SHEET_RESPONSE'||message.nonce!==nonce)return;
      // El HTML de Apps Script puede ejecutarse dentro de un iframe anidado;
      // validar el origen de Google y el nonce, no exigir e.source===frame.contentWindow.
      let host='';try{const u=new URL(e.origin);if(u.protocol!=='https:')return;host=u.hostname;}catch(_){return;}
      if(host!=='script.google.com'&&host!=='script.googleusercontent.com'&&!host.endsWith('.googleusercontent.com'))return;
      finish(message.payload?.ok?null:new Error(message.payload?.error||'Error del servidor'),message.payload);
    };
    const timer=setTimeout(()=>finish(new Error('No se recibió confirmación de Apps Script. Compruebe el despliegue, el acceso y la conexión.')),45000);
    window.addEventListener('message',receive);document.body.append(frame,form);form.submit();
  });}
  async function refresh(){
    try{if(!ensureURL())return;if(!(await connectionKey()))return;}
    catch(e){toast(e.message);return;}
    status('Sincronizando...');$('refresh').disabled=true;$('connect').disabled=true;
    try{
      const ans=await sendWithIframe({action:'read_verified',key:readKey});
      data=(ans.rows||[]).map(normalizeRow).filter(r=>r.Fecha&&r['Área']&&r['Punto / Equipo']);
      lastSync=new Date().toLocaleString('es-VE');status(`${data.length} registros cargados`,'success');page=1;render();
    }catch(e){
      if(/Clave de consulta incorrecta/.test(e.message))readKey='';
      status('Error de conexión','error');toast(e.message);console.error(e);
    }finally{$('refresh').disabled=false;$('connect').disabled=false;}
  }
  function connect(){refresh();}
  function configure(){
    const supplied=prompt('Pegue la URL /exec del Apps Script exclusivo para Vapor y Condensado:',apiURL());
    if(supplied===null)return;
    const endpoint=tx(supplied);
    if(!validateEndpoint(endpoint)){toast('La dirección debe ser https://script.google.com/macros/s/.../exec');return;}
    saveStoredURL(endpoint);readKey='';data=[];page=1;lastSync='';render();
    status('URL guardada · Conectar para consultar');
    toast('Dirección actualizada. Se guarda en este navegador, sin contraseñas.');
  }
  function parseWorkbook(wb,fileName){let result=[];
    for(const area of AREAS){const s=wb.SheetNames.find(n=>norm(n)===area);if(!s)continue;
      const grid=window.XLSX.utils.sheet_to_json(wb.Sheets[s],{header:1,raw:true,defval:''}),supervisor=tx(grid[2]?.[1]),operador=tx(grid[2]?.[4]),fecha=dateValue(grid[3]?.[1]),turno=tx(grid[3]?.[4]);
      let header=-1;for(let i=0;i<Math.min(grid.length,15);i++)if(norm(grid[i]?.[0])==='PUNTO / EQUIPO'){header=i;break;}
      if(header<0)continue;let filled=0;
      for(let i=header+1;i<grid.length;i++){const raw=grid[i]||[],point=tx(raw[0]);if(!point)continue;let v=raw.slice(1,8).map(tx);if(!v.some(Boolean))continue;
        filled++;const seq=String(i-header).padStart(2,'0'),id=[fecha||'SINFECHA',turno||'SINTURNO',area,seq].join('|');
        result.push({'Registro ID':id,Fecha:fecha,Turno:turno,Supervisor:supervisor,Operador:operador,'Área':area,'Punto / Equipo':point,...Object.fromEntries(FIELDS.map((h,j)=>[h,v[j]||''])),'#Aviso B1':v[5]||'','Observacion Operador':v[6]||'','Archivo origen':fileName,'Fecha de carga':''});
      }
      if(filled&&(!fecha||!turno||!supervisor||!operador))throw new Error(`${area}: falta completar supervisor, operador, día o turno en el formato Excel.`);
    }
    if(!result.length)throw new Error('No se encontraron registros diligenciados en las cinco hojas del checklist.');return result;
  }
  async function upload(e){const file=e.target.files?.[0];if(!file)return;
    try{if(!window.XLSX)throw new Error('No pudo cargarse la biblioteca de lectura de Excel. Compruebe la conexión a Internet.');
      const wb=window.XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:false});const rows=parseWorkbook(wb,file.name);
      if(!ensureURL())return;
      const key=await askSecret('Clave de actualización (WRITE_KEY)');if(!key)return;
      if(!confirm(`¿Guardar o actualizar ${rows.length} registros de «${file.name}» en la nueva Google Sheet?`))return;
      status('Guardando checklist...');const ans=await sendWithIframe({action:'upsert_verified',key,rows});
      if(ans?.ok){toast(`Guardado confirmado: ${ans.added||0} agregados, ${ans.updated||0} actualizados.`);await refresh();}
    }catch(err){status('No se confirmó la carga','error');toast(err.message);console.error(err);}finally{e.target.value='';}
  }
  function init(){AREAS.forEach(a=>$('ref-area').add(new Option(a,a)));renderRefChoices(true);render();
    $('refresh').addEventListener('click',refresh);$('excel-input').addEventListener('change',upload);
    for(const id of ['filter-area','filter-date','filter-turn','filter-status'])$(id).addEventListener('change',()=>{page=1;render();});
    $('filter-search').addEventListener('input',()=>{page=1;render();});
    $('clear').addEventListener('click',()=>{['filter-area','filter-date','filter-turn','filter-status'].forEach(id=>$(id).value='*');$('filter-search').value='';page=1;render();});
    $('areas').addEventListener('click',e=>{const t=e.target.closest('[data-area]');if(!t)return;const area=t.dataset.area;$('filter-area').value=$('filter-area').value===area?'*':area;page=1;render();$('historico').scrollIntoView({behavior:'smooth'});});
    $('prev-page').addEventListener('click',()=>{if(page>1){page--;render();}});$('next-page').addEventListener('click',()=>{page++;render();});
    $('page-select').addEventListener('change',e=>{page=Number(e.target.value)||1;render();});
    $('ref-area').addEventListener('change',()=>renderRefChoices(true));$('ref-point').addEventListener('change',renderRefs);$('ref-mode').addEventListener('change',renderRefs);
    $('history-body').addEventListener('click',e=>{const b=e.target.closest('[data-photo-area]');if(!b)return;$('ref-area').value=b.dataset.photoArea;renderRefChoices(true);$('ref-point').value=b.dataset.photoPoint;$('ref-mode').value='point';renderRefs();$('referencias').scrollIntoView({behavior:'smooth'});});
    $('reference-content').addEventListener('click',e=>{const b=e.target.closest('[data-image-file]');if(b)openPhoto(b.dataset.imageFile,b.dataset.imageLabel);});
    $('modal-close').addEventListener('click',closePhoto);$('modal').addEventListener('click',e=>{if(e.target.id==='modal')closePhoto();});document.addEventListener('keydown',e=>{if(e.key==='Escape')closePhoto();});
    document.querySelectorAll('.nav-item').forEach(link=>link.addEventListener('click',()=>{document.querySelectorAll('.nav-item').forEach(x=>x.classList.remove('active'));link.classList.add('active');}));
    status(apiURL()?'Clave de consulta requerida':'Pendiente de conectar');
  }
  document.addEventListener('DOMContentLoaded',init);
})();
