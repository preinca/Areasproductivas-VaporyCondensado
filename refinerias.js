'use strict';
(() => {
  // Módulo independiente: no altera ni mezcla registros de las cinco áreas productivas.
  const REFINERIES={'desmet':'DE SMET','tiuna':'TIUNA'};
  const FIELDS=['Condicion Fuga','Condicion by pass','Condicion Manometros','Condicion Aislamiento','Condicion Seguridad'];
  const HEADERS=['Registro ID','Fecha','Turno','Supervisor','Operador','Área','Punto / Equipo',...FIELDS,'#Aviso B1','Observacion Operador','Archivo origen','Fecha de carga'];
  const DIAGRAMS={
    'DE SMET':[
      ['Entrada Principal a planta','images/desmet/entrada_principal_planta.png','Entrada Principal a Planta'],
      ['Manifold de vapor para trazas','images/desmet/manifold_vapor_para_trazas.png','Manifold de vapor para trazas'],
      ['Sistema rociador de ácidos grasos 1 y 2','images/desmet/sistema_rociador_acidos_grasos.png','Sistema de rociador de ácidos grasos'],
      ['Manifold Vapores de contacto','images/desmet/manifold_vapores_de_contacto.png','Manifold vapores de contacto'],
      ['Manifold alimentacion de vapor generador de vacío','images/desmet/manifold_alimentacion_vapor_generador_vacio.png','Manifold alimentación de vapor generador de vacío']
    ],
    TIUNA:[
      ['Generación de vapor','images/tiuna/conjunto_tuberias_instrumentacion.png','Generación de Vapor'],
      ['Sistema de distribución de vapor Blanqueo','images/tiuna/distribucion_vapor_blanqueo_dos_niveles.png','Sistema de distribución de vapor de blanqueo'],
      ['Intercambiador de calor Blanqueo','images/tiuna/intercambiador_calor_e103.png','Intercambiador de calor del sistema de blanqueo (E-103)'],
      ['Manifold distribución de vapor Blanqueo','images/tiuna/manifold_distribucion_vapor_blanqueo.png','Manifold de distribución de vapor de blanqueo'],
      ['Alimentación de vapor Blanqueo','images/tiuna/alimentacion_vapor_sistema_blanqueo.png','Alimentación de vapor del sistema de blanqueo'],
      ['Intercambiadores de calor Blanqueo','images/tiuna/intercambiadores_calor_blanqueo.png','Intercambiadores de calor del sistema de blanqueo'],
      ['Blanqueo','images/tiuna/valvula_control_vapor_reactor_blanqueo.png','Sistema de blanqueo: válvula de control de vapor y reactor']
    ]
  };
  const $=id=>document.getElementById(id);
  const tx=v=>String(v??'').trim();
  const norm=v=>tx(v).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/\s+/g,' ');
  const esc=v=>String(v??'').replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
  const DATE_FORMAT=/^(\d{4})-(\d{2})-(\d{2})$/;
  function dateValue(value){
    if(value===undefined||value===null||value==='')return '';
    if(typeof value==='number'){
      const d=window.XLSX?.SSF?.parse_date_code(value);
      return d?`${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`:'';
    }
    const s=tx(value);let m;
    if(m=s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/))return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
    if(m=s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/)){
      const year=Number(m[3])<100?Number(m[3])+2000:Number(m[3]);
      return `${year}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
    }
    return '';
  }
  const fmtDate=v=>dateValue(v)?.split('-').reverse().join('/')||'—';
  const stateCell=v=>{
    const s=norm(v);
    if(!s)return 'pending';
    if(['N/A','NA','NO APLICA','NO APLICA.'].includes(s))return 'na';
    if(['OK','BIEN','NORMAL','SI','CONFORME'].includes(s))return 'ok';
    return 'issue';
  };
  const stateRow=r=>{const s=FIELDS.map(k=>stateCell(r[k]));return s.includes('issue')?'issue':s.includes('ok')||s.includes('na')?'ok':'pending';};
  const labelState=v=>`<span class="state ${stateCell(v)==='ok'?'':stateCell(v)}">${esc(tx(v)||'—')}</span>`;
  let rows=[],lastSync='',toastTimer=null;
  const pagination={desmet:1,tiuna:1};
  const pageSize=10;
  function status(slug,text,mode){
    const el=$(slug+'-sync-status');el.textContent=text;
    el.className='status-pill'+(mode?' '+mode:'');
  }
  function toast(message){
    const el=$('toast');el.textContent=message;el.hidden=false;
    clearTimeout(toastTimer);toastTimer=setTimeout(()=>{el.hidden=true},7000);
  }
  function choice(slug,name,values,placeholder){
    const el=$(slug+'-filter-'+name),old=el.value;
    el.replaceChildren(new Option(placeholder,'*'));
    for(const val of values)el.add(new Option(val,val));
    el.value=Array.from(el.options).some(o=>o.value===old)?old:'*';
  }
  function filtered(slug){
    const f={date:$(slug+'-filter-date').value,turn:$(slug+'-filter-turn').value,status:$(slug+'-filter-status').value,q:norm($(slug+'-filter-search').value)};
    return rows.filter(r=>{
      if(r['Área']!==REFINERIES[slug])return false;
      if(f.date!=='*'&&r.Fecha!==f.date)return false;
      if(f.turn!=='*'&&r.Turno!==f.turn)return false;
      if(f.status!=='*'&&stateRow(r)!==f.status)return false;
      return !f.q||[r['Punto / Equipo'],r['#Aviso B1'],r['Observacion Operador'],r.Operador,r.Supervisor,r['Área']].some(v=>norm(v).includes(f.q));
    }).sort((a,b)=>b.Fecha.localeCompare(a.Fecha)||b.Turno.localeCompare(a.Turno)||a['Punto / Equipo'].localeCompare(b['Punto / Equipo']));
  }
  function diagramFor(area,point){
    const entry=DIAGRAMS[area]?.find(v=>norm(v[0])===norm(point));
    return entry?{src:entry[1],title:entry[2]}:null;
  }
  // Referencias visuales por refinería: los doce diagramas ya publicados, sin duplicar archivos.
  function renderRefChoices(slug,reset=false){
    const area=REFINERIES[slug],list=DIAGRAMS[area]||[],select=$(slug+'-ref-point'),previous=select.value;
    select.replaceChildren();
    for(const [point] of list)select.add(new Option(point,point));
    select.value=!reset&&list.some(([point])=>point===previous)?previous:(list[0]?.[0]||'');
    renderRefs(slug);
  }
  function renderRefs(slug){
    const area=REFINERIES[slug],point=$(slug+'-ref-point').value,mode=$(slug+'-ref-mode').value,view=$(slug+'-reference-content');
    $(slug+'-ref-point').disabled=mode!=='point';
    if(mode==='hide'){view.hidden=true;view.replaceChildren();return;}
    view.hidden=false;
    const refs=mode==='area'?(DIAGRAMS[area]||[]):(DIAGRAMS[area]||[]).filter(([name])=>norm(name)===norm(point));
    const heading=`<div class="ref-heading"><h3>${esc(area)} · ${esc(mode==='area'?'Todos los diagramas':point)}</h3><span>${refs.length} diagrama(s)</span></div>`;
    const cards=refs.map(([,src,title])=>`
      <a class="refinery-diagram-card" href="${esc(src)}" target="_blank" rel="noopener noreferrer" data-refinery-diagram="${esc(src)}" data-refinery-title="${esc(title)}" aria-label="Ampliar diagrama: ${esc(title)}">
        <span class="refinery-diagram-preview">
          <img src="${esc(src)}" alt="Diagrama técnico: ${esc(title)}" loading="lazy" decoding="async">
          <span class="refinery-diagram-zoom" aria-hidden="true">⤢</span>
        </span>
        <span class="refinery-diagram-caption">${esc(title)}</span>
      </a>`).join('');
    view.innerHTML=heading+(refs.length?`<div class="refinery-diagram-gallery">${cards}</div>`:'<div class="empty">No hay un diagrama asociado a este punto.</div>');
  }
  function render(slug){
    const areaRows=rows.filter(r=>r['Área']===REFINERIES[slug]);
    choice(slug,'date',[...new Set(areaRows.map(r=>r.Fecha).filter(Boolean))].sort().reverse(),'Todas las fechas');
    choice(slug,'turn',[...new Set(areaRows.map(r=>r.Turno).filter(Boolean))].sort(),'Todos los turnos');
    const list=filtered(slug),latest=list.map(r=>r.Fecha).sort().at(-1)||'';
    $(slug+'-kpi-total').textContent=String(list.length);
    $(slug+'-kpi-issues').textContent=String(list.filter(r=>stateRow(r)==='issue').length);
    $(slug+'-kpi-b1').textContent=String(list.filter(r=>tx(r['#Aviso B1'])).length);
    $(slug+'-kpi-latest').textContent=fmtDate(latest);
    $(slug+'-kpi-latest-detail').textContent=latest?`${list.filter(r=>r.Fecha===latest).length} registro(s) de esa fecha`:'Sin información';
    const pages=Math.max(1,Math.ceil(list.length/pageSize));
    pagination[slug]=Math.max(1,Math.min(pagination[slug],pages));
    const begin=(pagination[slug]-1)*pageSize,end=Math.min(begin+pageSize,list.length);
    $(slug+'-history-count').textContent=`${list.length} registro${list.length===1?'':'s'}`;
    $(slug+'-page-meta').textContent=`${list.length?begin+1:0}–${end} de ${list.length}`;
    $(slug+'-prev-page').disabled=pagination[slug]<=1;
    $(slug+'-next-page').disabled=pagination[slug]>=pages;
    $(slug+'-page-select').innerHTML=Array.from({length:pages},(_,i)=>`<option value="${i+1}" ${i+1===pagination[slug]?'selected':''}>${i+1} / ${pages}</option>`).join('');
    $(slug+'-history-body').innerHTML=!list.length?'<tr><td colspan="13" class="empty">Sin registros que coincidan con los filtros.</td></tr>':list.slice(begin,end).map(r=>{
      const diagram=diagramFor(r['Área'],r['Punto / Equipo']);
      const btn=diagram?`<button class="photo-link" type="button" data-refinery-diagram="${esc(diagram.src)}" data-refinery-title="${esc(diagram.title)}">Ver diagrama</button>`:'';
      return `<tr><td>${esc(fmtDate(r.Fecha))}</td><td>${esc(r.Turno||'—')}</td><td><strong>${esc(r['Área'])}</strong></td><td>${esc(r['Punto / Equipo'])}${btn}</td>${FIELDS.map(f=>`<td>${labelState(r[f])}</td>`).join('')}<td>${esc(r['#Aviso B1']||'—')}</td><td title="${esc(r['Observacion Operador'])}">${esc(r['Observacion Operador']||'—')}</td><td>${esc(r.Supervisor||'—')}</td><td>${esc(r.Operador||'—')}</td></tr>`;
    }).join('');
    if(lastSync)$(slug+'-last-sync').textContent='Última sincronización: '+lastSync;
  }
  function apiURL(){
    let saved='';try{saved=localStorage.getItem('vc_checklist_api_url')||'';}catch(_){}
    return tx(window.CHECKLIST_API_URL||saved);
  }
  function send(payload){return new Promise((resolve,reject)=>{
    const endpoint=apiURL();
    if(!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec\/?$/.test(endpoint)){reject(new Error('La conexión con Google Apps Script no está configurada.'));return;}
    const nonce='ref_post_'+Date.now()+'_'+Math.random().toString(36).slice(2);
    const frame=document.createElement('iframe'),form=document.createElement('form');
    frame.hidden=true;frame.name=nonce;frame.setAttribute('aria-hidden','true');
    form.action=endpoint;form.method='POST';form.target=nonce;form.style.display='none';
    const inputs={payload:JSON.stringify(payload),nonce,origin:location.origin||'*'};
    for(const [name,value] of Object.entries(inputs)){
      const el=document.createElement('input');el.type='hidden';el.name=name;el.value=value;form.appendChild(el);
    }
    let done=false;
    const finish=(err,response)=>{if(done)return;done=true;clearTimeout(timer);window.removeEventListener('message',receive);form.remove();frame.remove();err?reject(err):resolve(response);};
    const receive=e=>{
      const m=e.data;if(m?.type!=='VC_SHEET_RESPONSE'||m.nonce!==nonce)return;
      let host='';try{const u=new URL(e.origin);if(u.protocol!=='https:')return;host=u.hostname;}catch(_){return;}
      if(host!=='script.google.com'&&host!=='script.googleusercontent.com'&&!host.endsWith('.googleusercontent.com'))return;
      finish(m.payload?.ok?null:new Error(m.payload?.error||'Error de Apps Script'),m.payload);
    };
    const timer=setTimeout(()=>finish(new Error('No se recibió confirmación de Apps Script. Revise el despliegue.')),45000);
    window.addEventListener('message',receive);document.body.append(frame,form);form.submit();
  });}
  function explain(error){
    if(/Acción POST no reconocida/i.test(error.message||''))return 'El panel ya está actualizado, pero es necesario publicar la versión nueva de Code.gs en Apps Script para activar las refinerías.';
    return error.message||'No se pudo conectar.';
  }
  async function refresh(){
    for(const slug of Object.keys(REFINERIES)){status(slug,'Sincronizando...');$(slug+'-refresh').disabled=true;}
    try{
      const result=await send({action:'read_refineries'});
      rows=(result.rows||[]).map(raw=>{
        const row={};for(const key of HEADERS)row[key]=tx(raw[key]);
        row.Fecha=dateValue(row.Fecha);
        row['Área']=norm(row['Área']);
        if(!row['Condicion Fuga']&&raw.Fugas)row['Condicion Fuga']=tx(raw.Fugas);
        return row;
      }).filter(r=>r.Fecha&&Object.values(REFINERIES).includes(r['Área'])&&r['Punto / Equipo']);
      lastSync=new Date().toLocaleString('es-VE');
      for(const slug of Object.keys(REFINERIES)){pagination[slug]=1;status(slug,`${rows.filter(r=>r['Área']===REFINERIES[slug]).length} registros cargados`,'success');render(slug);}
    }catch(error){
      for(const slug of Object.keys(REFINERIES))status(slug,'Sin conexión','error');
      toast(explain(error));console.error(error);
    }finally{for(const slug of Object.keys(REFINERIES))$(slug+'-refresh').disabled=false;}
  }
  function parseWorkbook(wb,fileName){
    const result=[],summaries=[];
    for(const area of Object.values(REFINERIES)){
      const sheetName=wb.SheetNames.find(name=>norm(name)===area);
      if(!sheetName)continue;
      const grid=window.XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,raw:true,defval:''});
      const supervisor=tx(grid[2]?.[1]),operador=tx(grid[2]?.[4]),fecha=dateValue(grid[3]?.[1]),turno=tx(grid[3]?.[4]);
      const header=grid.findIndex((line,i)=>i<15&&norm(line?.[0])==='PUNTO / EQUIPO');
      if(header<0)throw new Error(area+': no se encontró la cabecera «Punto / Equipo» del formato.');
      let count=0;
      for(let i=header+1;i<grid.length;i++){
        const row=grid[i]||[],point=tx(row[0]);
        if(!point)continue;
        const vals=row.slice(1,8).map(tx);
        if(!vals.some(Boolean))continue;
        for(let n=0;n<5;n++){
          if(vals[n]&&!['OK','NO OK','N/A','NA'].includes(norm(vals[n])))throw new Error(area+': estado no válido en «'+point+'». Utilice OK, NO OK o N/A.');
        }
        if(!fecha||!turno||!supervisor||!operador)throw new Error(area+': faltan supervisor, operador, día o turno.');
        count++;
        result.push({
          'Registro ID':[fecha,turno,area,String(i-header).padStart(2,'0')].join('|'),
          Fecha:fecha,Turno:turno,Supervisor:supervisor,Operador:operador,'Área':area,
          'Punto / Equipo':point,
          ...Object.fromEntries(FIELDS.map((field,n)=>[field,vals[n]||''])),
          '#Aviso B1':vals[5]||'','Observacion Operador':vals[6]||'',
          'Archivo origen':fileName,'Fecha de carga':''
        });
      }
      if(count)summaries.push(area+': '+count);
    }
    if(!result.length)throw new Error('No hay verificaciones diligenciadas en las hojas DE SMET o TIUNA.');
    return {result,summaries};
  }
  function askSecret(){return new Promise(resolve=>{
    const d=document.createElement('dialog');d.className='secret-dialog';
    const title=document.createElement('h3');title.textContent='Clave de actualización (WRITE_KEY)';
    const label=document.createElement('label');label.textContent='Clave privada';
    const input=document.createElement('input');input.type='password';input.required=true;input.autocomplete='off';
    label.appendChild(input);
    const actions=document.createElement('div');actions.className='secret-actions';
    const cancel=document.createElement('button');cancel.className='btn muted';cancel.textContent='Cancelar';cancel.type='button';
    const save=document.createElement('button');save.className='btn primary';save.textContent='Confirmar';save.type='button';
    actions.append(cancel,save);d.append(title,label,actions);document.body.appendChild(d);
    let done=false;const finish=v=>{if(done)return;done=true;resolve(v);d.close();d.remove();};
    cancel.addEventListener('click',()=>finish(''));save.addEventListener('click',()=>finish(input.value));
    input.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();finish(input.value);}});
    d.addEventListener('cancel',e=>{e.preventDefault();finish('');});d.showModal();input.focus();
  });}
  async function upload(e,slug){
    const file=e.target.files?.[0];if(!file)return;
    try{
      if(!window.XLSX)throw new Error('No pudo cargarse el lector Excel.');
      const book=window.XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:false});
      const {result,summaries}=parseWorkbook(book,file.name);
      const key=await askSecret();if(!key)return;
      if(!confirm(`¿Guardar o actualizar ${result.length} registros del archivo «${file.name}»?\n${summaries.join(' · ')}`))return;
      status(slug,'Guardando checklist...');
      const ans=await send({action:'upsert_refineries',key,rows:result});
      toast(`Guardado confirmado: ${ans.added||0} agregados, ${ans.updated||0} actualizados.`);
      await refresh();
    }catch(error){status(slug,'No se confirmó la carga','error');toast(explain(error));console.error(error);}
    finally{e.target.value='';}
  }
  function init(){
    for(const slug of Object.keys(REFINERIES)){
      $(slug+'-refresh').addEventListener('click',refresh);
      $(slug+'-excel-input').addEventListener('change',event=>upload(event,slug));
      for(const key of ['date','turn','status'])$(slug+'-filter-'+key).addEventListener('change',()=>{pagination[slug]=1;render(slug);});
      $(slug+'-filter-search').addEventListener('input',()=>{pagination[slug]=1;render(slug);});
      $(slug+'-clear').addEventListener('click',()=>{for(const key of ['date','turn','status'])$(slug+'-filter-'+key).value='*';$(slug+'-filter-search').value='';pagination[slug]=1;render(slug);});
      $(slug+'-prev-page').addEventListener('click',()=>{pagination[slug]=Math.max(1,pagination[slug]-1);render(slug);});
      $(slug+'-next-page').addEventListener('click',()=>{pagination[slug]++;render(slug);});
      $(slug+'-page-select').addEventListener('change',event=>{pagination[slug]=Number(event.target.value)||1;render(slug);});
      $(slug+'-ref-point').addEventListener('change',()=>renderRefs(slug));
      $(slug+'-ref-mode').addEventListener('change',()=>renderRefs(slug));
      renderRefChoices(slug,true);
      render(slug);
    }
    refresh();
  }
  document.addEventListener('DOMContentLoaded',init);
})();
