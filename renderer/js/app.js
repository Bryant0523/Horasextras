// ══════════════════════════════════════════
//  INDEXEDDB
// ══════════════════════════════════════════
const DB_NAME='HorasExtrasHK3', DB_VER=1;
let db=null;
function openDB(){
  return new Promise((ok,er)=>{
    const r=indexedDB.open(DB_NAME,DB_VER);
    r.onupgradeneeded=e=>{
      const d=e.target.result;
      ['sedes','empleados','resultados'].forEach(s=>{if(!d.objectStoreNames.contains(s))d.createObjectStore(s,{keyPath:'id'});});
    };
    r.onsuccess=e=>{db=e.target.result;ok(db);};
    r.onerror=e=>er(e.target.error);
  });
}
const txS=(s,m='readonly')=>db.transaction(s,m).objectStore(s);
const dbAll=s=>new Promise((ok,er)=>{const r=txS(s).getAll();r.onsuccess=()=>ok(r.result);r.onerror=()=>er(r.error);});
const dbPut=(s,o)=>new Promise((ok,er)=>{const r=txS(s,'readwrite').put(o);r.onsuccess=()=>ok();r.onerror=()=>er(r.error);});
const dbDel=(s,k)=>new Promise((ok,er)=>{const r=txS(s,'readwrite').delete(k);r.onsuccess=()=>ok();r.onerror=()=>er(r.error);});
const dbClear=s=>new Promise((ok,er)=>{const r=txS(s,'readwrite').clear();r.onsuccess=()=>ok();r.onerror=()=>er(r.error);});

// ══════════════════════════════════════════
//  ESTADO
// ══════════════════════════════════════════
const S={sedes:[],empleados:[],rawRows:[],csvHdr:[],resultados:[]};

// ══════════════════════════════════════════
//  BOOT
// ══════════════════════════════════════════
// Sede activa global
let sedeActiva = localStorage.getItem('sedeActiva') || '';

openDB().then(async()=>{
  document.getElementById('db-info').textContent='✓ BD activa';
  S.sedes     = await dbAll('sedes');
  S.empleados = await dbAll('empleados');
  S.resultados= await dbAll('resultados');
  renderSedes(); renderEmps(); fillSelects(); fillSedeActiva();
  // Apply saved sedeActiva
  if(sedeActiva && S.sedes.find(s=>s.id===sedeActiva)){
    document.getElementById('sede-activa').value = sedeActiva;
  } else {
    sedeActiva = '';
  }
  if(S.resultados.length){ renderStats(); renderRes(); spill('ok', S.resultados.length+' reg.'); }
  else if(S.sedes.length) spill('ok','Config. cargada');
  const t=localStorage.getItem('theme')||'dark';
  setTheme(t);
}).catch(e=>{document.getElementById('db-info').textContent='⚠ BD Error';console.error(e);});

// ══════════════════════════════════════════
//  TEMA
// ══════════════════════════════════════════
function setTheme(t){
  document.documentElement.setAttribute('data-theme',t);
  document.getElementById('theme-btn').textContent=t==='dark'?'☀️ Claro':'🌙 Oscuro';
  localStorage.setItem('theme',t);
}
function toggleTheme(){
  const cur=document.documentElement.getAttribute('data-theme');
  setTheme(cur==='dark'?'light':'dark');
}

// ══════════════════════════════════════════
//  NAV
// ══════════════════════════════════════════
const NTITLES={sedes:'Sedes y <span>Horarios</span>',empleados:'Gestión de <span>Empleados</span>',dashboard:'<span>Dashboard</span>',importar:'Importar <span>CSV</span>',resultados:'Ver <span>Resultados</span>',ausencias:'<span>Ausencias</span>',llegadas:'Reporte de <span>Llegadas</span>',michel:'Reporte de <span>Llegadas (Michel)</span>',exportar:'Exportar <span>Excel / PDF</span>'};

function onSedeActivaChange(){
  sedeActiva = document.getElementById('sede-activa').value;
  localStorage.setItem('sedeActiva', sedeActiva);
  // Sincronizar todos los filtros de sede en cada sección
  ['ef-sede','rf-sede','xf-sede','df-sede','lf-sede'].forEach(id=>{
    const el=document.getElementById(id);
    if(el) el.value = sedeActiva;
  });
  renderEmps(); renderRes(); renderConsol();
  // Actualizar dashboard si está visible
  const dashActive=document.getElementById('sec-dashboard')?.classList.contains('active');
  if(dashActive){ fillDashFilters(); renderDash(); }
  // Actualizar llegadas si está visible
  const llegActive=document.getElementById('sec-llegadas')?.classList.contains('active');
  if(llegActive){ fillLlegadasFilters(); renderLlegadas(); }
}

function fillSedeActiva(){
  const el = document.getElementById('sede-activa');
  if(!el) return;
  const v = el.value;
  el.innerHTML = '<option value="">Todas las sedes</option>';
  S.sedes.forEach(s=>{
    const o=document.createElement('option');
    o.value=s.id; o.textContent=s.nombre;
    el.appendChild(o);
  });
  el.value = v;
}

function go(name,el){
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.remove('active'));
  document.getElementById('sec-'+name).classList.add('active');
  el.classList.add('active');
  document.getElementById('pgTitle').innerHTML=NTITLES[name];
  if(['empleados','exportar'].includes(name)) fillSelects();
  if(name==='exportar') renderConsol();
  if(name==='dashboard'){ fillDashFilters(); renderDash(); }
  if(name==='ausencias'){ fillAusenciasFilters(); renderAusencias(); }
  if(name==='llegadas'){ fillLlegadasFilters(); renderLlegadas(); }
  if(name==='michel'){ fillMichelFilters(); renderMichel(); }
  // Sync sede filter on navigate
  ['ef-sede','rf-sede','xf-sede'].forEach(id=>{
    const fEl=document.getElementById(id);
    if(fEl && sedeActiva) fEl.value=sedeActiva;
  });
  closeSidebar();
}

// ══════════════════════════════════════════
//  TOAST / SPILL
// ══════════════════════════════════════════
let _tt;
function toast(msg,t='info'){const el=document.getElementById('toast');el.textContent=msg;el.className='show '+t;clearTimeout(_tt);_tt=setTimeout(()=>el.classList.remove('show'),3500);}
function spill(t,msg){const el=document.getElementById('spill');el.textContent=msg;el.className='status-pill '+(t==='ok'?'pill-ok':t==='warn'?'pill-warn':'pill-idle');}

// ── Sidebar móvil ──
function toggleSidebar(){
  document.getElementById('sidebar').classList.toggle('open');
  document.getElementById('sidebar-overlay').classList.toggle('open');
}
function closeSidebar(){
  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('sidebar-overlay').classList.remove('open');
}

// ── Limpiar todos los resultados ──
function confirmarLimpiarTodo(){ openModal('modal-clear'); closeSidebar(); }
async function limpiarTodo(){
  await dbClear('resultados');
  S.resultados=[];
  S.rawRows=[];S.csvHdr=[];
  _notas={};
  document.getElementById('prev-card').style.display='none';
  document.getElementById('csvf').value='';
  closeModal('modal-clear');
  renderStats();renderRes();
  spill('idle','Sin datos');
  _resPage=1;_llegPage=1;_ausPage=1;_resFiltered=[];_llegFiltered=[];_ausFiltered=[];
  document.getElementById('res-pagination').innerHTML='';
  document.getElementById('llegadas-pagination').innerHTML='';
  document.getElementById('aus-pagination').innerHTML='';
  document.getElementById('llegadas-stats-area').innerHTML='';
  document.getElementById('aus-stats-area').innerHTML='';
  document.getElementById('stats-area').innerHTML='';
  document.getElementById('consol-area').innerHTML='<div class="empty-state"><div class="eico">📋</div><p>Procesa datos primero.</p></div>';
  toast('Resultados eliminados','success');
}

// ══════════════════════════════════════════
//  TIEMPO UTILS
// ══════════════════════════════════════════
const DIAS=['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];
const DIAS_FULL=['Domingo','Lunes','Martes','Miércoles','Jueves','Viernes','Sábado'];
const t2m=t=>{if(!t)return 0;const[h,m]=t.split(':').map(Number);return h*60+m;};

// FORMATO PRINCIPAL: 00h 00m
function m2hm(min){
  if(min==null||min<=0) return '00h 00m';
  const h=Math.floor(min/60), m=Math.round(min%60);
  return `${String(h).padStart(2,'0')}h ${String(m).padStart(2,'0')}m`;
}
// para "no marcó salida" etc
function m2hmOrDash(min){return(!min||min<=0)?'—':m2hm(min);}

const fmtT=d=>d?String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0'):'--:--';
const dkey=d=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
const fmtFecha=s=>{const[y,m,d]=s.split('-');return`${d}/${m}/${y}`;};

function parseDT(s){
  if(!s)return null; s=s.toString().trim();
  const m1=s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\s+(\d{1,2}):(\d{2})/);
  if(m1)return new Date(+m1[1],+m1[2]-1,+m1[3],+m1[4],+m1[5]);
  const m2=s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})\s+(\d{1,2}):(\d{2})/);
  if(m2)return new Date(+m2[3],+m2[2]-1,+m2[1],+m2[4],+m2[5]);
  const d=new Date(s);return isNaN(d)?null:d;
}

// ══════════════════════════════════════════
//  HORARIO FORM HELPERS
// ══════════════════════════════════════════
function readHForm(prefix){
  const dias={};
  [0,1,2,3,4,5,6].forEach(dow=>{
    const r=document.getElementById(prefix+dow);if(!r)return;
    dias[dow]={activo:r.querySelector('.'+prefix.replace('dr-','dc').replace('edr-','edc')||'.dc').checked,
               entrada:r.querySelector('.'+prefix.replace('dr-','de').replace('edr-','ede')||'.de').value,
               salida:r.querySelector('.'+prefix.replace('dr-','ds').replace('edr-','eds')||'.ds').value,
               almuerzo:parseInt(r.querySelector('.'+prefix.replace('dr-','da').replace('edr-','eda')||'.da').value)||0};
  });
  return dias;
}

// simpler version
function readDias(idPrefix, cbClass, entClass, salClass, almClass){
  const dias={};
  [0,1,2,3,4,5,6].forEach(dow=>{
    const r=document.getElementById(idPrefix+dow);if(!r)return;
    dias[dow]={activo:r.querySelector(cbClass).checked,entrada:r.querySelector(entClass).value,salida:r.querySelector(salClass).value,almuerzo:parseInt(r.querySelector(almClass).value)||0};
  });
  return dias;
}
function writeDias(idPrefix, cbClass, entClass, salClass, almClass, dias){
  const DEF={activo:false,entrada:'08:00',salida:'17:00',almuerzo:60};
  [0,1,2,3,4,5,6].forEach(dow=>{
    const r=document.getElementById(idPrefix+dow);if(!r)return;
    const d=dias&&dias[dow]?dias[dow]:DEF;
    r.querySelector(cbClass).checked=!!d.activo;
    r.querySelector(entClass).value=d.entrada||'08:00';
    r.querySelector(salClass).value=d.salida||'17:00';
    r.querySelector(almClass).value=d.almuerzo??60;
  });
}
const DEF_DIAS={
  1:{activo:true,entrada:'08:00',salida:'17:00',almuerzo:60},
  2:{activo:true,entrada:'08:00',salida:'17:00',almuerzo:60},
  3:{activo:true,entrada:'08:00',salida:'17:00',almuerzo:60},
  4:{activo:true,entrada:'08:00',salida:'17:00',almuerzo:60},
  5:{activo:true,entrada:'08:00',salida:'16:00',almuerzo:60},
  6:{activo:false,entrada:'09:00',salida:'14:00',almuerzo:0},
  0:{activo:false,entrada:'08:00',salida:'17:00',almuerzo:60},
};

// ══════════════════════════════════════════
//  SEDES
// ══════════════════════════════════════════
async function guardarSede(){
  const nom=document.getElementById('s-nom').value.trim();
  const tol=parseInt(document.getElementById('s-tol').value)||0;
  const antici=parseInt(document.getElementById('s-antici').value)||0;
  const minExt=parseInt(document.getElementById('s-minext').value)??15;
  const noct=document.getElementById('s-noct').value||'21:00';
  const vhora=parseFloat(document.getElementById('s-vhora').value)||0;
  const eid=document.getElementById('s-eid').value;
  if(!nom){toast('Escribe el nombre','error');return;}
  const dias=readDias('dr-','.dc','.de','.ds','.da');
  if(!Object.values(dias).some(d=>d.activo)){toast('Activa al menos un día','error');return;}
  const s={id:eid||Date.now().toString(),nombre:nom,tolerancia:tol,antici,minExt,hnoct:noct,vhora,dias};
  await dbPut('sedes',s);
  const i=S.sedes.findIndex(x=>x.id===s.id);
  if(i>=0)S.sedes[i]=s;else S.sedes.push(s);
  renderSedes();fillSelects();fillSedeActiva();resetSedeForm();
  toast(`Sede "${nom}" guardada`,'success');
}
function editSede(id){
  const s=S.sedes.find(x=>x.id===id);if(!s)return;
  document.getElementById('s-nom').value=s.nombre;
  document.getElementById('s-tol').value=s.tolerancia||10;
  document.getElementById('s-antici').value=s.antici||0;
  document.getElementById('s-minext').value=s.minExt??15;
  document.getElementById('s-noct').value=s.hnoct||'21:00';
  document.getElementById('s-vhora').value=s.vhora||0;
  document.getElementById('s-eid').value=s.id;
  document.getElementById('btn-cancel-edit').style.display='inline-flex';
  writeDias('dr-','.dc','.de','.ds','.da',s.dias);
  document.getElementById('s-nom').scrollIntoView({behavior:'smooth'});
  toast('Modo edición','info');
}
async function delSede(id){
  await dbDel('sedes',id);S.sedes=S.sedes.filter(s=>s.id!==id);
  renderSedes();fillSelects();toast('Sede eliminada','info');
}
function resetSedeForm(){
  document.getElementById('s-nom').value='';document.getElementById('s-eid').value='';
  document.getElementById('btn-cancel-edit').style.display='none';
  writeDias('dr-','.dc','.de','.ds','.da',DEF_DIAS);
}
function renderSedes(){
  document.getElementById('sc').textContent=S.sedes.length;
  const el=document.getElementById('sedes-list');
  if(!S.sedes.length){el.innerHTML='<div class="empty-state"><div class="eico">🏢</div><p>Agrega tu primera sede.</p></div>';return;}
  el.innerHTML=S.sedes.map(s=>{
    const d=s.dias||{};
    const pills=[1,2,3,4,5,6,0].map(dow=>{
      const dc=d[dow];const lbl=DIAS[dow];
      if(!dc)return`<span class="day-pill">${lbl}</span>`;
      const tip=dc.activo?(dc.entrada+'–'+dc.salida+(dc.almuerzo?' alm.'+dc.almuerzo+'min':' sin alm.')):'No laboral';
      const sub=dc.activo?(' <small style="font-size:9px;">'+dc.entrada.slice(0,5)+'</small>'):'';
      return`<span class="day-pill ${dc.activo?'on':''}" title="${tip}">${lbl}${sub}</span>`;
    }).join('');
    return`<div class="sede-card">
      <div class="sede-card-header">
        <div><div class="sede-card-name">${s.nombre}</div>
        <div style="font-size:11px;color:var(--muted);font-family:'DM Mono',monospace;margin-top:3px;">Tolerancia: ${s.tolerancia} min · Mín. extras: ${s.minExt??15} min · Noct. desde: ${s.hnoct}</div>
        <div style="margin-top:8px;">${pills}</div></div>
        <div style="display:flex;gap:8px;flex-shrink:0;margin-left:12px;">
          <button class="icon-btn" onclick="editSede('${s.id}')">✏️ Editar</button>
          <button class="icon-btn del" onclick="delSede('${s.id}')">✕</button>
        </div>
      </div></div>`;
  }).join('');
}
async function loadDemoSedes(){
  const mk=(vi,sa)=>({
    1:{activo:true,entrada:'08:00',salida:'17:00',almuerzo:60},
    2:{activo:true,entrada:'08:00',salida:'17:00',almuerzo:60},
    3:{activo:true,entrada:'08:00',salida:'17:00',almuerzo:60},
    4:{activo:true,entrada:'08:00',salida:'17:00',almuerzo:60},
    5:{activo:true,entrada:'08:00',salida:vi,almuerzo:60},
    6:{activo:sa,entrada:'09:00',salida:'14:00',almuerzo:0},
    0:{activo:false,entrada:'08:00',salida:'17:00',almuerzo:60},
  });
  const demo=[
    {id:'ds1',nombre:'Sede Centro',tolerancia:10,minExt:15,hnoct:'21:00',dias:mk('16:00',true)},
    {id:'ds2',nombre:'Sede Norte', tolerancia:15,minExt:15,hnoct:'21:00',dias:mk('16:00',false)},
    {id:'ds3',nombre:'Sede Sur',   tolerancia:10,minExt:15,hnoct:'21:00',dias:mk('16:00',true)},
  ];
  for(const s of demo)await dbPut('sedes',s);
  S.sedes=await dbAll('sedes');
  renderSedes();fillSelects();toast('3 sedes de ejemplo cargadas','success');
}

// ══════════════════════════════════════════
//  EMPLEADOS
// ══════════════════════════════════════════
let empSchedOpen=false;
let _activeSchedTab='A';
function switchSchedTab(t){
  _activeSchedTab=t;
  document.getElementById('sched-panel-A').style.display=t==='A'?'block':'none';
  document.getElementById('sched-panel-B').style.display=t==='B'?'block':'none';
  document.getElementById('tab-ha').style.background=t==='A'?'var(--accent)':'transparent';
  document.getElementById('tab-ha').style.color=t==='A'?'#000':'var(--muted2)';
  document.getElementById('tab-hb').style.background=t==='B'?'var(--accent3)':'transparent';
  document.getElementById('tab-hb').style.color=t==='B'?'#fff':'var(--muted2)';
}
function toggleEmpSched(){
  empSchedOpen=!empSchedOpen;
  document.getElementById('emp-sched-box').classList.toggle('open',empSchedOpen);
  document.getElementById('emp-sched-arrow').textContent=empSchedOpen?'▼':'▶';
}

async function guardarEmp(){
  const id=document.getElementById('e-id').value.trim();
  const nom=document.getElementById('e-nom').value.trim();
  const sid=document.getElementById('e-sede').value;
  const cargo=document.getElementById('e-cargo').value.trim();
  const descansos=document.getElementById('e-descansos').value.split(',').map(d=>d.trim()).filter(Boolean);
  const eid=document.getElementById('e-eid').value;
  if(!id||!nom||!sid){toast('ID, nombre y sede son obligatorios','error');return;}
  // Horario A
  let diasPersonal=null, diasPersonalB=null, hbDesde=null, hbHasta=null, haDesde=null;
  if(empSchedOpen){
    diasPersonal=readDias('edr-','.edc','.ede','.eds','.eda');
    if(!Object.values(diasPersonal).some(d=>d.activo)) diasPersonal=null;
    haDesde=document.getElementById('e-ha-desde').value||null;
    // Horario B
    const bDesde=document.getElementById('e-hb-desde').value||null;
    if(bDesde){
      diasPersonalB=readDias('edrb-','.edcb','.edeb','.edsb','.edab');
      if(!Object.values(diasPersonalB).some(d=>d.activo)) diasPersonalB=null;
      hbDesde=bDesde;
      hbHasta=document.getElementById('e-hb-hasta').value||null;
    }
  }
  const e={id,nombre:nom,sedeId:sid,cargo,descansos,diasPersonal,haDesde,diasPersonalB,hbDesde,hbHasta};
  await dbPut('empleados',e);
  const i=S.empleados.findIndex(x=>x.id===id);
  if(i>=0)S.empleados[i]=e;else S.empleados.push(e);
  renderEmps();resetEmpForm();
  toast(`"${nom}" guardado`,'success');
}

async function delEmp(id){await dbDel('empleados',id);S.empleados=S.empleados.filter(e=>e.id!==id);renderEmps();}

function editEmp(id){
  const e=S.empleados.find(x=>x.id===id);if(!e)return;
  document.getElementById('e-id').value=e.id;
  document.getElementById('e-nom').value=e.nombre;
  document.getElementById('e-sede').value=e.sedeId;
  document.getElementById('e-cargo').value=e.cargo||'';
  document.getElementById('e-descansos').value=(e.descansos||[]).join(', ');
  document.getElementById('e-eid').value=e.id;
  document.getElementById('btn-cancel-emp').style.display='inline-flex';
  if(e.diasPersonal){
    if(!empSchedOpen) toggleEmpSched();
    writeDias('edr-','.edc','.ede','.eds','.eda',e.diasPersonal);
    document.getElementById('e-ha-desde').value=e.haDesde||'';
    if(e.diasPersonalB&&e.hbDesde){
      writeDias('edrb-','.edcb','.edeb','.edsb','.edab',e.diasPersonalB);
      document.getElementById('e-hb-desde').value=e.hbDesde||'';
      document.getElementById('e-hb-hasta').value=e.hbHasta||'';
    }
  }
  document.getElementById('e-id').scrollIntoView({behavior:'smooth'});
  toast('Editando empleado','info');
}

function resetEmpForm(){
  document.getElementById('e-id').value='';document.getElementById('e-nom').value='';
  document.getElementById('e-cargo').value='';document.getElementById('e-descansos').value='';document.getElementById('e-eid').value='';
  document.getElementById('btn-cancel-emp').style.display='none';
  document.getElementById('e-ha-desde').value='';
  document.getElementById('e-hb-desde').value='';
  document.getElementById('e-hb-hasta').value='';
  if(empSchedOpen) toggleEmpSched();
  switchSchedTab('A');
}

function renderEmps(){
  const q=(document.getElementById('ef-q')?.value||'').toLowerCase();
  // Use sedeActiva if no explicit filter chosen
  const sfRaw=document.getElementById('ef-sede')?.value||'';
  const sf=sfRaw||sedeActiva||'';
  const f=S.empleados.filter(e=>(!sf||e.sedeId===sf)&&(e.nombre.toLowerCase().includes(q)||e.id.includes(q)));
  document.getElementById('ec').textContent=S.empleados.length;
  const tb=document.getElementById('etb');
  if(!f.length){tb.innerHTML='<tr><td colspan="6"><div class="empty-state"><div class="eico">👤</div><p>Sin empleados.</p></div></td></tr>';return;}
  tb.innerHTML=f.map(e=>{
    const s=S.sedes.find(x=>x.id===e.sedeId);
    return`<tr>
      <td>${e.id}</td><td class="name">${e.nombre}</td>
      <td>${s?s.nombre:'<span class="badge badge-red">Sin sede</span>'}</td>
      <td>${e.cargo||'<span style="color:var(--muted)">—</span>'}</td>
      <td>${e.diasPersonal?(e.diasPersonalB?'<span class="badge badge-blue">Horario A+B</span>':'<span class="badge badge-blue">Personalizado</span>'):'<span class="badge badge-gray">De sede</span>'}</td>
      <td style="display:flex;gap:6px;">
        <button class="icon-btn" onclick="editEmp('${e.id}')">✏️</button>
        <button class="icon-btn del" onclick="delEmp('${e.id}')">✕</button>
      </td>
    </tr>`;
  }).join('');
}

// Pending import list (filled when modal opens)
let _importPending = [];

function importarEmpCSV(){
  if(!S.rawRows.length){toast('Carga un CSV primero','error');return;}
  if(!S.sedes.length){toast('Crea al menos una sede primero','error');return;}
  const ic=detectCol(S.csvHdr,['card no','card_no','cardno']);
  const nc=detectCol(S.csvHdr,['name','nombre','employee']);
  if(!ic&&!nc){toast('Sin col. de ID/nombre en CSV','error');return;}
  const seen=new Set(S.empleados.map(e=>e.id));
  _importPending=[];
  const seenNew=new Set();
  for(const row of S.rawRows){
    const rawId=(ic?row[ic]:row[nc]||'').toString().trim();
    const rawNom=(nc?row[nc]:'').toString().trim()||rawId;
    if(!rawId||seen.has(rawId)||seenNew.has(rawId))continue;
    seenNew.add(rawId);
    _importPending.push({id:rawId,nombre:rawNom});
  }
  if(!_importPending.length){toast('No hay empleados nuevos en el CSV','info');return;}
  // Fill modal
  document.getElementById('modal-import-count').textContent=_importPending.length;
  const mSede=document.getElementById('modal-import-sede');
  mSede.innerHTML='<option value="">— selecciona —</option>';
  S.sedes.forEach(s=>{const o=document.createElement('option');o.value=s.id;o.textContent=s.nombre;mSede.appendChild(o);});
  // Pre-select active sede
  if(sedeActiva) mSede.value=sedeActiva;
  // Preview list
  document.getElementById('modal-import-preview').innerHTML=
    _importPending.slice(0,15).map(e=>`<div style="padding:3px 0;border-bottom:1px solid var(--border);">${e.id} · ${e.nombre}</div>`).join('')+
    (_importPending.length>15?`<div style="color:var(--muted);padding:4px 0;">… y ${_importPending.length-15} más</div>`:'');
  openModal('modal-import');
}

async function confirmarImport(){
  const sid=document.getElementById('modal-import-sede').value;
  if(!sid){toast('Selecciona una sede','error');return;}
  let n=0;
  for(const e of _importPending){
    const emp={id:e.id,nombre:e.nombre,sedeId:sid,cargo:'',diasPersonal:null};
    await dbPut('empleados',emp);
    if(!S.empleados.find(x=>x.id===e.id)) S.empleados.push(emp);
    n++;
  }
  closeModal('modal-import');
  renderEmps();
  toast(`${n} empleados importados a "${S.sedes.find(s=>s.id===sid)?.nombre}"`, 'success');
}

function openModal(id){document.getElementById(id).classList.add('open');}
function closeModal(id){document.getElementById(id).classList.remove('open');}

// ══════════════════════════════════════════
//  SELECTS
// ══════════════════════════════════════════
function fillSelects(){
  ['e-sede','ef-sede','rf-sede','xf-sede'].forEach(id=>{
    const el=document.getElementById(id);if(!el)return;
    const v=el.value;
    el.innerHTML=`<option value="">${id==='e-sede'?'— selecciona —':'Todas las sedes'}</option>`;
    S.sedes.forEach(s=>{const o=document.createElement('option');o.value=s.id;o.textContent=s.nombre;el.appendChild(o);});
    // Apply active sede if set
    if(sedeActiva && id!=='e-sede') el.value=sedeActiva;
    else if(v) el.value=v;
  });
  fillSedeActiva();
}

// ══════════════════════════════════════════
//  CSV
// ══════════════════════════════════════════
function detectCol(hdr,cands){
  for(const c of cands){const f=hdr.find(h=>h.toLowerCase().replace(/[\s_.]/g,'').includes(c.replace(/[\s_.]/g,'')));if(f)return f;}
  return null;
}
const dzEl=document.getElementById('dz');
['dragover','dragenter'].forEach(e=>dzEl.addEventListener(e,ev=>{ev.preventDefault();dzEl.classList.add('drag');}));
['dragleave','drop'].forEach(e=>dzEl.addEventListener(e,()=>dzEl.classList.remove('drag')));
dzEl.addEventListener('drop',ev=>{ev.preventDefault();const f=ev.dataTransfer.files[0];if(f)readFile(f);});
function handleFile(e){readFile(e.target.files[0]);}
function readFile(f){const r=new FileReader();r.onload=ev=>parseCSV(ev.target.result);r.readAsText(f,'utf-8');}

function parseCSV(txt){
  const lines=txt.split(/\r?\n/).filter(l=>l.trim());
  if(!lines.length){toast('Archivo vacío','error');return;}
  let sep=',';const fl=lines[0];
  if(fl.split(';').length>fl.split(',').length)sep=';';
  else if(fl.split('\t').length>fl.split(',').length)sep='\t';
  const split=line=>{const r=[];let c='';let q=false;for(const ch of line){if(ch==='"'){q=!q;}else if(ch===sep&&!q){r.push(c.trim());c='';}else c+=ch;}r.push(c.trim());return r;};
  let hi=0;
  for(let i=0;i<Math.min(5,lines.length);i++){if(split(lines[i]).some(c=>/time|fecha|hora|name|nombre|card/i.test(c))){hi=i;break;}}
  const hdr=split(lines[hi]).map(h=>h.replace(/["\uFEFF]/g,'').trim());
  const rows=[];
  for(let i=hi+1;i<lines.length;i++){const v=split(lines[i]);if(v.length<2)continue;const row={};hdr.forEach((h,ix)=>row[h]=(v[ix]||'').replace(/"/g,'').trim());rows.push(row);}
  S.csvHdr=hdr;S.rawRows=rows;
  const tc=detectCol(hdr,['time','fecha','datetime','date','hora']);
  const ic=detectCol(hdr,['card no','card_no','cardno']);
  const nc=detectCol(hdr,['name','nombre','employee']);
  document.getElementById('prev-card').style.display='block';
  document.getElementById('csv-stats').innerHTML=`<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:12px;">
    <span class="badge badge-blue">${rows.length} registros</span>
    ${tc?`<span class="badge badge-green">✓ Tiempo: ${tc}</span>`:'<span class="badge badge-red">✗ Sin col. tiempo</span>'}
    ${ic?`<span class="badge badge-green">✓ ID: ${ic}</span>`:''}
    ${nc?`<span class="badge badge-green">✓ Nombre: ${nc}</span>`:''}
  </div>`;
  document.getElementById('csv-prev').innerHTML=`<table><thead><tr>${hdr.map(h=>`<th>${h}</th>`).join('')}</tr></thead><tbody>${rows.slice(0,6).map(r=>`<tr>${hdr.map(h=>`<td>${r[h]||''}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  toast(`CSV cargado: ${rows.length} registros`,'success');
  spill('warn',`${rows.length} filas listas`);
}
function clearCSV(){S.rawRows=[];S.csvHdr=[];document.getElementById('prev-card').style.display='none';document.getElementById('csvf').value='';}

// ══════════════════════════════════════════
//  DEMO
// ══════════════════════════════════════════
function loadDemo(){
  const names=['Ana García','Luis Martínez','Carlos López','María Rodríguez','Pedro Jiménez','Sofía Torres'];
  const ids=['00001','00002','00003','00004','00005','00006'];
  const hdr=['No.','Card No.','Name','Department','Time','Card Reader','Verification Mode'];
  const pad=n=>String(n).padStart(2,'0');
  const rows=[];let seq=1;const now=new Date();
  for(let day=25;day>=1;day--){
    const date=new Date(now.getFullYear(),now.getMonth(),day);
    const dow=date.getDay();if(dow===0)continue;
    const isSat=dow===6,isFri=dow===5;
    const baseSal=isSat?14:isFri?16:17;
    names.forEach((name,i)=>{
      const late=Math.random()<0.2,ot=Math.random()<0.3;
      const inDt=new Date(date);inDt.setHours(isSat?9:8,(late?14:0)+Math.floor(Math.random()*6),0);
      const outDt=new Date(date);outDt.setHours(baseSal,(ot?50:0)+Math.floor(Math.random()*8),0);
      const fmt=d=>`${d.getFullYear()}/${pad(d.getMonth()+1)}/${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:00`;
      rows.push({'No.':seq++,'Card No.':ids[i],'Name':name,'Department':i<3?'Ventas':'Admin','Time':fmt(inDt),'Card Reader':'Entrada','Verification Mode':'Face'});
      rows.push({'No.':seq++,'Card No.':ids[i],'Name':name,'Department':i<3?'Ventas':'Admin','Time':fmt(outDt),'Card Reader':'Salida','Verification Mode':'Face'});
    });
  }
  S.csvHdr=hdr;S.rawRows=rows;
  document.getElementById('prev-card').style.display='block';
  document.getElementById('csv-stats').innerHTML=`<div style="display:flex;gap:10px;"><span class="badge badge-blue">${rows.length} reg. demo</span><span class="badge badge-green">✓ Time · Card No. · Name</span></div>`;
  document.getElementById('csv-prev').innerHTML=`<table><thead><tr>${hdr.map(h=>`<th>${h}</th>`).join('')}</tr></thead><tbody>${rows.slice(0,6).map(r=>`<tr>${hdr.map(h=>`<td>${r[h]||''}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  if(!S.sedes.length)loadDemoSedes().then(()=>{if(!S.empleados.length)addDemoEmps();});
  else if(!S.empleados.length)addDemoEmps();
  toast('Datos demo cargados','success');
}
async function addDemoEmps(){
  const names=['Ana García','Luis Martínez','Carlos López','María Rodríguez','Pedro Jiménez','Sofía Torres'];
  const ids=['00001','00002','00003','00004','00005','00006'];
  const sid=S.sedes[0]?.id||'ds1';
  for(let i=0;i<names.length;i++){
    const e={id:ids[i],nombre:names[i],sedeId:sid,cargo:i<3?'Vendedor':'Administrativo',diasPersonal:null};
    await dbPut('empleados',e);if(!S.empleados.find(x=>x.id===ids[i]))S.empleados.push(e);
  }
  renderEmps();fillSelects();
}

// ══════════════════════════════════════════
//  PROCESAR
// ══════════════════════════════════════════
async function procesar(){
  if(!S.rawRows.length){toast('No hay CSV cargado','error');return;}
  if(!S.sedes.length){toast('Configura al menos una sede','error');return;}
  const pbar=document.getElementById('pbar'),pfill=document.getElementById('pfill');
  pbar.style.display='block';pfill.style.width='5%';
  const tc=detectCol(S.csvHdr,['time','fecha','datetime','date','hora']);
  const ic=detectCol(S.csvHdr,['card no','card_no','cardno','no.']);
  const nc=detectCol(S.csvHdr,['name','nombre','employee']);
  if(!tc){toast('No se encontró columna de tiempo','error');return;}

  const byKey={};
  S.rawRows.forEach(row=>{
    const rawId=(ic?row[ic]:'').toString().trim();
    const rawNom=(nc?row[nc]:'').toString().trim();
    const ek=rawId||rawNom;if(!ek)return;
    const dt=parseDT(row[tc]);if(!dt||isNaN(dt))return;
    const k=ek+'_'+dkey(dt);
    if(!byKey[k])byKey[k]={empId:rawId,empNom:rawNom,date:dt,marks:[]};
    // Deduplicate: skip if same minute already recorded
    const dtMin=dt.getHours()*60+dt.getMinutes();
    const alreadyHasMin=byKey[k].marks.some(m=>m.getHours()*60+m.getMinutes()===dtMin);
    if(!alreadyHasMin) byKey[k].marks.push(dt);
  });

  pfill.style.width='40%';
  const resultados=[];

  Object.values(byKey).forEach(({empId,empNom,date,marks})=>{
    marks.sort((a,b)=>a-b);
    const emp=S.empleados.find(e=>e.id===empId||e.nombre.toLowerCase()===empNom.toLowerCase());
    const sedeId=emp?.sedeId||S.sedes[0]?.id;
    const sede=S.sedes.find(s=>s.id===sedeId)||S.sedes[0];
    if(!sede)return;

    const dow=date.getDay();
    // Resolver horario vigente: B si la fecha cae en su rango, si no A, si no sede
    const fechaStr=dkey(date);
    let diasSrc=sede.dias||{};
    let horarioUsado='Sede';
    if(emp?.diasPersonal){
      diasSrc=emp.diasPersonal;
      horarioUsado='Horario A';
    }
    if(emp?.diasPersonalB&&emp?.hbDesde){
      const dentroB=fechaStr>=emp.hbDesde&&(!emp.hbHasta||fechaStr<=emp.hbHasta);
      if(dentroB){diasSrc=emp.diasPersonalB;horarioUsado='Horario B';}
    }
    const dc=diasSrc[dow];
    const esLab=!!(dc?.activo);

    const entM=t2m(dc?.entrada||'08:00');
    const salM=t2m(dc?.salida||'17:00');
    const almM=parseInt(dc?.almuerzo??60);
    const tolM=sede.tolerancia||0;
    const minExtM=sede.minExt??15;
    const antici=sede.antici||0;  // 1 = contar entrada anticipada como HH.EE
    const hnoct=t2m(sede.hnoct||'21:00');

    const p=marks[0],u=marks[marks.length-1];
    const pM=p.getHours()*60+p.getMinutes();
    const uM=u.getHours()*60+u.getMinutes();

    // Tardanza
    const tardanza=esLab?Math.max(0,pM-entM-tolM):0;

    // Sin salida: solo 1 marca única, o primera y última en mismo minuto
    const sinSalida = marks.length < 2 ||
      (marks[0].getHours()===marks[marks.length-1].getHours() &&
       marks[0].getMinutes()===marks[marks.length-1].getMinutes());

    // Almuerzo tomado
    let almTom=almM;
    if(marks.length>=4){
      let mg=0;
      for(let i=0;i<marks.length-1;i++){const g=(marks[i+1]-marks[i])/60000;if(g>15&&g>mg)mg=g;}
      if(mg>15)almTom=Math.round(mg);
    }

    // Horas trabajadas brutas
    const trabajadasBruto=sinSalida?null:uM-pM;
    // Trabajadas netas = bruto - almuerzo
    const trabajadasNeto=sinSalida?null:Math.max(0,trabajadasBruto-almTom);

    // HH.EE — minutos DESPUÉS de la salida + (si antici=1) minutos ANTES de la entrada
    const extrasSalida = sinSalida ? null : Math.max(0, uM - salM);
    const extrasEntrada = (antici && esLab && !sinSalida) ? Math.max(0, entM - pM) : 0;
    const rawExtras = sinSalida ? null : (extrasSalida||0) + extrasEntrada;
    const extrasMin = sinSalida ? null : (rawExtras >= minExtM ? rawExtras : 0);
    let heDiur=0, heNoct=0;
    if(extrasMin>0){
      // Diurnas: desde entrada anticipada hasta hora nocturna, y desde salida hasta nocturna
      const extEnd = uM; // hora de salida real
      const extStart = antici ? Math.min(pM, entM) : salM; // si anticipa, desde entrada real
      heDiur = extStart < hnoct ? Math.min(extrasMin, hnoct - extStart) : 0;
      heNoct = Math.max(0, extEnd - hnoct);
    }

    let estado='Normal';
    if(!esLab)estado='No laboral';
    else if(sinSalida)estado='Sin salida';
    else if(tardanza>0&&extrasMin>0)estado='Tardanza + Extras';
    else if(tardanza>0)estado='Tardanza';
    else if(extrasMin>0&&esFestivo(dkey(date)))estado='Festivo + Extras';
    else if(extrasMin>0)estado='Con extras';

    resultados.push({
      id:empId+'_'+dkey(date),
      empId,empNom:emp?.nombre||empNom,
      sedeId,sedeNom:sede.nombre,
      fecha:dkey(date),dia:DIAS_FULL[dow],diaCort:DIAS[dow],
      entrada:fmtT(p),
      salida:sinSalida?'--:--':fmtT(u),
      sinSalida,
      tardanza,almTom,
      trabajadasNeto,
      extrasMin:extrasMin??0,extrasEntrada,heDiur,heNoct,
      estado,esLab,
      horarioUsado,
      entradaEsperada:dc?.entrada||'--:--',
      salidaEsperada:dc?.salida||'--:--',
    });
  });

  pfill.style.width='75%';
  // ── MERGE: mantener resultados de otras sedes, reemplazar solo los del CSV actual ──
  // Identificar qué empIds y fechas trae este CSV
  const newIds=new Set(resultados.map(r=>r.id));
  // Filtrar los resultados existentes quitando solo los que se van a reemplazar
  const existentes=S.resultados.filter(r=>!newIds.has(r.id));
  const merged=[...existentes,...resultados];
  // Guardar en BD solo los nuevos/actualizados (los existentes ya están)
  for(const r of resultados) await dbPut('resultados',r);
  S.resultados=merged;

  const meses=[...new Set(resultados.map(r=>r.fecha.slice(0,7)))].sort();
  ['rf-mes','xf-mes'].forEach(id=>{
    const el=document.getElementById(id);if(!el)return;
    const v=el.value;el.innerHTML='<option value="">Todos los meses</option>';
    meses.forEach(m=>{const o=document.createElement('option');o.value=m;o.textContent=m;el.appendChild(o);});
    if(v)el.value=v;
  });
  fillSelects();
  pfill.style.width='100%';
  setTimeout(()=>pbar.style.display='none',600);
  renderStats();renderRes();
  spill('ok',`${merged.length} días en BD`);
  toast(`+${resultados.length} registros procesados · Total: ${merged.length}`,'success');
  go('resultados',document.querySelector('[data-sec=resultados]'));
}

// ══════════════════════════════════════════
//  STATS
// ══════════════════════════════════════════
function renderStats(){
  const r=S.resultados.filter(x=>x.esLab);
  const te=r.reduce((s,x)=>s+x.extrasMin,0);
  const tt=r.reduce((s,x)=>s+x.tardanza,0);
  const hd=r.reduce((s,x)=>s+x.heDiur,0);
  const hn=r.reduce((s,x)=>s+x.heNoct,0);
  document.getElementById('stats-area').innerHTML=`<div class="stats-grid">
    <div class="stat-card"><div class="stat-label">Días laborales</div><div class="stat-val">${r.length}</div><div class="stat-sub">procesados</div></div>
    <div class="stat-card red"><div class="stat-label">Total HH.EE</div><div class="stat-val">${(te/60).toFixed(1)}h</div><div class="stat-sub">${r.filter(x=>x.extrasMin>0).length} días con extras</div></div>
    <div class="stat-card amber"><div class="stat-label">Tardanzas</div><div class="stat-val">${r.filter(x=>x.tardanza>0).length}</div><div class="stat-sub">${m2hm(tt)} acumulado</div></div>
    <div class="stat-card blue"><div class="stat-label">HE Diurnas</div><div class="stat-val">${(hd/60).toFixed(1)}h</div><div class="stat-sub">${m2hm(hd)}</div></div>
    <div class="stat-card purple"><div class="stat-label">HE Nocturnas</div><div class="stat-val">${(hn/60).toFixed(1)}h</div><div class="stat-sub">${m2hm(hn)}</div></div>
  </div>`;
}

// ══════════════════════════════════════════
//  TABLA RESULTADOS
// ══════════════════════════════════════════
const EBADGE={
  'Normal':'<span class="badge badge-green">Normal</span>',
  'Con extras':'<span class="badge badge-amber">Con extras</span>',
  'Festivo + Extras':'<span class="badge badge-blue">Festivo+HE</span>',
  'Tardanza':'<span class="badge badge-red">Tardanza</span>',
  'Tardanza + Extras':'<span class="badge badge-red">T+HE</span>',
  'Sin salida':'<span class="badge badge-gray">Sin salida</span>',
  'Incompleto':'<span class="badge badge-gray">Incompleto</span>',
};
let _resPage=1, _resPageSize=15, _resFiltered=[];
function renderRes(resetPage=true){
  const q=(document.getElementById('rf-q')?.value||'').toLowerCase();
  const sfRaw2=document.getElementById('rf-sede')?.value||'';
  const sf=sfRaw2||sedeActiva||'';
  const mf=document.getElementById('rf-mes')?.value||'';
  const tf=document.getElementById('rf-tipo')?.value||'';
  const desde=document.getElementById('rf-desde')?.value||'';
  const hasta=document.getElementById('rf-hasta')?.value||'';
  _resFiltered=S.resultados.filter(r=>{
    if(!r.esLab)return false;
    if(sf&&r.sedeId!==sf)return false;
    if(mf&&!r.fecha.startsWith(mf))return false;
    if(desde&&r.fecha<desde)return false;
    if(hasta&&r.fecha>hasta)return false;
    if(q&&!r.empNom.toLowerCase().includes(q)&&!r.empId.includes(q))return false;
    if(tf==='extras'&&r.extrasMin<=0)return false;
    if(tf==='tardanza'&&r.tardanza<=0)return false;
    if(tf==='incompleto'&&r.estado!=='Sin salida')return false;
    return true;
  });
  if(resetPage) _resPage=1;
  const tb=document.getElementById('rtb');
  if(!_resFiltered.length){
    tb.innerHTML='<tr><td colspan="14"><div class="empty-state"><div class="eico">📊</div><p>Sin resultados para los filtros.</p></div></td></tr>';
    document.getElementById('res-pagination').innerHTML='';
    return;
  }
  const start=(_resPage-1)*_resPageSize;
  const page=_resFiltered.slice(start,start+_resPageSize);
  tb.innerHTML=page.map(r=>{
    const ss=r.sinSalida;
    const nota=_notas[r.id]||'';
    return`<tr>
      <td class="name">${r.empNom}</td>
      <td style="color:var(--muted2);font-size:11px;">${r.sedeNom}</td>
      <td style="font-size:11px;">${fmtFecha(r.fecha)}</td>
      <td style="color:var(--muted);">${r.diaCort}</td>
      <td>${r.entrada}</td>
      <td ${ss?'style="color:var(--muted)"':''}>${ss?'no marcó':''}${!ss?r.salida:''}</td>
      <td class="${!ss&&r.trabajadasNeto>0?'':'rn'}">${ss?'<span style="color:var(--muted)">no marcó salida</span>':m2hm(r.trabajadasNeto)}</td>
      <td class="${r.tardanza>0?'rt':'rn'}">${r.tardanza>0?m2hm(r.tardanza):'00h 00m'}</td>
      <td class="${r.almTom>0?'rg':'rn'}">${m2hm(r.almTom)}</td>
      <td class="${r.heDiur>0?'rx':'rn'}">${m2hm(r.heDiur)}</td>
      <td style="${r.heNoct>0?'color:var(--accent3);font-weight:600':'color:var(--muted2)'}">${m2hm(r.heNoct)}</td>
      <td class="${r.extrasMin>0?'rx':'rn'}" style="font-weight:700;">${ss?'<span style="color:var(--muted)">no marcó salida</span>':m2hm(r.extrasMin)}</td>
      <td>${EBADGE[r.estado]||r.estado}</td>
      <td><button class="nota-btn${nota?' has-nota':''}" onclick="openNotaPopup('${r.id}',event)" title="${nota||'Agregar nota'}">${nota?'📝 Ver':'＋ Nota'}</button></td>
    </tr>`;
  }).join('');
  renderPagination('res-pagination',_resFiltered.length,_resPage,_resPageSize,
    p=>{_resPage=p;renderRes(false);},
    sz=>{_resPageSize=sz;_resPage=1;renderRes(false);}
  );
}

// ══════════════════════════════════════════
//  CONSOLIDADO
// ══════════════════════════════════════════
function renderConsol(){
  if(!S.resultados.length)return;
  const m={};
  const sfC=sedeActiva||'';
  S.resultados.filter(r=>r.esLab&&(!sfC||r.sedeId===sfC)).forEach(r=>{
    if(!m[r.empId])m[r.empId]={nom:r.empNom,sede:r.sedeNom,ext:0,tard:0,hd:0,hn:0,dias:0};
    const e=m[r.empId];e.ext+=r.extrasMin;e.tard+=r.tardanza;e.hd+=r.heDiur;e.hn+=r.heNoct;e.dias++;
  });
  const rows=Object.values(m).sort((a,b)=>b.ext-a.ext);
  document.getElementById('consol-area').innerHTML=`<div class="table-wrap"><table>
    <thead><tr><th>Empleado</th><th>Sede</th><th>Días</th><th>HE Total</th><th>HE Diurnas</th><th>HE Nocturnas</th><th>Tardanzas</th></tr></thead>
    <tbody>${rows.map(r=>`<tr>
      <td class="name">${r.nom}</td>
      <td style="color:var(--muted2);font-size:11px;">${r.sede}</td>
      <td>${r.dias}</td>
      <td class="${r.ext>0?'rx':'rn'}" style="font-weight:700;">${m2hm(r.ext)}</td>
      <td class="${r.hd>0?'rx':'rn'}">${m2hm(r.hd)}</td>
      <td style="${r.hn>0?'color:var(--accent3)':'color:var(--muted2)'}">${m2hm(r.hn)}</td>
      <td class="${r.tard>0?'rt':'rn'}">${m2hm(r.tard)}</td>
    </tr>`).join('')}</tbody></table></div>`;
}

// ══════════════════════════════════════════
//  EXPORTAR
// ══════════════════════════════════════════
function getXData(){
  const sfRaw3=document.getElementById('xf-sede')?.value||'';
  const sf=sfRaw3||sedeActiva||'';
  const mf=document.getElementById('xf-mes')?.value||'';
  return S.resultados.filter(r=>r.esLab&&(!sf||r.sedeId===sf)&&(!mf||r.fecha.startsWith(mf)));
}

// ── Estilos Excel compartidos ──
function makeHeaderStyle(bgHex,fgHex='FFFFFF'){
  return{font:{bold:true,color:{rgb:fgHex}},fill:{fgColor:{rgb:bgHex}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{top:{style:'thin',color:{rgb:'CCCCCC'}},bottom:{style:'thin',color:{rgb:'CCCCCC'}},left:{style:'thin',color:{rgb:'CCCCCC'}},right:{style:'thin',color:{rgb:'CCCCCC'}}}};
}
function makeCellStyle(align='center'){
  return{alignment:{horizontal:align,vertical:'center'},border:{top:{style:'thin',color:{rgb:'EEEEEE'}},bottom:{style:{style:'thin'},color:{rgb:'EEEEEE'}},left:{style:'thin',color:{rgb:'EEEEEE'}},right:{style:'thin',color:{rgb:'EEEEEE'}}}};
}

// ── Hoja tipo "Reporte de Horas Extras" por empleado ──
function buildSheetEmpleado(empNom, rows, opts){
  const {incT,incG,incTrab}=opts;
  const aoa=[];
  aoa.push(['Reporte de Horas Extras']);
  aoa.push([`Colaborador: ${empNom}`]);
  aoa.push([]);
  const cab=['Fecha','Día de la Semana','Entrada','Salida'];
  if(incTrab) cab.push('Horas trabajadas');
  if(incT)    cab.push('Tardanza');
  cab.push('Horas extras');
  aoa.push(cab);
  for(const r of rows){
    const ss=r.sinSalida;
    const fila=[fmtFecha(r.fecha),r.dia,r.entrada,ss?'--:--':r.salida];
    if(incTrab) fila.push(ss?'no marcó':m2hm(r.trabajadasNeto));
    if(incT)    fila.push(r.tardanza>0?m2hm(r.tardanza):'00h 00m');
    fila.push(ss?'no marcó':m2hm(r.extrasMin));
    aoa.push(fila);
  }
  aoa.push([]);
  const totalDias=rows.length;
  const totalExtras=rows.filter(r=>!r.sinSalida).reduce((s,r)=>s+r.extrasMin,0);
  aoa.push([`Total de días registrados: ${totalDias}`]);
  aoa.push([`Total horas extras: ${m2hm(totalExtras)}`]);
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  const lastCol=cab.length-1;
  const ncols=cab.length;
  ws['!cols']=[{wch:14},{wch:18},{wch:10},{wch:10},{wch:18},{wch:14},{wch:16},{wch:16}].slice(0,ncols);
  ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:lastCol}},{s:{r:1,c:0},e:{r:1,c:lastCol}}];
  // Título
  const tc=ws['A1'];if(tc)tc.s={font:{bold:true,sz:14,color:{rgb:'FFFFFF'}},fill:{fgColor:{rgb:'1A3A2A'}},alignment:{horizontal:'center',vertical:'center'}};
  const cc=ws['A2'];if(cc)cc.s={font:{bold:true,sz:12,color:{rgb:'FFFFFF'}},fill:{fgColor:{rgb:'1A3A2A'}},alignment:{horizontal:'center',vertical:'center'}};
  // Cabecera fila 3
  for(let ci=0;ci<ncols;ci++){const a=XLSX.utils.encode_cell({r:3,c:ci});if(ws[a])ws[a].s={font:{bold:true,color:{rgb:'FFFFFF'},sz:11},fill:{fgColor:{rgb:'00A572'}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{bottom:{style:'medium',color:{rgb:'007A50'}}}};}
  // Filas de datos desde fila 4
  for(let ri=0;ri<rows.length;ri++){
    const r=rows[ri];const esTarde=r.tardanza>0;const isEven=ri%2===0;
    const bg=esTarde?'FFE5E5':(isEven?'F4FBF7':'FFFFFF');
    const fontColor=esTarde?'CC0000':'222222';
    for(let ci=0;ci<ncols;ci++){
      const a=XLSX.utils.encode_cell({r:4+ri,c:ci});
      if(ws[a])ws[a].s={fill:{fgColor:{rgb:bg}},font:{color:{rgb:fontColor},bold:ci===ncols-1&&r.extrasMin>0},alignment:{horizontal:ci<2?'left':'center',vertical:'center'},border:{bottom:{style:'thin',color:{rgb:'DDDDDD'}},right:{style:'thin',color:{rgb:'EEEEEE'}}}};
    }
  }
  // Totales
  const tr1=4+rows.length+1,tr2=4+rows.length+2;
  for(let ci=0;ci<ncols;ci++){
    const a1=XLSX.utils.encode_cell({r:tr1,c:ci});const a2=XLSX.utils.encode_cell({r:tr2,c:ci});
    if(ws[a1])ws[a1].s={font:{bold:true,color:{rgb:'1A3A2A'},sz:11},fill:{fgColor:{rgb:'E8F5EF'}},alignment:{horizontal:'left'}};
    if(ws[a2])ws[a2].s={font:{bold:true,color:{rgb:'007A50'},sz:11},fill:{fgColor:{rgb:'E8F5EF'}},alignment:{horizontal:'left'}};
  }
  return ws;
}

async function exportXLSX(){
  if(!S.resultados.length){toast('Sin datos','error');return;}
  const data=getXData();
  const formato=document.getElementById('xf-formato').value;
  const incT=document.getElementById('xc-t').checked;
  const incG=document.getElementById('xc-g').checked;
  const incD=document.getElementById('xc-d').checked;
  const incTrab=document.getElementById('xc-trab').checked;
  const opts={incT,incG,incD,incTrab};

  if(formato==='por_empleado'){
    // Group by sede → one workbook per sede, one sheet per employee
    const bySede={};
    data.forEach(r=>{
      if(!bySede[r.sedeId]) bySede[r.sedeId]={nom:r.sedeNom,emps:{}};
      if(!bySede[r.sedeId].emps[r.empId]) bySede[r.sedeId].emps[r.empId]={nom:r.empNom,rows:[]};
      bySede[r.sedeId].emps[r.empId].rows.push(r);
    });
    let totalEmps=0;
    for(const [,sedeData] of Object.entries(bySede)){
      const wb=XLSX.utils.book_new();
      for(const [,empData] of Object.entries(sedeData.emps)){
        empData.rows.sort((a,b)=>a.fecha.localeCompare(b.fecha));
        const ws=buildSheetEmpleado(empData.nom,empData.rows,opts);
        const shName=empData.nom.substring(0,31).replace(/[\\/\?\*\[\]:]/g,'_');
        XLSX.utils.book_append_sheet(wb,ws,shName);
        totalEmps++;
      }
      const fname='reporte_'+sedeData.nom.replace(/[^a-zA-Z0-9]/g,'_')+'_'+new Date().toISOString().slice(0,10)+'.xlsx';
      XLSX.writeFile(wb,fname,{cellStyles:true});
    }
    toast(`${totalEmps} empleados exportados en ${Object.keys(bySede).length} archivo(s)`,'success');

  } else {
    // Consolidado clásico
    const det=data.map(r=>{
      const ss=r.sinSalida;
      const o={'Empleado':r.empNom,'ID':r.empId,'Sede':r.sedeNom,'Fecha':fmtFecha(r.fecha),'Día':r.dia,'Entrada':r.entrada,'Salida':ss?'no marcó salida':r.salida};
      if(incTrab)o['Horas trabajadas']=ss?'no marcó salida':m2hm(r.trabajadasNeto);
      if(incT)o['Tardanza']=m2hm(r.tardanza);
      if(incG)o['Almuerzo (min)']=r.almTom;
      if(incD){o['HE Diurnas']=m2hm(r.heDiur);o['HE Nocturnas']=m2hm(r.heNoct);}
      o['Total HE']=ss?'no marcó salida':m2hm(r.extrasMin);
      o['Estado']=r.estado;
      return o;
    });
    const byEmp2={};
    data.forEach(r=>{
      if(!byEmp2[r.empId])byEmp2[r.empId]={'Empleado':r.empNom,'ID':r.empId,'Sede':r.sedeNom,'Días':0,'HE Total':'','HE Diurnas':'','HE Nocturnas':'','Tardanza total':'',_ext:0,_hd:0,_hn:0,_t:0};
      const e=byEmp2[r.empId];e['Días']++;e._ext+=r.extrasMin;e._hd+=r.heDiur;e._hn+=r.heNoct;e._t+=r.tardanza;
    });
    const cons=Object.values(byEmp2).map(e=>({...e,'HE Total':m2hm(e._ext),'HE Diurnas':m2hm(e._hd),'HE Nocturnas':m2hm(e._hn),'Tardanza total':m2hm(e._t)}));
    ['_ext','_hd','_hn','_t'].forEach(k=>cons.forEach(e=>delete e[k]));
    const wb=XLSX.utils.book_new();
    const wsDet=XLSX.utils.json_to_sheet(det);
    const wsCons=XLSX.utils.json_to_sheet(cons);
    // Estilos headers consolidado
    const detAoa=[Object.keys(det[0]||{}),...det.map(Object.values)];
    const consAoa=[Object.keys(cons[0]||{}),...cons.map(Object.values)];
    applyXlsxStylesFlat(wsDet,Object.keys(det[0]||{}).length);
    applyXlsxStylesFlat(wsCons,Object.keys(cons[0]||{}).length);
    XLSX.utils.book_append_sheet(wb,wsDet,'Detalle Diario');
    XLSX.utils.book_append_sheet(wb,wsCons,'Por Empleado');
    XLSX.writeFile(wb, `horas_extras_consolidado_${new Date().toISOString().slice(0,10)}.xlsx`, {cellStyles:true});
    toast('Excel consolidado exportado','success');
  }
}


// ══════════════════════════════════════════
//  FESTIVOS COLOMBIA (fijos + pascua)
// ══════════════════════════════════════════
function easterDate(y){
  // algoritmo de Meeus/Jones/Butcher
  const a=y%19,b=Math.floor(y/100),c=y%100;
  const d=Math.floor(b/4),e=b%4,f=Math.floor((b+8)/25);
  const g=Math.floor((b-f+1)/3),h=(19*a+b-d-g+15)%30;
  const i=Math.floor(c/4),k=c%4,l=(32+2*e+2*i-h-k)%7;
  const m=Math.floor((a+11*h+22*l)/451);
  const month=Math.floor((h+l-7*m+114)/31);
  const day=((h+l-7*m+114)%31)+1;
  return new Date(y,month-1,day);
}
function nextMonday(d){const r=new Date(d);const dow=r.getDay();if(dow===1)return r;r.setDate(r.getDate()+((8-dow)%7));return r;}
function getFestivos(y){
  const f=[];
  const add=(m,d)=>f.push(new Date(y,m-1,d));
  const addLunes=(m,d)=>f.push(nextMonday(new Date(y,m-1,d)));
  // Fijos
  add(1,1);add(5,1);add(7,20);add(8,7);add(12,8);add(12,25);
  // Emiliani (traslado a lunes si no caen en lunes)
  addLunes(1,6);addLunes(3,19);addLunes(6,29);addLunes(8,15);
  addLunes(10,12);addLunes(11,1);addLunes(11,11);
  // Semana santa / Pascua
  const pascua=easterDate(y);
  const juevesSanto=new Date(pascua);juevesSanto.setDate(pascua.getDate()-3);
  const viernesSanto=new Date(pascua);viernesSanto.setDate(pascua.getDate()-2);
  const ascension=new Date(pascua);ascension.setDate(pascua.getDate()+43);
  const corpusChristi=new Date(pascua);corpusChristi.setDate(pascua.getDate()+64);
  const sagradoCorazon=new Date(pascua);sagradoCorazon.setDate(pascua.getDate()+71);
  f.push(juevesSanto,viernesSanto);
  f.push(nextMonday(ascension),nextMonday(corpusChristi),nextMonday(sagradoCorazon));
  return f.map(d=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`);
}
// Cache festivos por año
const _festivosCache={};
function esFestivo(fecha){
  const y=parseInt(fecha.slice(0,4));
  if(!_festivosCache[y])_festivosCache[y]=new Set(getFestivos(y));
  return _festivosCache[y].has(fecha);
}

// ══════════════════════════════════════════
//  RECARGOS (CST Colombia)
//  HE diurna ord   +25%  → factor 1.25
//  HE nocturna ord +75%  → factor 1.75
//  HE dom/fest diurna   +100% → factor 2.00
//  HE dom/fest nocturna +150% → factor 2.50
// ══════════════════════════════════════════
function calcRecargos(r,vhora){
  // r = resultado de un día
  const fest=esFestivo(r.fecha)||new Date(r.fecha+'T00:00').getDay()===0;
  const factor_diur = fest ? 2.00 : 1.25;
  const factor_noct = fest ? 2.50 : 1.75;
  const valDiur  = vhora ? (r.heDiur/60)*vhora*factor_diur  : null;
  const valNoct  = vhora ? (r.heNoct/60)*vhora*factor_noct  : null;
  return {fest,factor_diur,factor_noct,valDiur,valNoct};
}

// ══════════════════════════════════════════
//  ALERTAS UMBRAL
// ══════════════════════════════════════════
const ALERTA_UMBRAL_HORAS = 20; // horas mensuales por empleado que disparan alerta

function checkAlertas(sf=''){
  if(!S.resultados.length) return [];
  const alertas=[];
  const base=S.resultados.filter(r=>r.esLab&&(!sf||r.sedeId===sf));
  // Alerta 1: Empleado supera umbral de HH.EE en el mes
  const byEmpMes={};
  base.forEach(r=>{
    const mes=r.fecha.slice(0,7);
    const k=r.empId+'_'+mes;
    if(!byEmpMes[k])byEmpMes[k]={nom:r.empNom,mes,ext:0};
    byEmpMes[k].ext+=r.extrasMin;
  });
  Object.values(byEmpMes).forEach(e=>{
    if(e.ext/60>=ALERTA_UMBRAL_HORAS)
      alertas.push({tipo:'umbral',msg:`⚠️ <strong>${e.nom}</strong> superó ${ALERTA_UMBRAL_HORAS}h de HH.EE en ${e.mes} (${(e.ext/60).toFixed(1)}h)`});
  });
  // Alerta 2: Empleados sin salida frecuente
  const byEmpSS={};
  base.filter(r=>r.sinSalida).forEach(r=>{if(!byEmpSS[r.empId])byEmpSS[r.empId]={nom:r.empNom,n:0};byEmpSS[r.empId].n++;});
  Object.values(byEmpSS).filter(e=>e.n>=3).forEach(e=>{
    alertas.push({tipo:'sin_salida',msg:`🔴 <strong>${e.nom}</strong> tiene ${e.n} días sin marcar salida`});
  });
  // Alerta 3: Festivos trabajados
  const festTrab=base.filter(r=>esFestivo(r.fecha)&&r.extrasMin>0);
  if(festTrab.length)
    alertas.push({tipo:'festivo',msg:`📅 Hay <strong>${festTrab.length}</strong> registro(s) con HH.EE en día festivo — requieren recargo especial`});
  return alertas;
}

// ══════════════════════════════════════════
//  DASHBOARD
// ══════════════════════════════════════════
let _dashChart=null;

function fillDashFilters(){
  const dSede=document.getElementById('df-sede');
  if(!dSede)return;
  // Sede: poblar y forzar sedeActiva si está seleccionada
  dSede.innerHTML='<option value="">Todas las sedes</option>';
  S.sedes.forEach(s=>{const o=document.createElement('option');o.value=s.id;o.textContent=s.nombre;dSede.appendChild(o);});
  if(sedeActiva) dSede.value=sedeActiva;

  // Mes: solo meses que existen en la sede activa
  const sfActual=dSede.value||'';
  const dMes=document.getElementById('df-mes');
  const mv=dMes.value;
  dMes.innerHTML='<option value="">Todos los meses</option>';
  const meses=[...new Set(
    S.resultados.filter(r=>!sfActual||r.sedeId===sfActual).map(r=>r.fecha.slice(0,7))
  )].sort();
  meses.forEach(m=>{const o=document.createElement('option');o.value=m;o.textContent=m;dMes.appendChild(o);});
  if(mv&&meses.includes(mv))dMes.value=mv;

  // Empleado calendario: solo los de la sede activa
  const dCalEmp=document.getElementById('df-cal-emp');
  const cv=dCalEmp.value;
  dCalEmp.innerHTML='<option value="">— Empleado —</option>';
  const emps=[...new Map(
    S.resultados.filter(r=>!sfActual||r.sedeId===sfActual).map(r=>[r.empId,r.empNom])
  ).entries()];
  emps.forEach(([id,nom])=>{const o=document.createElement('option');o.value=id;o.textContent=nom;dCalEmp.appendChild(o);});
  // Si el empleado seleccionado ya no pertenece a la sede activa, resetear
  if(cv&&emps.some(([id])=>id===cv)) dCalEmp.value=cv;
}

function getdashData(){
  const sf=document.getElementById('df-sede')?.value||sedeActiva||'';
  const mf=document.getElementById('df-mes')?.value||'';
  return S.resultados.filter(r=>r.esLab&&(!sf||r.sedeId===sf)&&(!mf||r.fecha.startsWith(mf)));
}

function renderDash(){
  if(!S.resultados.length){
    document.getElementById('dash-chart').style.display='none';
    document.getElementById('dash-chart-empty').style.display='block';
    document.getElementById('dash-ranking').innerHTML='<div class="empty-state"><div class="eico">🏆</div><p>Sin datos.</p></div>';
    document.getElementById('dash-calendar').innerHTML='<div class="empty-state"><p>Sin datos.</p></div>';
    document.getElementById('dash-recargos').innerHTML='<div class="empty-state"><p>Sin datos.</p></div>';
    document.getElementById('dash-alerts-area').innerHTML='';
    document.getElementById('dash-stats-area').innerHTML='';
    return;
  }
  fillDashFilters();
  const data=getdashData();
  if(!data.length){
    document.getElementById('dash-chart').style.display='none';
    document.getElementById('dash-chart-empty').style.display='block';
    return;
  }

  // ── Alertas ──
  const sf=document.getElementById('df-sede')?.value||sedeActiva||'';
  const alertas=checkAlertas(sf);
  document.getElementById('dash-alerts-area').innerHTML = alertas.length
    ? `<div style="display:flex;flex-direction:column;gap:8px;margin-bottom:18px;">`
      +alertas.map(a=>`<div style="background:${a.tipo==='umbral'?'rgba(245,158,11,.1)':a.tipo==='festivo'?'rgba(79,95,196,.1)':'rgba(220,38,38,.1)'};border:1px solid ${a.tipo==='umbral'?'var(--warn)':a.tipo==='festivo'?'var(--accent3)':'var(--danger)'};border-radius:8px;padding:10px 14px;font-size:12px;font-family:'DM Mono',monospace;color:var(--text);">${a.msg}</div>`).join('')
      +`</div>`
    : '';

  // ── Stats mini ──
  const te=data.reduce((s,x)=>s+x.extrasMin,0);
  const festDias=data.filter(r=>esFestivo(r.fecha));
  document.getElementById('dash-stats-area').innerHTML=`<div class="stats-grid" style="margin-bottom:18px;">
    <div class="stat-card"><div class="stat-label">Registros filtrados</div><div class="stat-val">${data.length}</div><div class="stat-sub">días laborales</div></div>
    <div class="stat-card red"><div class="stat-label">HH.EE período</div><div class="stat-val">${(te/60).toFixed(1)}h</div><div class="stat-sub">${data.filter(x=>x.extrasMin>0).length} días con extras</div></div>
    <div class="stat-card blue"><div class="stat-label">Días festivos</div><div class="stat-val">${festDias.length}</div><div class="stat-sub">en el período</div></div>
  </div>`;

  // ── Gráfica barras por empleado ──
  const byEmp={};
  data.forEach(r=>{
    if(!byEmp[r.empId])byEmp[r.empId]={nom:r.empNom,ext:0,diur:0,noct:0};
    byEmp[r.empId].ext+=r.extrasMin;
    byEmp[r.empId].diur+=r.heDiur;
    byEmp[r.empId].noct+=r.heNoct;
  });
  const sorted=Object.values(byEmp).sort((a,b)=>b.ext-a.ext).slice(0,15);
  const labels=sorted.map(e=>e.nom.split(' ')[0]+(e.nom.split(' ')[1]?' '+e.nom.split(' ')[1][0]+'.':''));
  const isDark=document.documentElement.getAttribute('data-theme')==='dark';
  const gridColor=isDark?'rgba(255,255,255,0.07)':'rgba(0,0,0,0.07)';
  const textColor=isDark?'#9ca3af':'#6b7280';
  document.getElementById('dash-chart').style.display='block';
  document.getElementById('dash-chart-empty').style.display='none';
  if(_dashChart)_dashChart.destroy();
  _dashChart=new Chart(document.getElementById('dash-chart'),{
    type:'bar',
    data:{
      labels,
      datasets:[
        {label:'HE Diurnas',data:sorted.map(e=>+(e.diur/60).toFixed(2)),backgroundColor:'rgba(0,229,160,0.75)',borderRadius:4},
        {label:'HE Nocturnas',data:sorted.map(e=>+(e.noct/60).toFixed(2)),backgroundColor:'rgba(123,140,255,0.75)',borderRadius:4},
      ]
    },
    options:{
      responsive:true,maintainAspectRatio:true,
      plugins:{legend:{labels:{color:textColor,font:{family:'DM Mono, monospace',size:11}}},tooltip:{callbacks:{label:ctx=>`${ctx.dataset.label}: ${ctx.parsed.y}h`}}},
      scales:{
        x:{stacked:true,ticks:{color:textColor,font:{family:'DM Mono',size:11}},grid:{color:gridColor}},
        y:{stacked:true,ticks:{color:textColor,font:{family:'DM Mono',size:11},callback:v=>v+'h'},grid:{color:gridColor}}
      }
    }
  });

  // ── Ranking ──
  document.getElementById('dash-ranking').innerHTML=sorted.length
    ? sorted.map((e,i)=>`<div style="display:flex;align-items:center;gap:10px;padding:7px 0;border-bottom:1px solid var(--border);">
        <span style="font-size:13px;font-family:'DM Mono',monospace;color:var(--muted);width:22px;">${i+1}</span>
        <span style="font-family:'Syne',sans-serif;font-weight:700;flex:1;font-size:13px;">${e.nom}</span>
        <span style="font-family:'DM Mono',monospace;font-size:12px;color:${e.ext>0?'var(--accent)':'var(--muted2)'};">${(e.ext/60).toFixed(1)}h</span>
        <div style="width:60px;background:var(--border);border-radius:4px;height:6px;overflow:hidden;"><div style="width:${Math.min(100,(e.ext/sorted[0].ext*100)).toFixed(0)}%;height:100%;background:var(--accent);border-radius:4px;"></div></div>
      </div>`).join('')
    : '<div class="empty-state"><p>Sin datos.</p></div>';

  // ── Calendario ──
  renderCalendar();

  // ── Recargos ──
  renderRecargos(data);
}

function renderCalendar(){
  const empId=document.getElementById('df-cal-emp')?.value||'';
  const mf=document.getElementById('df-mes')?.value||'';
  const sf=document.getElementById('df-sede')?.value||sedeActiva||'';
  const calEl=document.getElementById('dash-calendar');
  if(!empId){calEl.innerHTML='<p style="font-size:12px;color:var(--muted);font-family:\'DM Mono\',monospace;padding:10px 0;">Selecciona un empleado.</p>';return;}
  const empData=S.resultados.filter(r=>r.empId===empId&&r.esLab&&(!mf||r.fecha.startsWith(mf))&&(!sf||r.sedeId===sf));
  if(!empData.length){calEl.innerHTML='<div class="empty-state"><p>Sin registros para este filtro.</p></div>';return;}
  // Group by month
  const byMes={};
  empData.forEach(r=>{const m=r.fecha.slice(0,7);if(!byMes[m])byMes[m]=[];byMes[m].push(r);});
  let html='';
  for(const [mes,dias] of Object.entries(byMes).sort()){
    const [y,m]=mes.split('-').map(Number);
    const firstDay=new Date(y,m-1,1).getDay();
    const daysInMonth=new Date(y,m,0).getDate();
    const byDay={};dias.forEach(r=>byDay[parseInt(r.fecha.slice(8))]=r);
    const festivos=new Set(getFestivos(y).filter(f=>f.startsWith(mes)).map(f=>parseInt(f.slice(8))));
    html+=`<div style="margin-bottom:16px;">
      <div style="font-size:11px;font-family:'DM Mono',monospace;color:var(--muted2);margin-bottom:8px;letter-spacing:1px;">${mes}</div>
      <div style="display:grid;grid-template-columns:repeat(7,1fr);gap:3px;text-align:center;">
        ${['D','L','M','X','J','V','S'].map(d=>`<div style="font-size:9px;color:var(--muted);font-family:'DM Mono',monospace;padding:2px;">${d}</div>`).join('')}
        ${Array(firstDay).fill('<div></div>').join('')}
        ${Array.from({length:daysInMonth},(_,i)=>{
          const d=i+1;const r=byDay[d];const fest=festivos.has(d);
          const dow=new Date(y,m-1,d).getDay();const isWeekend=dow===0||dow===6;
          let bg='transparent',color='var(--muted)',border='transparent',title='';
          if(r){
            if(r.extrasMin>0){bg=fest?'rgba(123,140,255,0.25)':'rgba(0,229,160,0.2)';color=fest?'var(--accent3)':'var(--accent)';border=fest?'var(--accent3)':'var(--accent)';}
            else if(r.tardanza>0){bg='rgba(255,79,79,0.15)';color='var(--danger)';border='var(--danger)';}
            else if(r.sinSalida){bg='rgba(107,114,128,0.15)';color='var(--muted2)';}
            else{bg='rgba(16,185,129,0.1)';color='var(--success)';}
            title=`${r.entrada}–${r.salida} HE:${m2hm(r.extrasMin)}`;
          } else if(fest){bg='rgba(123,140,255,0.1)';color='var(--accent3)';}
          else if(isWeekend){color='var(--muted)';}
          return`<div title="${title}" style="font-size:10px;font-family:'DM Mono',monospace;padding:3px 1px;border-radius:4px;background:${bg};color:${color};border:1px solid ${border};cursor:${r?'pointer':'default'};">${d}${fest?'*':''}</div>`;
        }).join('')}
      </div>
      <div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:6px;font-size:9px;font-family:'DM Mono',monospace;color:var(--muted);">
        <span style="color:var(--accent);">■ Con extras</span>
        <span style="color:var(--accent3);">■ Festivo+extras</span>
        <span style="color:var(--danger);">■ Tardanza</span>
        <span style="color:var(--success);">■ Normal</span>
        <span>* Festivo</span>
      </div>
    </div>`;
  }
  calEl.innerHTML=html;
}

function renderRecargos(data){
  // Agrupar por empleado y calcular recargos
  const byEmp={};
  data.forEach(r=>{
    const sede=S.sedes.find(s=>s.id===r.sedeId);
    const vhora=sede?.vhora||0;
    const rec=calcRecargos(r,vhora);
    if(!byEmp[r.empId])byEmp[r.empId]={nom:r.empNom,sede:r.sedeNom,vhora,
      heDiurOrd:0,heNoctOrd:0,heDiurFest:0,heNoctFest:0,
      valDiurOrd:0,valNoctOrd:0,valDiurFest:0,valNoctFest:0};
    const e=byEmp[r.empId];
    if(rec.fest){e.heDiurFest+=r.heDiur;e.heNoctFest+=r.heNoct;}
    else{e.heDiurOrd+=r.heDiur;e.heNoctOrd+=r.heNoct;}
    if(vhora){
      if(rec.fest){e.valDiurFest+=(r.heDiur/60)*vhora*2.00;e.valNoctFest+=(r.heNoct/60)*vhora*2.50;}
      else{e.valDiurOrd+=(r.heDiur/60)*vhora*1.25;e.valNoctOrd+=(r.heNoct/60)*vhora*1.75;}
    }
  });
  const rows=Object.values(byEmp).filter(e=>e.heDiurOrd+e.heNoctOrd+e.heDiurFest+e.heNoctFest>0);
  const fmtCOP=v=>v>0?new Intl.NumberFormat('es-CO',{style:'currency',currency:'COP',minimumFractionDigits:0}).format(v):'—';
  if(!rows.length){document.getElementById('dash-recargos').innerHTML='<div class="empty-state"><p>Sin horas extras en el período.</p></div>';return;}
  document.getElementById('dash-recargos').innerHTML=`<div class="table-wrap"><table>
    <thead><tr>
      <th>Empleado</th><th>Sede</th>
      <th>HE Diur. Ord. <span style="color:var(--accent);font-size:9px;">(+25%)</span></th>
      <th>HE Noct. Ord. <span style="color:var(--accent3);font-size:9px;">(+75%)</span></th>
      <th>HE Diur. Fest. <span style="color:var(--warn);font-size:9px;">(+100%)</span></th>
      <th>HE Noct. Fest. <span style="color:var(--danger);font-size:9px;">(+150%)</span></th>
      <th>Valor total</th>
    </tr></thead>
    <tbody>${rows.sort((a,b)=>(b.heDiurOrd+b.heNoctOrd+b.heDiurFest+b.heNoctFest)-(a.heDiurOrd+a.heNoctOrd+a.heDiurFest+a.heNoctFest)).map(e=>{
      const totalMin=e.heDiurOrd+e.heNoctOrd+e.heDiurFest+e.heNoctFest;
      const totalVal=e.valDiurOrd+e.valNoctOrd+e.valDiurFest+e.valNoctFest;
      return`<tr>
        <td class="name">${e.nom}</td>
        <td style="color:var(--muted2);font-size:11px;">${e.sede}</td>
        <td class="${e.heDiurOrd>0?'rx':'rn'}">${m2hm(e.heDiurOrd)}${e.vhora&&e.valDiurOrd>0?`<br><span style="font-size:10px;color:var(--muted2);">${fmtCOP(e.valDiurOrd)}</span>`:''}</td>
        <td style="${e.heNoctOrd>0?'color:var(--accent3);font-weight:600':'color:var(--muted2)'}">${m2hm(e.heNoctOrd)}${e.vhora&&e.valNoctOrd>0?`<br><span style="font-size:10px;color:var(--muted2);">${fmtCOP(e.valNoctOrd)}</span>`:''}</td>
        <td style="${e.heDiurFest>0?'color:var(--warn);font-weight:600':'color:var(--muted2)'}">${m2hm(e.heDiurFest)}${e.vhora&&e.valDiurFest>0?`<br><span style="font-size:10px;color:var(--muted2);">${fmtCOP(e.valDiurFest)}</span>`:''}</td>
        <td style="${e.heNoctFest>0?'color:var(--danger);font-weight:600':'color:var(--muted2)'}">${m2hm(e.heNoctFest)}${e.vhora&&e.valNoctFest>0?`<br><span style="font-size:10px;color:var(--muted2);">${fmtCOP(e.valNoctFest)}</span>`:''}</td>
        <td style="font-weight:700;color:${totalMin>0?'var(--text)':'var(--muted2)'};">${m2hm(totalMin)}${e.vhora&&totalVal>0?`<br><span style="font-size:11px;color:var(--accent);">${fmtCOP(totalVal)}</span>`:''}</td>
      </tr>`;
    }).join('')}</tbody>
  </table></div>`;
}

function exportCSV(){
  if(!S.resultados.length){toast('Sin datos','error');return;}
  const data=getXData();
  const heads=['Empleado','ID','Sede','Fecha','Día','Entrada','Salida','Horas trabajadas','Tardanza','Almuerzo min','HE Diurnas','HE Nocturnas','Total HE','Estado'];
  let csv=heads.join(',')+'\n';
  data.forEach(r=>{
    const ss=r.sinSalida;
    csv+=`"${r.empNom}","${r.empId}","${r.sedeNom}","${fmtFecha(r.fecha)}","${r.dia}","${r.entrada}","${ss?'no marcó salida':r.salida}","${ss?'no marcó salida':m2hm(r.trabajadasNeto)}","${m2hm(r.tardanza)}","${r.almTom}","${m2hm(r.heDiur)}","${m2hm(r.heNoct)}","${ss?'no marcó salida':m2hm(r.extrasMin)}","${r.estado}"\n`;
  });
  const a=document.createElement('a');
  a.href=URL.createObjectURL(new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8;'}));
  a.download=`horas_extras_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  toast('CSV exportado','success');
}

// ══════════════════════════════════════════
//  REPORTE DE LLEGADAS — filtros
// ══════════════════════════════════════════
function fillLlegadasFilters(){
  const lSede=document.getElementById('lf-sede');
  if(!lSede)return;
  const sv=lSede.value;
  lSede.innerHTML='<option value="">Todas las sedes</option>';
  S.sedes.forEach(s=>{const o=document.createElement('option');o.value=s.id;o.textContent=s.nombre;lSede.appendChild(o);});
  lSede.value=sedeActiva||sv||'';
  const lMes=document.getElementById('lf-mes');
  if(!lMes)return;
  const mv=lMes.value;
  lMes.innerHTML='<option value="">Todos los meses</option>';
  const meses=[...new Set(S.resultados.map(r=>r.fecha.slice(0,7)))].sort();
  meses.forEach(m=>{const o=document.createElement('option');o.value=m;o.textContent=m;lMes.appendChild(o);});
  if(mv)lMes.value=mv;
}

// ── Paginación helper ──
// Usa un dispatcher global en lugar de toString() de funciones
const _paginationHandlers = {};
function renderPagination(containerId, total, currentPage, pageSize, onPage, onSize){
  const totalPages = Math.ceil(total / pageSize);
  const el = document.getElementById(containerId);
  if(!el) return;
  if(totalPages <= 1 && total <= pageSize){ el.innerHTML=''; return; }
  // Guardar handlers por containerId
  _paginationHandlers[containerId] = { onPage, onSize };
  const start = (currentPage-1)*pageSize + 1;
  const end   = Math.min(currentPage*pageSize, total);
  // Generar botones de página
  let pages = [];
  if(totalPages <= 7){ for(let i=1;i<=totalPages;i++) pages.push(i); }
  else{
    pages = [1];
    if(currentPage > 3) pages.push('…');
    for(let i=Math.max(2,currentPage-1); i<=Math.min(totalPages-1,currentPage+1); i++) pages.push(i);
    if(currentPage < totalPages-2) pages.push('…');
    pages.push(totalPages);
  }
  el.innerHTML = `
    <button class="page-btn" ${currentPage===1?'disabled':''} onclick="_pgGo('${containerId}',${currentPage-1})">‹</button>
    ${pages.map(p => p==='…'
      ? `<span style="color:var(--muted);font-family:'DM Mono',monospace;padding:0 4px;">…</span>`
      : `<button class="page-btn${p===currentPage?' active':''}" onclick="_pgGo('${containerId}',${p})">${p}</button>`
    ).join('')}
    <button class="page-btn" ${currentPage===totalPages?'disabled':''} onclick="_pgGo('${containerId}',${currentPage+1})">›</button>
    <span class="page-info">${start}–${end} de ${total}</span>
    <select class="page-size-sel" onchange="_pgSize('${containerId}',parseInt(this.value))">
      ${[15,25,50,100].map(s=>`<option value="${s}"${s===pageSize?' selected':''}>${s} / pág</option>`).join('')}
    </select>`;
}
function _pgGo(id, page){
  const h = _paginationHandlers[id];
  if(h) h.onPage(page);
}
function _pgSize(id, size){
  const h = _paginationHandlers[id];
  if(h) h.onSize(size);
}

let _llegPage=1, _llegPageSize=15, _llegFiltered=[];
function renderLlegadas(resetPage=true){
  const q=(document.getElementById('lf-q')?.value||'').toLowerCase();
  const sf=document.getElementById('lf-sede')?.value||sedeActiva||'';
  const mf=document.getElementById('lf-mes')?.value||'';
  const tf=document.getElementById('lf-tipo')?.value||'';

  _llegFiltered=S.resultados.filter(r=>{
    if(!r.esLab)return false;
    if(sf&&r.sedeId!==sf)return false;
    if(mf&&!r.fecha.startsWith(mf))return false;
    const desde=document.getElementById('lf-desde')?.value||'';
    const hasta=document.getElementById('lf-hasta')?.value||'';
    if(desde&&r.fecha<desde)return false;
    if(hasta&&r.fecha>hasta)return false;
    if(q&&!r.empNom.toLowerCase().includes(q)&&!r.empId.includes(q))return false;
    if(tf==='puntual'&&r.tardanza>0)return false;
    if(tf==='tarde'&&r.tardanza<=0)return false;
    if(tf==='sin_salida'&&!r.sinSalida)return false;
    return true;
  });
  if(resetPage) _llegPage=1;

  // ── Stats ──
  const total=_llegFiltered.length;
  const puntuales=_llegFiltered.filter(r=>r.tardanza===0&&!r.sinSalida).length;
  const tardes=_llegFiltered.filter(r=>r.tardanza>0).length;
  const sinSal=_llegFiltered.filter(r=>r.sinSalida).length;
  const tardTotal=_llegFiltered.reduce((s,r)=>s+r.tardanza,0);
  document.getElementById('llegadas-stats-area').innerHTML=`<div class="stats-grid" style="margin-bottom:18px;">
    <div class="stat-card"><div class="stat-label">Días analizados</div><div class="stat-val">${total}</div><div class="stat-sub">días laborales</div></div>
    <div class="stat-card" style="--ac:var(--success);"><div class="stat-label">Puntuales</div><div class="stat-val" style="color:var(--success);">${puntuales}</div><div class="stat-sub">${total?((puntuales/total)*100).toFixed(0):0}% del total</div></div>
    <div class="stat-card amber"><div class="stat-label">Con tardanza</div><div class="stat-val">${tardes}</div><div class="stat-sub">${m2hm(tardTotal)} acumulado</div></div>
    <div class="stat-card red"><div class="stat-label">Sin salida</div><div class="stat-val">${sinSal}</div><div class="stat-sub">no marcaron salida</div></div>
  </div>`;

  // ── Tabla detalle paginada ──
  const tb=document.getElementById('llegadas-tb');
  if(!_llegFiltered.length){
    tb.innerHTML='<tr><td colspan="12"><div class="empty-state"><div class="eico">🕐</div><p>Sin registros para los filtros.</p></div></td></tr>';
    document.getElementById('llegadas-pagination').innerHTML='';
    return;
  }

  const fmtDiff=(min,tipo)=>{
    if(min===0)return`<span style="color:var(--success);">✓ Puntual</span>`;
    const abs=Math.abs(min);const hm=m2hm(abs);
    if(tipo==='entrada'){
      return min>0?`<span style="color:var(--danger);">+${hm} tarde</span>`:`<span style="color:var(--accent);">-${hm} antes</span>`;
    } else {
      return min>0?`<span style="color:var(--accent);">+${hm} extra</span>`:`<span style="color:var(--warn);">-${hm} antes</span>`;
    }
  };

  const start=(_llegPage-1)*_llegPageSize;
  const page=_llegFiltered.slice(start,start+_llegPageSize);
  tb.innerHTML=page.map(r=>{
    const entEsp=r.entradaEsperada||'--:--';
    const salEsp=r.salidaEsperada||'--:--';
    const difEnt=r.tardanza;
    const difSal=r.sinSalida?null:(t2m(r.salida)-t2m(salEsp));
    const horBadge=r.horarioUsado==='Horario B'?`<span class="badge badge-blue">B</span>`:r.horarioUsado==='Horario A'?`<span class="badge badge-gray">A</span>`:`<span class="badge badge-gray">Sede</span>`;
    return`<tr>
      <td class="name">${r.empNom}</td>
      <td style="color:var(--muted2);font-size:11px;">${r.sedeNom}</td>
      <td style="font-size:11px;">${fmtFecha(r.fecha)}</td>
      <td style="color:var(--muted);">${r.diaCort}</td>
      <td style="font-family:'DM Mono',monospace;">${r.entrada}</td>
      <td style="font-family:'DM Mono',monospace;color:var(--muted2);">${entEsp}</td>
      <td>${fmtDiff(difEnt,'entrada')}</td>
      <td style="font-family:'DM Mono',monospace;">${r.sinSalida?'<span style="color:var(--muted)">no marcó</span>':r.salida}</td>
      <td style="font-family:'DM Mono',monospace;color:var(--muted2);">${salEsp}</td>
      <td>${r.sinSalida?'<span style="color:var(--muted)">—</span>':fmtDiff(difSal,'salida')}</td>
      <td>${horBadge}</td>
      <td>${EBADGE[r.estado]||r.estado}</td>
    </tr>`;
  }).join('');
  renderPagination('llegadas-pagination',_llegFiltered.length,_llegPage,_llegPageSize,
    p=>{_llegPage=p;renderLlegadas(false);},
    sz=>{_llegPageSize=sz;_llegPage=1;renderLlegadas(false);}
  );

  // ── Resumen por empleado (usa todos los filtrados, no solo la página) ──
  const byEmp={};
  _llegFiltered.forEach(r=>{
    if(!byEmp[r.empId])byEmp[r.empId]={nom:r.empNom,sede:r.sedeNom,dias:0,puntuales:0,tardes:0,sinSal:0,tardTotal:0,usaB:false};
    const e=byEmp[r.empId];
    e.dias++;
    if(r.tardanza===0&&!r.sinSalida)e.puntuales++;
    if(r.tardanza>0)e.tardes++;
    if(r.sinSalida)e.sinSal++;
    e.tardTotal+=r.tardanza;
    if(r.horarioUsado==='Horario B')e.usaB=true;
  });
  const resRows=Object.values(byEmp).sort((a,b)=>b.tardes-a.tardes);
  document.getElementById('llegadas-resumen').innerHTML=`<div class="table-wrap"><table>
    <thead><tr>
      <th>Empleado</th><th>Sede</th><th>Días</th>
      <th>Puntuales</th><th>% Puntualidad</th>
      <th>Con tardanza</th><th>Tardanza total</th>
      <th>Sin salida</th><th>Horario</th>
    </tr></thead>
    <tbody>${resRows.map(e=>{
      const pct=e.dias?((e.puntuales/e.dias)*100).toFixed(0):0;
      const pctColor=parseInt(pct)>=90?'var(--success)':parseInt(pct)>=70?'var(--warn)':'var(--danger)';
      return`<tr>
        <td class="name">${e.nom}</td>
        <td style="color:var(--muted2);font-size:11px;">${e.sede}</td>
        <td>${e.dias}</td>
        <td style="color:var(--success);">${e.puntuales}</td>
        <td><span style="font-weight:700;color:${pctColor};">${pct}%</span>
          <div style="width:50px;background:var(--border);border-radius:3px;height:4px;display:inline-block;vertical-align:middle;margin-left:6px;overflow:hidden;">
            <div style="width:${pct}%;height:100%;background:${pctColor};border-radius:3px;"></div>
          </div></td>
        <td class="${e.tardes>0?'rt':'rn'}">${e.tardes}</td>
        <td class="${e.tardTotal>0?'rt':'rn'}">${m2hm(e.tardTotal)}</td>
        <td style="${e.sinSal>0?'color:var(--muted2)':''}">${e.sinSal}</td>
        <td>${e.usaB?'<span class="badge badge-blue">A+B</span>':'<span class="badge badge-gray">Std</span>'}</td>
      </tr>`;
    }).join('')}</tbody>
  </table></div>`;
}

// ══════════════════════════════════════════
//  EXPORT LLEGADAS XLSX — con diseño
// ══════════════════════════════════════════
function exportLlegadasXLSX(){
  if(!_llegFiltered.length){toast('Sin datos de llegadas','error');return;}
  const data=_llegFiltered;
  const wb=XLSX.utils.book_new();

  // ── Hoja 1: Detalle ──
  const aoa=[];
  aoa.push(['REPORTE DE LLEGADAS Y PUNTUALIDAD']);
  aoa.push([`Generado: ${new Date().toLocaleDateString('es-CO')}`]);
  aoa.push([]);
  aoa.push(['Empleado','Sede','Fecha','Día','Entrada real','Entrada esperada','Dif. entrada','Salida real','Salida esperada','Dif. salida','Horario','Estado']);
  data.forEach(r=>{
    const entEsp=r.entradaEsperada||'--:--';
    const salEsp=r.salidaEsperada||'--:--';
    const difEnt=r.tardanza===0?'Puntual':(r.tardanza>0?`+${m2hm(r.tardanza)} tarde`:`-${m2hm(Math.abs(r.tardanza))} antes`);
    const difSal=r.sinSalida?'no marcó':((t2m(r.salida)-t2m(salEsp))===0?'Puntual':((t2m(r.salida)-t2m(salEsp))>0?`+${m2hm(t2m(r.salida)-t2m(salEsp))} extra`:`-${m2hm(Math.abs(t2m(r.salida)-t2m(salEsp)))} antes`));
    aoa.push([r.empNom,r.sedeNom,fmtFecha(r.fecha),r.dia,r.entrada,entEsp,difEnt,r.sinSalida?'no marcó':r.salida,salEsp,difSal,r.horarioUsado||'Sede',r.estado]);
  });
  const ws1=XLSX.utils.aoa_to_sheet(aoa);
  ws1['!cols']=[{wch:22},{wch:16},{wch:12},{wch:10},{wch:13},{wch:15},{wch:18},{wch:12},{wch:14},{wch:18},{wch:10},{wch:16}];
  ws1['!merges']=[{s:{r:0,c:0},e:{r:0,c:11}},{s:{r:1,c:0},e:{r:1,c:11}}];
  // Estilos fila título
  applyXlsxStyles(ws1, aoa);
  XLSX.utils.book_append_sheet(wb,ws1,'Detalle Llegadas');

  // ── Hoja 2: Resumen por empleado ──
  const byEmp={};
  data.forEach(r=>{
    if(!byEmp[r.empId])byEmp[r.empId]={nom:r.empNom,sede:r.sedeNom,dias:0,puntuales:0,tardes:0,sinSal:0,tardTotal:0};
    const e=byEmp[r.empId];e.dias++;
    if(r.tardanza===0&&!r.sinSalida)e.puntuales++;
    if(r.tardanza>0)e.tardes++;
    if(r.sinSalida)e.sinSal++;
    e.tardTotal+=r.tardanza;
  });
  const aoa2=[];
  aoa2.push(['RESUMEN DE PUNTUALIDAD POR EMPLEADO']);
  aoa2.push([`Generado: ${new Date().toLocaleDateString('es-CO')}`]);
  aoa2.push([]);
  aoa2.push(['Empleado','Sede','Días registrados','Puntuales','% Puntualidad','Con tardanza','Tardanza total','Sin salida']);
  Object.values(byEmp).sort((a,b)=>b.tardes-a.tardes).forEach(e=>{
    const pct=e.dias?((e.puntuales/e.dias)*100).toFixed(1):0;
    aoa2.push([e.nom,e.sede,e.dias,e.puntuales,pct+'%',e.tardes,m2hm(e.tardTotal),e.sinSal]);
  });
  const ws2=XLSX.utils.aoa_to_sheet(aoa2);
  ws2['!cols']=[{wch:22},{wch:16},{wch:16},{wch:10},{wch:14},{wch:14},{wch:14},{wch:10}];
  ws2['!merges']=[{s:{r:0,c:0},e:{r:0,c:7}},{s:{r:1,c:0},e:{r:1,c:7}}];
  applyXlsxStyles(ws2, aoa2);
  XLSX.utils.book_append_sheet(wb,ws2,'Resumen Empleados');

  XLSX.writeFile(wb, `reporte_llegadas_${new Date().toISOString().slice(0,10)}.xlsx`, {cellStyles:true});
  toast('Reporte de llegadas exportado','success');
}

// ── Estilos para hojas generadas con json_to_sheet (sin filas de título) ──
function applyXlsxStylesFlat(ws, numCols){
  // Header row = row 0
  for(let c=0;c<numCols;c++){
    const addr=XLSX.utils.encode_cell({r:0,c});
    if(ws[addr]){ws[addr].s={font:{bold:true,color:{rgb:'FFFFFF'},sz:11},fill:{fgColor:{rgb:'00A572'}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{bottom:{style:'medium',color:{rgb:'007A50'}}}};}
  }
  // Data rows — alternate
  if(!ws['!ref'])return;
  const range=XLSX.utils.decode_range(ws['!ref']);
  for(let row=1;row<=range.e.r;row++){
    const isEven=row%2===0;
    for(let c=0;c<=range.e.c;c++){
      const addr=XLSX.utils.encode_cell({r:row,c});
      if(ws[addr]){ws[addr].s={fill:{fgColor:{rgb:isEven?'F4FBF7':'FFFFFF'}},alignment:{horizontal:c===0?'left':'center',vertical:'center'},border:{bottom:{style:'thin',color:{rgb:'DDDDDD'}},right:{style:'thin',color:{rgb:'EEEEEE'}}}};}
    }
  }
}

function applyXlsxStyles(ws, aoa){
  // Título en fila 0
  const titleCell=ws['A1'];
  if(titleCell){titleCell.s={font:{bold:true,sz:13,color:{rgb:'FFFFFF'}},fill:{fgColor:{rgb:'1A3A2A'}},alignment:{horizontal:'center',vertical:'center'}};}
  // Subtítulo fila 1
  const subCell=ws['A2'];
  if(subCell){subCell.s={font:{italic:true,sz:10,color:{rgb:'888888'}},alignment:{horizontal:'center'}};}
  // Cabecera (fila 3, índice 3 en aoa)
  const headerRow=aoa.findIndex(r=>Array.isArray(r)&&r.length>4&&typeof r[0]==='string'&&r[0].length>0&&!r[0].startsWith('REPORTE')&&!r[0].startsWith('Generado'));
  if(headerRow>=0){
    const cols=aoa[headerRow].length;
    for(let c=0;c<cols;c++){
      const addr=XLSX.utils.encode_cell({r:headerRow,c});
      if(ws[addr]){ws[addr].s={font:{bold:true,color:{rgb:'FFFFFF'},sz:11},fill:{fgColor:{rgb:'00A572'}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{bottom:{style:'medium',color:{rgb:'007A50'}}}};}
    }
  }
  // Filas de datos — alternar fondo
  for(let row=headerRow+1;row<aoa.length;row++){
    const cols=aoa[row].length;
    const isEven=(row-headerRow)%2===0;
    for(let c=0;c<cols;c++){
      const addr=XLSX.utils.encode_cell({r:row,c});
      if(ws[addr]){ws[addr].s={fill:{fgColor:{rgb:isEven?'F4FBF7':'FFFFFF'}},alignment:{horizontal:c===0?'left':'center',vertical:'center'},border:{bottom:{style:'thin',color:{rgb:'DDDDDD'}},right:{style:'thin',color:{rgb:'EEEEEE'}}}};}
    }
  }
}

// ══════════════════════════════════════════
//  NOTAS / JUSTIFICACIONES
// ══════════════════════════════════════════
let _notas = {}; // { resultadoId: "texto" }
let _notaCurrentId = null;

function openNotaPopup(id, event){
  _notaCurrentId = id;
  const popup = document.getElementById('nota-popup');
  const txt = document.getElementById('nota-popup-text');
  const lbl = document.getElementById('nota-popup-label');
  const r = S.resultados.find(x=>x.id===id);
  lbl.textContent = r ? `${r.empNom} · ${fmtFecha(r.fecha)}` : 'Nota';
  txt.value = _notas[id] || '';
  // Position near the button
  const btn = event.currentTarget;
  const rect = btn.getBoundingClientRect();
  popup.style.display = 'block';
  popup.style.top = Math.min(rect.bottom + 6, window.innerHeight - 180) + 'px';
  popup.style.left = Math.min(rect.left, window.innerWidth - 300) + 'px';
  txt.focus();
  event.stopPropagation();
}
function closeNotaPopup(){
  document.getElementById('nota-popup').style.display = 'none';
  _notaCurrentId = null;
}
function saveNota(){
  if(!_notaCurrentId) return;
  const txt = document.getElementById('nota-popup-text').value.trim();
  if(txt) _notas[_notaCurrentId] = txt;
  else delete _notas[_notaCurrentId];
  closeNotaPopup();
  renderRes(false); // refresh sin resetear página
  toast('Nota guardada','success');
}
// Cerrar popup al click fuera
document.addEventListener('click', e=>{
  const popup = document.getElementById('nota-popup');
  if(popup && popup.style.display !== 'none' && !popup.contains(e.target)) closeNotaPopup();
});

// ══════════════════════════════════════════
//  AUSENCIAS
// ══════════════════════════════════════════
let _ausPage=1, _ausPageSize=15, _ausFiltered=[];

function fillAusenciasFilters(){
  const aSede=document.getElementById('af-sede'); if(!aSede)return;
  aSede.innerHTML='<option value="">Todas las sedes</option>';
  S.sedes.forEach(s=>{const o=document.createElement('option');o.value=s.id;o.textContent=s.nombre;aSede.appendChild(o);});
  if(sedeActiva) aSede.value=sedeActiva;
  const aMes=document.getElementById('af-mes');
  aMes.innerHTML='<option value="">Todos los meses</option>';
  const sf=aSede.value||'';
  const meses=[...new Set(S.resultados.filter(r=>!sf||r.sedeId===sf).map(r=>r.fecha.slice(0,7)))].sort();
  meses.forEach(m=>{const o=document.createElement('option');o.value=m;o.textContent=m;aMes.appendChild(o);});
}

// Genera todos los días laborables esperados para todos los empleados
// dentro del rango de fechas que cubre S.resultados
function calcAusencias(){
  if(!S.resultados.length || !S.empleados.length) return [];
  // Determinar rango de fechas de los resultados
  const fechas = S.resultados.map(r=>r.fecha).sort();
  const minFecha = fechas[0], maxFecha = fechas[fechas.length-1];
  // Set de días registrados por empId_fecha
  const registrados = new Set(S.resultados.map(r=>r.empId+'_'+r.fecha));
  const ausencias = [];
  S.empleados.forEach(emp=>{
    const sede = S.sedes.find(s=>s.id===emp.sedeId);
    if(!sede) return;
    // Iterar cada día entre minFecha y maxFecha
    let cur = new Date(minFecha+'T00:00');
    const fin = new Date(maxFecha+'T00:00');
    while(cur <= fin){
      const fechaStr = dkey(cur);
      const dow = cur.getDay();
      // Resolver horario vigente (igual que en procesar)
      let diasSrc = sede.dias||{};
      if(emp.diasPersonal) diasSrc = emp.diasPersonal;
      if(emp.diasPersonalB && emp.hbDesde){
        const dentroB = fechaStr>=emp.hbDesde && (!emp.hbHasta||fechaStr<=emp.hbHasta);
        if(dentroB) diasSrc = emp.diasPersonalB;
      }
      const dc = diasSrc[dow];
      const esLab = !!(dc?.activo);
      // Si es laborable y NO hay registro → ausencia
      const esDescanso=(emp.descansos||[]).includes(fechaStr);
      if(esLab && !registrados.has(emp.id+'_'+fechaStr) && !esFestivo(fechaStr) && !esDescanso){
        const horas = dc ? ((t2m(dc.salida)-t2m(dc.entrada)-parseInt(dc.almuerzo||0))/60).toFixed(1) : '—';
        ausencias.push({
          empId: emp.id, empNom: emp.nombre,
          sedeId: emp.sedeId, sedeNom: sede.nombre,
          fecha: fechaStr, dia: DIAS_FULL[dow], diaCort: DIAS[dow],
          horario: dc ? `${dc.entrada}–${dc.salida}` : '—',
          horas,
        });
      }
      cur.setDate(cur.getDate()+1);
    }
  });
  return ausencias.sort((a,b)=>a.fecha.localeCompare(b.fecha)||a.empNom.localeCompare(b.empNom));
}

function renderAusencias(resetPage=true){
  const q=(document.getElementById('af-q')?.value||'').toLowerCase();
  const sf=document.getElementById('af-sede')?.value||sedeActiva||'';
  const mf=document.getElementById('af-mes')?.value||'';
  const desde=document.getElementById('af-desde')?.value||'';
  const hasta=document.getElementById('af-hasta')?.value||'';
  if(resetPage) _ausPage=1;

  const todas = calcAusencias();
  _ausFiltered = todas.filter(a=>{
    if(sf && a.sedeId!==sf) return false;
    if(mf && !a.fecha.startsWith(mf)) return false;
    if(desde && a.fecha<desde) return false;
    if(hasta && a.fecha>hasta) return false;
    if(q && !a.empNom.toLowerCase().includes(q)) return false;
    return true;
  });

  // Stats
  const total=_ausFiltered.length;
  const empsCon=[...new Set(_ausFiltered.map(a=>a.empId))].length;
  const horasPerdidas=_ausFiltered.reduce((s,a)=>s+parseFloat(a.horas||0),0);
  document.getElementById('aus-stats-area').innerHTML=`<div class="stats-grid" style="margin-bottom:18px;">
    <div class="stat-card red"><div class="stat-label">Días ausentes</div><div class="stat-val">${total}</div><div class="stat-sub">sin registro</div></div>
    <div class="stat-card amber"><div class="stat-label">Empleados afectados</div><div class="stat-val">${empsCon}</div><div class="stat-sub">con al menos 1 ausencia</div></div>
    <div class="stat-card blue"><div class="stat-label">Horas no trabajadas</div><div class="stat-val">${horasPerdidas.toFixed(1)}h</div><div class="stat-sub">estimado contrato</div></div>
  </div>`;

  const tb=document.getElementById('aus-tb');
  if(!_ausFiltered.length){
    tb.innerHTML='<tr><td colspan="7"><div class="empty-state"><div class="eico">✅</div><p>Sin ausencias detectadas para estos filtros.</p></div></td></tr>';
    document.getElementById('aus-pagination').innerHTML='';
    document.getElementById('aus-resumen').innerHTML='';
    return;
  }
  const start=(_ausPage-1)*_ausPageSize;
  const page=_ausFiltered.slice(start,start+_ausPageSize);
  tb.innerHTML=page.map(a=>{
    const nota=_notas['aus_'+a.empId+'_'+a.fecha]||'';
    return`<tr>
      <td class="name">${a.empNom}</td>
      <td style="color:var(--muted2);font-size:11px;">${a.sedeNom}</td>
      <td style="font-size:11px;">${fmtFecha(a.fecha)}</td>
      <td style="color:var(--muted);">${a.diaCort}</td>
      <td style="font-family:'DM Mono',monospace;color:var(--muted2);">${a.horario}</td>
      <td style="font-family:'DM Mono',monospace;">${a.horas}h</td>
      <td><button class="nota-btn${nota?' has-nota':''}" onclick="openNotaAus('${a.empId}','${a.fecha}',event)" title="${nota||'Agregar justificación'}">${nota?'📝 '+nota.slice(0,18)+(nota.length>18?'…':''):'＋ Justificar'}</button></td>
    </tr>`;
  }).join('');
  renderPagination('aus-pagination',_ausFiltered.length,_ausPage,_ausPageSize,
    p=>{_ausPage=p;renderAusencias(false);},
    sz=>{_ausPageSize=sz;_ausPage=1;renderAusencias(false);}
  );

  // Resumen por empleado
  const byEmp={};
  _ausFiltered.forEach(a=>{
    if(!byEmp[a.empId])byEmp[a.empId]={nom:a.empNom,sede:a.sedeNom,dias:0,horas:0};
    byEmp[a.empId].dias++;
    byEmp[a.empId].horas+=parseFloat(a.horas||0);
  });
  document.getElementById('aus-resumen').innerHTML=`<div class="table-wrap"><table>
    <thead><tr><th>Empleado</th><th>Sede</th><th>Días ausentes</th><th>Horas no trabajadas</th></tr></thead>
    <tbody>${Object.values(byEmp).sort((a,b)=>b.dias-a.dias).map(e=>`<tr>
      <td class="name">${e.nom}</td>
      <td style="color:var(--muted2);font-size:11px;">${e.sede}</td>
      <td style="color:var(--danger);font-weight:700;">${e.dias}</td>
      <td style="font-family:'DM Mono',monospace;">${e.horas.toFixed(1)}h</td>
    </tr>`).join('')}</tbody>
  </table></div>`;
}

function openNotaAus(empId, fecha, event){
  const key = 'aus_'+empId+'_'+fecha;
  _notaCurrentId = key;
  const popup = document.getElementById('nota-popup');
  const txt = document.getElementById('nota-popup-text');
  const lbl = document.getElementById('nota-popup-label');
  const emp = S.empleados.find(e=>e.id===empId);
  lbl.textContent = `Justificación · ${emp?.nombre||empId} · ${fmtFecha(fecha)}`;
  txt.value = _notas[key] || '';
  const btn = event.currentTarget;
  const rect = btn.getBoundingClientRect();
  popup.style.display = 'block';
  popup.style.top = Math.min(rect.bottom+6, window.innerHeight-180)+'px';
  popup.style.left = Math.min(rect.left, window.innerWidth-300)+'px';
  txt.focus();
  event.stopPropagation();
}

function exportAusenciasXLSX(){
  if(!_ausFiltered.length){toast('Sin ausencias para exportar','error');return;}
  const wb=XLSX.utils.book_new();
  const aoa=[];
  aoa.push(['REPORTE DE AUSENCIAS']);
  aoa.push([`Generado: ${new Date().toLocaleDateString('es-CO')}`]);
  aoa.push([]);
  aoa.push(['Empleado','Sede','Fecha','Día','Horario esperado','Horas contrato','Justificación']);
  _ausFiltered.forEach(a=>{
    const nota=_notas['aus_'+a.empId+'_'+a.fecha]||'';
    aoa.push([a.empNom,a.sedeNom,fmtFecha(a.fecha),a.dia,a.horario,a.horas+'h',nota]);
  });
  aoa.push([]);
  // Resumen
  aoa.push(['RESUMEN POR EMPLEADO']);
  aoa.push(['Empleado','Sede','Días ausentes','Horas no trabajadas']);
  const byEmp={};
  _ausFiltered.forEach(a=>{
    if(!byEmp[a.empId])byEmp[a.empId]={nom:a.empNom,sede:a.sedeNom,dias:0,horas:0};
    byEmp[a.empId].dias++;byEmp[a.empId].horas+=parseFloat(a.horas||0);
  });
  Object.values(byEmp).sort((a,b)=>b.dias-a.dias).forEach(e=>{
    aoa.push([e.nom,e.sede,e.dias,e.horas.toFixed(1)+'h']);
  });
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols']=[{wch:22},{wch:16},{wch:12},{wch:10},{wch:16},{wch:14},{wch:30}];
  ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:6}},{s:{r:1,c:0},e:{r:1,c:6}}];
  applyXlsxStyles(ws,aoa);
  XLSX.utils.book_append_sheet(wb,ws,'Ausencias');
  XLSX.writeFile(wb, `ausencias_${new Date().toISOString().slice(0,10)}.xlsx`, {cellStyles:true});
  toast('Reporte de ausencias exportado','success');
}

// ══════════════════════════════════════════
//  EXPORTAR PDF (llegadas + ausencias)
// ══════════════════════════════════════════
function buildPDF(title, subtitle, headers, rows, filename){
  const {jsPDF} = window.jspdf;
  const doc = new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});
  const accent = [0,165,114]; // #00A572
  const dark   = [26,58,42];
  const W = doc.internal.pageSize.getWidth();

  // Header bar
  doc.setFillColor(...dark);
  doc.rect(0,0,W,20,'F');
  doc.setTextColor(255,255,255);
  doc.setFontSize(13); doc.setFont('helvetica','bold');
  doc.text(title, 14, 13);
  doc.setFontSize(9); doc.setFont('helvetica','normal');
  doc.text(subtitle, W-14, 13, {align:'right'});

  // Table
  doc.autoTable({
    startY: 24,
    head: [headers],
    body: rows,
    styles:{fontSize:8, cellPadding:2.5, textColor:[30,30,30]},
    headStyles:{fillColor:accent, textColor:[255,255,255], fontStyle:'bold', halign:'center'},
    alternateRowStyles:{fillColor:[244,251,247]},
    columnStyles:{0:{fontStyle:'bold'}},
    margin:{left:14,right:14},
    didDrawPage:(data)=>{
      // Footer
      doc.setFontSize(7); doc.setTextColor(150,150,150);
      doc.text(`Página ${data.pageNumber}  ·  Control HH.EE Hikvision  ·  ${new Date().toLocaleDateString('es-CO')}`,
        W/2, doc.internal.pageSize.getHeight()-6, {align:'center'});
    }
  });
  doc.save(filename);
}

function exportLlegadasPDF(){
  if(!_llegFiltered.length){toast('Sin datos de llegadas','error');return;}
  const headers=['Empleado','Sede','Fecha','Día','Entrada','Esp.','Dif. entrada','Salida','Esp.','Dif. salida','Horario','Estado'];
  const rows=_llegFiltered.map(r=>{
    const entEsp=r.entradaEsperada||'--:--';
    const salEsp=r.salidaEsperada||'--:--';
    const difEnt=r.tardanza===0?'Puntual':(r.tardanza>0?'+'+m2hm(r.tardanza)+' tarde':'-'+m2hm(Math.abs(r.tardanza))+' antes');
    const difSal=r.sinSalida?'—':((t2m(r.salida)-t2m(salEsp))===0?'Puntual':((t2m(r.salida)-t2m(salEsp))>0?'+'+m2hm(t2m(r.salida)-t2m(salEsp))+' extra':'-'+m2hm(Math.abs(t2m(r.salida)-t2m(salEsp)))+' antes'));
    return[r.empNom,r.sedeNom,fmtFecha(r.fecha),r.diaCort,r.entrada,entEsp,difEnt,r.sinSalida?'—':r.salida,salEsp,difSal,r.horarioUsado||'Sede',r.estado];
  });
  buildPDF('Reporte de Llegadas y Puntualidad',`Generado: ${new Date().toLocaleDateString('es-CO')} · ${_llegFiltered.length} registros`,headers,rows,`llegadas_${new Date().toISOString().slice(0,10)}.pdf`);
  toast('PDF generado','success');
}

function exportAusenciasPDF(){
  if(!_ausFiltered.length){toast('Sin ausencias para exportar','error');return;}
  const headers=['Empleado','Sede','Fecha','Día','Horario esperado','Horas contrato','Justificación'];
  const rows=_ausFiltered.map(a=>[a.empNom,a.sedeNom,fmtFecha(a.fecha),a.diaCort,a.horario,a.horas+'h',_notas['aus_'+a.empId+'_'+a.fecha]||'']);
  buildPDF('Reporte de Ausencias',`Generado: ${new Date().toLocaleDateString('es-CO')} · ${_ausFiltered.length} ausencias`,headers,rows,`ausencias_${new Date().toISOString().slice(0,10)}.pdf`);
  toast('PDF generado','success');
}

// ══════════════════════════════════════════
//  REPORTE MICHEL — tabla por empleado
// ══════════════════════════════════════════
function fillMichelFilters(){
  const mSede=document.getElementById('mf-sede'); if(!mSede)return;
  mSede.innerHTML='<option value="">Todas las sedes</option>';
  S.sedes.forEach(s=>{const o=document.createElement('option');o.value=s.id;o.textContent=s.nombre;mSede.appendChild(o);});
  if(sedeActiva) mSede.value=sedeActiva;
  const mMes=document.getElementById('mf-mes');
  mMes.innerHTML='<option value="">Todos los meses</option>';
  const sf=mSede.value||'';
  const meses=[...new Set(S.resultados.filter(r=>!sf||r.sedeId===sf).map(r=>r.fecha.slice(0,7)))].sort();
  meses.forEach(m=>{const o=document.createElement('option');o.value=m;o.textContent=m;mMes.appendChild(o);});
}

function getMichelData(){
  const sf=document.getElementById('mf-sede')?.value||sedeActiva||'';
  const mf=document.getElementById('mf-mes')?.value||'';
  const desde=document.getElementById('mf-desde')?.value||'';
  const hasta=document.getElementById('mf-hasta')?.value||'';
  return S.resultados.filter(r=>{
    if(!r.esLab) return false;
    if(sf&&r.sedeId!==sf) return false;
    if(mf&&!r.fecha.startsWith(mf)) return false;
    if(desde&&r.fecha<desde) return false;
    if(hasta&&r.fecha>hasta) return false;
    return true;
  }).sort((a,b)=>a.fecha.localeCompare(b.fecha));
}

function renderMichel(){
  const area=document.getElementById('michel-area');
  if(!S.resultados.length){area.innerHTML='<div class="empty-state"><div class="eico">📋</div><p>Procesa un CSV primero.</p></div>';return;}
  const data=getMichelData();
  if(!data.length){area.innerHTML='<div class="empty-state"><div class="eico">📋</div><p>Sin registros para estos filtros.</p></div>';return;}

  // Agrupar por empleado
  const byEmp={};
  data.forEach(r=>{
    if(!byEmp[r.empId]) byEmp[r.empId]={nom:r.empNom,sede:r.sedeNom,rows:[]};
    byEmp[r.empId].rows.push(r);
  });

  area.innerHTML=Object.values(byEmp).map(emp=>{
    const filas=emp.rows.map(r=>{
      const tarde=r.tardanza>0;
      const estilo=tarde?'background:#FFF0F0;color:#c0392b;font-weight:600;':'';
      return`<tr style="${estilo}">
        <td style="padding:8px 12px;font-family:'DM Mono',monospace;font-size:12px;border-bottom:1px solid #eee;${tarde?'color:#c0392b;':''}">${fmtFecha(r.fecha)}</td>
        <td style="padding:8px 12px;font-family:'DM Mono',monospace;font-size:12px;border-bottom:1px solid #eee;${tarde?'color:#c0392b;':''}">${r.dia}</td>
        <td style="padding:8px 12px;font-family:'DM Mono',monospace;font-size:12px;border-bottom:1px solid #eee;font-weight:${tarde?'700':'400'};${tarde?'color:#c0392b;':''}">${r.entrada}${tarde?` <span style="font-size:10px;">(+${m2hm(r.tardanza)})</span>`:''}</td>
      </tr>`;
    }).join('');
    return`<div class="card" style="margin-bottom:18px;max-width:480px;">
      <div style="background:#1A3A2A;color:#fff;padding:12px 16px;border-radius:8px 8px 0 0;margin:-22px -22px 14px;">
        <div style="font-size:10px;font-family:'DM Mono',monospace;letter-spacing:2px;opacity:.7;margin-bottom:2px;">REPORTE DE HORAS DE LLEGADA</div>
        <div style="font-size:14px;font-weight:700;">Colaborador: ${emp.nom}</div>
        <div style="font-size:11px;font-family:'DM Mono',monospace;opacity:.6;margin-top:2px;">${emp.sede}</div>
      </div>
      <table style="width:100%;border-collapse:collapse;">
        <thead><tr style="background:#00A572;">
          <th style="padding:8px 12px;text-align:left;color:#fff;font-size:11px;font-family:'DM Mono',monospace;letter-spacing:1px;">Fecha</th>
          <th style="padding:8px 12px;text-align:left;color:#fff;font-size:11px;font-family:'DM Mono',monospace;letter-spacing:1px;">Día de la Semana</th>
          <th style="padding:8px 12px;text-align:left;color:#fff;font-size:11px;font-family:'DM Mono',monospace;letter-spacing:1px;">Hora de Llegada</th>
        </tr></thead>
        <tbody>${filas}</tbody>
        <tfoot><tr>
          <td colspan="3" style="padding:10px 12px;font-family:'DM Mono',monospace;font-size:12px;font-weight:700;border-top:2px solid #1A3A2A;color:#1A3A2A;">
            Total de días registrados: ${emp.rows.length}
          </td>
        </tr></tfoot>
      </table>
    </div>`;
  }).join('');
}

function exportMichelXLSX(){
  const data=getMichelData();
  if(!data.length){toast('Sin datos','error');return;}
  const wb=XLSX.utils.book_new();
  const byEmp={};
  data.forEach(r=>{
    if(!byEmp[r.empId]) byEmp[r.empId]={nom:r.empNom,sede:r.sedeNom,rows:[]};
    byEmp[r.empId].rows.push(r);
  });
  Object.values(byEmp).forEach(emp=>{
    const aoa=[];
    aoa.push(['Reporte de Horas de Llegada']);
    aoa.push([`Colaborador: ${emp.nom}`]);
    aoa.push([]);
    aoa.push(['Fecha','Día de la Semana','Hora de Llegada']);
    emp.rows.forEach(r=>aoa.push([fmtFecha(r.fecha),r.dia,r.entrada]));
    aoa.push([]);
    aoa.push([`Total de días registrados: ${emp.rows.length}`]);
    const ws=XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols']=[{wch:14},{wch:18},{wch:16}];
    ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:2}},{s:{r:1,c:0},e:{r:1,c:2}}];
    // Estilos
    if(ws['A1']) ws['A1'].s={font:{bold:true,sz:13,color:{rgb:'FFFFFF'}},fill:{fgColor:{rgb:'1A3A2A'}},alignment:{horizontal:'center'}};
    if(ws['A2']) ws['A2'].s={font:{bold:true,sz:11,color:{rgb:'1A3A2A'}},alignment:{horizontal:'center'}};
    ['A4','B4','C4'].forEach(addr=>{if(ws[addr])ws[addr].s={font:{bold:true,color:{rgb:'FFFFFF'},sz:11},fill:{fgColor:{rgb:'00A572'}},alignment:{horizontal:'center'}};});
    // Filas rojas para tardanzas
    emp.rows.forEach((r,i)=>{
      if(r.tardanza>0){
        ['A','B','C'].forEach(col=>{
          const addr=col+(i+5);
          if(ws[addr]) ws[addr].s={fill:{fgColor:{rgb:'FFF0F0'}},font:{color:{rgb:'C0392B'},bold:true},alignment:{horizontal:'center'}};
        });
      } else {
        ['A','B','C'].forEach(col=>{
          const addr=col+(i+5);
          if(ws[addr]) ws[addr].s={fill:{fgColor:{rgb:i%2===0?'F4FBF7':'FFFFFF'}},alignment:{horizontal:'center'}};
        });
      }
    });
    const shName=emp.nom.substring(0,31).replace(/[\\/\?\*\[\]:]/g,'_');
    XLSX.utils.book_append_sheet(wb,ws,shName);
  });
  XLSX.writeFile(wb, `llegadas_michel_${new Date().toISOString().slice(0,10)}.xlsx`, {cellStyles:true});
  toast('Excel Michel exportado','success');
}

function exportMichelPDF(){
  const data=getMichelData();
  if(!data.length){toast('Sin datos','error');return;}
  const {jsPDF}=window.jspdf;
  const doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});
  const byEmp={};
  data.forEach(r=>{
    if(!byEmp[r.empId]) byEmp[r.empId]={nom:r.empNom,sede:r.sedeNom,rows:[]};
    byEmp[r.empId].rows.push(r);
  });
  const emps=Object.values(byEmp);
  emps.forEach((emp,ei)=>{
    if(ei>0) doc.addPage();
    const W=doc.internal.pageSize.getWidth();
    // Header
    doc.setFillColor(26,58,42);
    doc.rect(0,0,W,22,'F');
    doc.setTextColor(255,255,255);
    doc.setFontSize(8); doc.setFont('helvetica','normal');
    doc.text('REPORTE DE HORAS DE LLEGADA',14,10);
    doc.setFontSize(13); doc.setFont('helvetica','bold');
    doc.text(`Colaborador: ${emp.nom}`,14,18);
    doc.setFontSize(8); doc.setFont('helvetica','normal');
    doc.setTextColor(200,200,200);
    doc.text(emp.sede, W-14, 18, {align:'right'});
    // Table
    const tableRows=emp.rows.map(r=>[fmtFecha(r.fecha),r.dia,r.entrada+(r.tardanza>0?` (+${m2hm(r.tardanza)})` : '')]);
    doc.autoTable({
      startY:27,
      head:[['Fecha','Día de la Semana','Hora de Llegada']],
      body:tableRows,
      styles:{fontSize:10,cellPadding:3,textColor:[30,30,30]},
      headStyles:{fillColor:[0,165,114],textColor:[255,255,255],fontStyle:'bold',halign:'center'},
      alternateRowStyles:{fillColor:[244,251,247]},
      didParseCell:(data)=>{
        if(data.section==='body'){
          const r=emp.rows[data.row.index];
          if(r&&r.tardanza>0){
            data.cell.styles.textColor=[192,57,43];
            data.cell.styles.fillColor=[255,240,240];
            data.cell.styles.fontStyle='bold';
          }
        }
      },
      margin:{left:14,right:14},
    });
    // Footer
    const finalY=doc.lastAutoTable.finalY+8;
    doc.setFont('helvetica','bold'); doc.setFontSize(11); doc.setTextColor(26,58,42);
    doc.text(`Total de días registrados: ${emp.rows.length}`, 14, finalY);
    // Page footer
    doc.setFontSize(7); doc.setTextColor(150,150,150); doc.setFont('helvetica','normal');
    doc.text(`Control HH.EE · ${new Date().toLocaleDateString('es-CO')}`, W/2, doc.internal.pageSize.getHeight()-6, {align:'center'});
  });
  doc.save(`llegadas_michel_${new Date().toISOString().slice(0,10)}.pdf`);
  toast('PDF Michel generado','success');
}

// ══════════════════════════════════════════
//  EXPORT HH.EE PDF
// ══════════════════════════════════════════
function exportExtrassPDF(){
  const data=getXData();
  if(!data.length){toast('Sin datos','error');return;}
  const incT=document.getElementById('xc-t').checked;
  const incD=document.getElementById('xc-d').checked;
  const incTrab=document.getElementById('xc-trab').checked;
  const formato=document.getElementById('xf-formato').value;

  const {jsPDF}=window.jspdf;

  if(formato==='por_empleado'){
    const byEmp={};
    data.forEach(r=>{
      if(!byEmp[r.empId]) byEmp[r.empId]={nom:r.empNom,sede:r.sedeNom,rows:[]};
      byEmp[r.empId].rows.push(r);
    });
    const doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});
    const emps=Object.values(byEmp);
    emps.forEach((emp,ei)=>{
      if(ei>0) doc.addPage();
      const W=doc.internal.pageSize.getWidth();
      doc.setFillColor(26,58,42); doc.rect(0,0,W,22,'F');
      doc.setTextColor(255,255,255);
      doc.setFontSize(8); doc.setFont('helvetica','normal');
      doc.text('REPORTE DE HORAS EXTRAS',14,10);
      doc.setFontSize(13); doc.setFont('helvetica','bold');
      doc.text(`Colaborador: ${emp.nom}`,14,18);
      doc.setFontSize(8); doc.setFont('helvetica','normal');
      doc.setTextColor(200,200,200);
      doc.text(emp.sede, W-14, 18, {align:'right'});
      const heads=['Fecha','Día','Entrada','Salida'];
      if(incTrab) heads.push('H. Trabajadas');
      if(incT) heads.push('Tardanza');
      if(incD){heads.push('HE Diurnas');heads.push('HE Nocturnas');}
      heads.push('Total HE');
      emp.rows.sort((a,b)=>a.fecha.localeCompare(b.fecha));
      const rows=emp.rows.map(r=>{
        const ss=r.sinSalida;
        const row=[fmtFecha(r.fecha),r.diaCort,r.entrada,ss?'—':r.salida];
        if(incTrab) row.push(ss?'—':m2hm(r.trabajadasNeto));
        if(incT) row.push(r.tardanza>0?m2hm(r.tardanza):'00h 00m');
        if(incD){row.push(m2hm(r.heDiur));row.push(m2hm(r.heNoct));}
        row.push(ss?'—':m2hm(r.extrasMin));
        return row;
      });
      doc.autoTable({
        startY:27,head:[heads],body:rows,
        styles:{fontSize:9,cellPadding:2.5,textColor:[30,30,30]},
        headStyles:{fillColor:[0,165,114],textColor:[255,255,255],fontStyle:'bold',halign:'center'},
        alternateRowStyles:{fillColor:[244,251,247]},
        columnStyles:{0:{fontStyle:'bold'}},
        margin:{left:14,right:14},
      });
      const totalExtras=emp.rows.filter(r=>!r.sinSalida).reduce((s,r)=>s+r.extrasMin,0);
      const fy=doc.lastAutoTable.finalY+6;
      doc.setFont('helvetica','bold');doc.setFontSize(10);doc.setTextColor(26,58,42);
      doc.text(`Total días: ${emp.rows.length}   ·   Total HH.EE: ${m2hm(totalExtras)}`,14,fy);
      doc.setFontSize(7);doc.setTextColor(150,150,150);doc.setFont('helvetica','normal');
      doc.text(`Control HH.EE · ${new Date().toLocaleDateString('es-CO')}`,W/2,doc.internal.pageSize.getHeight()-6,{align:'center'});
    });
    doc.save(`horas_extras_por_empleado_${new Date().toISOString().slice(0,10)}.pdf`);
  } else {
    // Consolidado
    const heads=['Empleado','Sede','Fecha','Día','Entrada','Salida','Total HE','Estado'];
    const rows=data.map(r=>[r.empNom,r.sedeNom,fmtFecha(r.fecha),r.diaCort,r.entrada,r.sinSalida?'—':r.salida,r.sinSalida?'—':m2hm(r.extrasMin),r.estado]);
    buildPDF('Reporte Consolidado de Horas Extras',`${new Date().toLocaleDateString('es-CO')} · ${data.length} registros`,heads,rows,`horas_extras_consolidado_${new Date().toISOString().slice(0,10)}.pdf`);
  }
  toast('PDF exportado','success');
}

// ══════════════════════════════════════════
//  REPORTE DE HORAS DE LLEGADA (estilo Michel)
// ══════════════════════════════════════════
function exportReporteMichel(fmt){
  if(!_llegFiltered.length){toast('Sin datos de llegadas','error');return;}
  const byEmp={};
  _llegFiltered.forEach(r=>{
    if(!byEmp[r.empId]) byEmp[r.empId]={nom:r.empNom,sede:r.sedeNom,rows:[]};
    byEmp[r.empId].rows.push(r);
  });
  const emps=Object.values(byEmp);

  if(fmt==='xlsx'){
    const wb=XLSX.utils.book_new();
    emps.forEach(emp=>{
      emp.rows.sort((a,b)=>a.fecha.localeCompare(b.fecha));
      const aoa=[];
      aoa.push(['Reporte de Horas de Llegada']);
      aoa.push(['Colaborador: '+emp.nom]);
      aoa.push([]);
      aoa.push(['Fecha','Dia de la Semana','Hora de Llegada']);
      emp.rows.forEach(r=>aoa.push([fmtFecha(r.fecha),r.dia,r.entrada]));
      aoa.push([]);
      aoa.push(['Total de dias registrados: '+emp.rows.length]);
      const ws=XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols']=[{wch:14},{wch:18},{wch:16}];
      ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:2}},{s:{r:1,c:0},e:{r:1,c:2}}];
      const t=ws['A1'];if(t)t.s={font:{bold:true,sz:14,color:{rgb:'FFFFFF'}},fill:{fgColor:{rgb:'1A3A2A'}},alignment:{horizontal:'center',vertical:'center'}};
      const c=ws['A2'];if(c)c.s={font:{bold:true,sz:12,color:{rgb:'FFFFFF'}},fill:{fgColor:{rgb:'1A3A2A'}},alignment:{horizontal:'center',vertical:'center'}};
      ['A4','B4','C4'].forEach(addr=>{if(ws[addr])ws[addr].s={font:{bold:true,color:{rgb:'FFFFFF'},sz:11},fill:{fgColor:{rgb:'00A572'}},alignment:{horizontal:'center',vertical:'center'}};});
      emp.rows.forEach((r,i)=>{
        const late=r.tardanza>0;
        const bg=late?'FFE5E5':(i%2===0?'F4FBF7':'FFFFFF');
        const fc=late?'CC0000':'222222';
        const bord={bottom:{style:'thin',color:{rgb:'DDDDDD'}},right:{style:'thin',color:{rgb:'EEEEEE'}}};
        const rowIdx=4+i+1;
        ['A','B','C'].forEach((col,ci)=>{
          const a=col+rowIdx;
          if(ws[a])ws[a].s={fill:{fgColor:{rgb:bg}},font:{color:{rgb:fc},bold:late},alignment:{horizontal:ci===0?'left':'center',vertical:'center'},border:bord};
        });
      });
      const totAddr=XLSX.utils.encode_cell({r:4+emp.rows.length+1,c:0});
      if(ws[totAddr])ws[totAddr].s={font:{bold:true,color:{rgb:'1A3A2A'},sz:11},fill:{fgColor:{rgb:'E8F5EF'}},alignment:{horizontal:'left'}};
      const shName=emp.nom.substring(0,31).replace(/[\/\?\*\[\]:]/g,'_');
      XLSX.utils.book_append_sheet(wb,ws,shName);
    });
    XLSX.writeFile(wb, 'reporte_llegadas_'+new Date().toISOString().slice(0,10)+'.xlsx', {cellStyles:true});
    toast('Excel generado: '+emps.length+' empleado(s)','success');

  } else {
    const {jsPDF}=window.jspdf;
    const doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});
    const W=doc.internal.pageSize.getWidth();
    const H=doc.internal.pageSize.getHeight();
    // Encabezado global
    doc.setFillColor(26,58,42); doc.rect(0,0,W,18,'F');
    doc.setTextColor(255,255,255);
    doc.setFontSize(11); doc.setFont('helvetica','bold');
    doc.text('REPORTE DE HORAS DE LLEGADA',14,12);
    doc.setFontSize(8); doc.setFont('helvetica','normal');
    doc.setTextColor(180,220,200);
    doc.text(new Date().toLocaleDateString('es-CO'), W-14, 12, {align:'right'});
    let curY=24;
    emps.forEach((emp,ei)=>{
      emp.rows.sort((a,b)=>a.fecha.localeCompare(b.fecha));
      const spaceNeeded=10+emp.rows.length*6.5+14;
      if(ei>0 && curY+spaceNeeded>H-12){ doc.addPage(); curY=14; }
      // Mini header por empleado
      doc.setFillColor(0,165,114); doc.roundedRect(14,curY,W-28,8,1.5,1.5,'F');
      doc.setTextColor(255,255,255); doc.setFontSize(9); doc.setFont('helvetica','bold');
      doc.text('Colaborador: '+emp.nom, 18, curY+5.5);
      doc.setFontSize(7.5); doc.setFont('helvetica','normal');
      doc.setTextColor(220,255,240);
      doc.text(emp.sede, W-16, curY+5.5, {align:'right'});
      curY+=10;
      doc.autoTable({
        startY:curY,
        head:[['Fecha','Dia de la Semana','Hora de Llegada']],
        body:emp.rows.map(r=>[fmtFecha(r.fecha),r.dia,r.entrada]),
        styles:{fontSize:8,cellPadding:1.8,textColor:[30,30,30]},
        headStyles:{fillColor:[26,58,42],textColor:[255,255,255],fontStyle:'bold',halign:'center',fontSize:8},
        alternateRowStyles:{fillColor:[244,251,247]},
        columnStyles:{0:{halign:'left',fontStyle:'bold',cellWidth:38},1:{halign:'center',cellWidth:48},2:{halign:'center',cellWidth:38}},
        margin:{left:14,right:14},
        tableWidth:'fixed',
        didParseCell:(data)=>{
          if(data.section==='body'){
            const r=emp.rows[data.row.index];
            if(r&&r.tardanza>0){
              data.cell.styles.fillColor=[255,229,229];
              data.cell.styles.textColor=[180,0,0];
              data.cell.styles.fontStyle='bold';
            }
          }
        }
      });
      curY=doc.lastAutoTable.finalY+2;
      doc.setFont('helvetica','italic'); doc.setFontSize(7.5); doc.setTextColor(26,58,42);
      doc.text('Total de dias registrados: '+emp.rows.length, 14, curY+4);
      curY+=10;
    });
    const totalPgs=doc.internal.getNumberOfPages();
    for(let p=1;p<=totalPgs;p++){
      doc.setPage(p);
      doc.setFontSize(6.5); doc.setTextColor(160,160,160); doc.setFont('helvetica','normal');
      doc.text('Control HH.EE  |  '+new Date().toLocaleDateString('es-CO')+'  |  Pag. '+p+'/'+totalPgs, W/2, H-4, {align:'center'});
    }
    doc.save('reporte_llegadas_'+new Date().toISOString().slice(0,10)+'.pdf');
    toast('PDF generado: '+emps.length+' empleado(s)','success');
  }
}