const EVALUATIONS = [
  {order:0, code:'data0', label:'Recopilación Examen 2026-II', short:'Recop.', path:'data0/resultados.xlsx', color:'#755282', text:'#FFFFFF'},
  {order:1, code:'data1', label:'Simulacro I', short:'Sim. I', path:'data1/resultados.xlsx', color:'#B9964A', text:'#FFFFFF'},
  {order:2, code:'data2', label:'Simulacro II', short:'Sim. II', path:'data2/resultados.xlsx', color:'#617F9D', text:'#FFFFFF'},
  {order:3, code:'data3', label:'Simulacro III', short:'Sim. III', path:'data3/resultados.xlsx', color:'#A56868', text:'#FFFFFF'},
  {order:4, code:'data4', label:'Simulacro IV', short:'Sim. IV', path:'data4/resultados.xlsx', color:'#688A72', text:'#FFFFFF'}
];

const AREA_CONFIG = [
  {name:'Análisis textual', score:'puntaje_analisis_textual', hits:'aciertos_analisis_textual', max:25},
  {name:'Matemáticas', score:'puntaje_matematicas', hits:'aciertos_matematicas', max:25},
  {name:'Ciencias naturales', score:'puntaje_ciencias_naturales', hits:'aciertos_ciencias_naturales', max:25},
  {name:'Ciencias sociales', score:'puntaje_ciencias_sociales', hits:'aciertos_ciencias_sociales', max:25},
  {name:'Análisis de imagen', score:'puntaje_analisis_imagen', hits:'aciertos_analisis_imagen', max:20}
];

const PREUNAL_URL = 'https://docs.google.com/forms/d/e/1FAIpQLScTS2tK3FG7hsLIgsh9g7CSE81Yql0fLYq3TBjXUt5ZbkAHag/viewform';
const PUNTAJES_UNAL_URL = 'https://proyectiapreu.github.io/historico_puntajesunal/';
const STORE = new Map();
let currentDoc = '';
let currentRecords = [];
let currentDetail = [];
let currentAreaProgress = AREA_CONFIG[0].name;

const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
const num = v => { const n = Number(String(v ?? '').replace(',','.')); return Number.isFinite(n) ? n : null; };
const normDoc = v => String(v ?? '').trim().replace(/\.0$/,'').replace(/[.\-\s]/g,'').toUpperCase();
const getName = row => String(row?.nombre ?? '').trim();
const globalValue = v => { const n=num(v); return n===null ? null : Math.trunc(n); };
const fmtGlobal = v => { const n=globalValue(v); return n===null ? '—' : n.toLocaleString('es-CO'); };
const fmtArea = v => { const n=num(v); return n===null ? '—' : n.toFixed(1); };
const scoreOf = (row, area) => num(row?.[AREA_CONFIG.find(a=>a.name===area)?.score]);
const evByOrder = order => EVALUATIONS.find(e=>e.order===order);

function showOnly(id){
  ['#menu-view','#report-view','#progress-view'].forEach(x=>$(x).classList.toggle('hidden',x!==id));
  window.scrollTo({top:430,behavior:'smooth'});
}
function sheetRows(wb,name){ const ws=wb.Sheets[name]; return ws ? XLSX.utils.sheet_to_json(ws,{defval:'',raw:true}) : []; }

async function loadWorkbook(ev){
  try{
    const r=await fetch(ev.path,{cache:'no-store'});
    if(!r.ok) throw new Error(`${r.status}`);
    const ab=await r.arrayBuffer();
    const wb=XLSX.read(ab,{type:'array'});
    STORE.set(ev.order,{
      resultados:sheetRows(wb,'resultados'),
      areas:sheetRows(wb,'areas'),
      temas:sheetRows(wb,'temas'),
      preguntas:sheetRows(wb,'preguntas'),
      metadatos:sheetRows(wb,'metadatos')
    });
  }catch(err){
    console.warn(`No se pudo cargar ${ev.path}`,err);
    STORE.set(ev.order,{resultados:[],areas:[],temas:[],preguntas:[],metadatos:[]});
  }
}

async function init(){
  const estado=$('#estado');
  if(typeof XLSX==='undefined'){
    estado.textContent='No se pudo cargar el lector de Excel. Recarga la página.';
    return;
  }
  estado.textContent='Cargando resultados…';
  await Promise.all(EVALUATIONS.map(loadWorkbook));
  estado.textContent='';
  $('#buscar').addEventListener('click',searchStudent);
  $('#documento').addEventListener('keydown',e=>{if(e.key==='Enter')searchStudent();});
}
document.addEventListener('DOMContentLoaded',init);

function resultFor(order,doc){
  const data=STORE.get(order); if(!data) return null;
  return data.resultados.find(r=>normDoc(r.documento)===doc) || null;
}
function searchStudent(){
  const doc=normDoc($('#documento').value);
  const estado=$('#estado');
  if(!doc){ estado.textContent='Ingresa tu documento de identidad.'; return; }
  const found=[];
  for(const ev of EVALUATIONS){ const row=resultFor(ev.order,doc); if(row) found.push({ev,row,data:STORE.get(ev.order)}); }
  if(!found.length){
    currentDoc=''; currentRecords=[];
    estado.textContent='No encontré resultados asociados a ese documento. Revisa el número e inténtalo de nuevo.';
    ['#menu-view','#report-view','#progress-view'].forEach(x=>$(x).classList.add('hidden'));
    return;
  }
  currentDoc=doc; currentRecords=found; estado.textContent=''; renderMenu(found);
}
function displayName(found){
  for(let i=found.length-1;i>=0;i--){ const n=getName(found[i].row); if(n) return n; }
  return '';
}

function studentOwnership(name){
  return name ? `<div class="student-ownership"><span>Estos resultados pertenecen a</span><strong>${esc(name)}</strong></div>` : '';
}

function renderMenu(found){
  const name=displayName(found);
  const available=new Set(found.map(x=>x.ev.order));
  const cards=EVALUATIONS.map((ev,i)=>{
    const on=available.has(ev.order);
    const sub=ev.order===0?'Resultado de la recopilación del examen 2026-II':'Resultado individual del simulacro';
    return `<button class="option-card evaluation-option" style="--eval-color:${ev.color};--eval-text:${ev.text}" ${on?`onclick="showEvaluation(${ev.order})"`:'disabled'}>
      <span class="option-index">${i+1}</span><strong>${esc(ev.label)}</strong><p>${sub}</p><span class="option-status">${on?'Ver resultado':'Aún no disponible'}</span>
    </button>`;
  }).join('');
  $('#menu-view').innerHTML=`
    <div class="card menu-card">
      <div class="menu-header"><div><h2>Elige qué resultado quieres consultar</h2>${studentOwnership(name)}</div></div>
      <div class="options-grid">${cards}</div>
      <div class="menu-progress-zone">
        <div><span class="menu-progress-kicker">Seguimiento del curso</span><strong>Revisa cómo han cambiado tus puntajes</strong></div>
        <button class="menu-progress-button" onclick="showProgress()">Ver mi progresión</button>
      </div>
    </div>`;
  showOnly('#menu-view');
}

function toolbar(title,subtitle,ev=null){
  return `<div class="report-toolbar" ${ev?`style="--eval-color:${ev.color}"`:''}>
    <div>${ev?`<span class="evaluation-badge" style="--eval-color:${ev.color};--eval-text:${ev.text}"><i></i>${esc(ev.short)}</span>`:''}<h2>${esc(title)}</h2>${subtitle?`<p>${subtitle}</p>`:''}</div>
    <div class="toolbar-actions">
      ${ev?`<button class="btn-small btn-progress" onclick="showProgress()">Ver mi progresión</button>`:''}
      <button class="btn-small" onclick="renderMenu(currentRecords)">Volver a los resultados</button>
    </div>
  </div>`;
}

function areaRows(row,ev){
  return `<div class="area-list">${AREA_CONFIG.map(a=>{
    const score=num(row[a.score]); const hits=num(row[a.hits]);
    const width=score===null?0:Math.max(7,Math.min(100,50+(score-10)*24));
    return `<div class="area-row">
      <div class="area-name"><strong>${a.name}</strong><span>${hits===null?'':`${hits} de ${a.max} respuestas correctas`}</span></div>
      <div class="area-bar"><span style="width:${width}%"></span></div>
      <div class="area-score"><small>Puntaje</small><strong>${fmtArea(score)}</strong></div>
    </div>`;
  }).join('')}</div>`;
}
function questionsFor(data,doc){ return (data.preguntas||[]).filter(t=>normDoc(t.documento)===doc); }
function topicsFor(data,doc){ return (data.temas||[]).filter(t=>normDoc(t.documento)===doc); }
function naturalTotal(row,data,subarea,field,fallback){
  const fromRow=num(row[field]); if(fromRow!==null) return fromRow;
  const qs=questionsFor(data,currentDoc).filter(q=>String(q.area)==='Ciencias naturales' && String(q.subarea)===subarea);
  return qs.length || fallback;
}
function naturalChips(row,data){
  const f=num(row.aciertos_fisica),q=num(row.aciertos_quimica),b=num(row.aciertos_biologia);
  if(f===null&&q===null&&b===null) return '';
  const tf=naturalTotal(row,data,'Física','total_fisica',11), tq=naturalTotal(row,data,'Química','total_quimica',6), tb=naturalTotal(row,data,'Biología','total_biologia',8);
  return `<div class="nat-chips">
    ${f!==null?`<span class="nat-chip">Física: ${f}/${tf}</span>`:''}
    ${q!==null?`<span class="nat-chip">Química: ${q}/${tq}</span>`:''}
    ${b!==null?`<span class="nat-chip">Biología: ${b}/${tb}</span>`:''}
  </div>`;
}

function reviewTopics(data){
  const topics=topicsFor(data,currentDoc).map(t=>({...t,pct:num(t.porcentaje_acierto),errors:num(t.errores),hits:num(t.aciertos),total:num(t.total_preguntas)}));
  topics.sort((a,b)=>(a.pct??101)-(b.pct??101) || (b.errors??0)-(a.errors??0));
  if(!topics.length) return '<p class="method-note">No hay información temática disponible para este resultado.</p>';
  return `<ul class="priority-list">${topics.slice(0,4).map(t=>`<li><strong>${esc(t.tema||'Tema')}</strong>${t.area?` · ${esc(t.area)}`:''}: ${t.hits===null?'—':t.hits}/${t.total===null?'—':t.total} respuestas correctas.</li>`).join('')}</ul>`;
}
function quickReading(row){
  return `El puntaje global estimado registrado para esta prueba es <strong>${fmtGlobal(row.puntaje_global)}</strong>. Revisa los puntajes de las cinco áreas y el detalle por pregunta para identificar en qué contenidos tuviste más y menos aciertos.`;
}

function showEvaluation(order){
  const item=currentRecords.find(x=>x.ev.order===order); if(!item) return;
  const {ev,row,data}=item; const name=getName(row)||displayName(currentRecords); const total=num(row.aciertos_totales); currentDetail=questionsFor(data,currentDoc);
  $('#report-view').innerHTML=`
    ${toolbar(ev.label,'',ev)}
    ${studentOwnership(name)}
    <div class="grid summary-grid">
      <article class="card score-card evaluation-score" style="--eval-color:${ev.color}"><span class="label">Puntaje global estimado</span><strong class="big-score">${fmtGlobal(row.puntaje_global)}</strong></article>
      <article class="card metric-card"><span class="label">Respuestas correctas</span><strong>${total===null?'—':`${total}/120`}</strong></article>
    </div>
    <section class="card wide"><div class="section-title"><div><h2>Resultado por área</h2></div></div>${areaRows(row,ev)}${naturalChips(row,data)}</section>
    <section class="preunal-cta"><div><span>Convierte tu resultado en un plan de preparación</span><strong>Prepárate para el examen de admisión UNAL con Proyectia</strong></div><a href="${PREUNAL_URL}" target="_blank" rel="noopener">Inscríbete en nuestro PREUNAL</a></section>
    <section class="grid two"><article class="card"><h2>Temas para revisar</h2>${reviewTopics(data)}</article><article class="card"><h2>Lectura del resultado</h2><p class="friendly-reading">${quickReading(row)}</p></article></section>
    ${currentDetail.length?detailSection():''}`;
  if(currentDetail.length) renderDetail('todas');
  showOnly('#report-view');
}

function detailSection(){
  return `<section class="card wide"><div class="section-title"><div><h2>Detalle por pregunta</h2><span>Desplázate dentro de la ventana para revisar las preguntas y la respuesta correcta.</span></div><label class="filter-label">Área <select id="area-filter" class="area-filter" onchange="renderDetail(this.value)"><option value="todas">Todas las áreas</option>${AREA_CONFIG.map(a=>`<option value="${a.name}">${a.name}</option>`).join('')}</select></label></div><div class="table-wrap detail-window"><table class="responsive-table detail-table"><thead><tr><th>#</th><th>Área</th><th>Tema</th><th>Tu respuesta</th><th>Resultado</th><th>Correcta</th></tr></thead><tbody id="detail-body"></tbody></table></div></section>`;
}
function renderDetail(area){
  const rows=currentDetail.filter(r=>area==='todas'||String(r.area)===area);
  const body=$('#detail-body'); if(!body) return;
  body.innerHTML=rows.map(r=>`<tr>
    <td data-label="Pregunta">${esc(r.pregunta)}</td>
    <td data-label="Área">${esc(r.area)}</td>
    <td data-label="Tema">${esc(r.tema)}</td>
    <td data-label="Tu respuesta">${esc(r.respuesta_estudiante||'—')}</td>
    <td data-label="Resultado"><span class="pill ${String(r.resultado).toLowerCase()==='correcta'?'ok':'bad'}">${esc(r.resultado)}</span></td>
    <td data-label="Correcta">${esc(r.respuesta_correcta||'—')}</td>
  </tr>`).join('');
}

function availableSeries(){
  return currentRecords.slice().sort((a,b)=>a.ev.order-b.ev.order).map(x=>({ev:x.ev,row:x.row,data:x.data,score:globalValue(x.row.puntaje_global)})).filter(x=>x.score!==null);
}
function chartGeometry(width){
  const W=Math.max(280,Math.floor(width||760));
  const H=Math.round(Math.max(255,Math.min(360,W*.42)));
  const L=Math.round(Math.max(42,Math.min(70,W*.075)));
  const R=Math.round(Math.max(12,Math.min(30,W*.03)));
  const T=Math.round(Math.max(24,Math.min(32,W*.035)));
  const B=Math.round(Math.max(58,Math.min(76,W*.085)));
  const labelSize=Math.max(9.5,Math.min(12.5,W/72));
  const valueSize=Math.max(10,Math.min(13,W/68));
  const pointRadius=Math.max(5,Math.min(7,W/120));
  return {W,H,L,R,T,B,labelSize,valueSize,pointRadius};
}
function fixedX(order,W,L,R){ return L + order*(W-L-R)/(EVALUATIONS.length-1); }
function axisLabels(g){
  return EVALUATIONS.map(ev=>`<text class="chart-label chart-eval-label" x="${fixedX(ev.order,g.W,g.L,g.R)}" y="${g.H-g.B+30}" text-anchor="middle" style="fill:${ev.color};font-size:${g.labelSize}px">${esc(ev.short)}</text>`).join('');
}
function pointMarks(vals,x,y,format,g){
  return vals.map(v=>`<circle cx="${x(v.ev.order)}" cy="${y(v.score)}" r="${g.pointRadius}" fill="${v.ev.color}" stroke="#351046" stroke-width="2"/><text class="chart-value" x="${x(v.ev.order)}" y="${y(v.score)-13}" text-anchor="middle" style="font-size:${g.valueSize}px">${format(v.score)}</text>`).join('');
}
function globalChartSvg(records,width){
  const vals=records.filter(r=>r.score!==null);
  if(!vals.length) return `<p class="method-note">No hay puntajes disponibles.</p>`;
  const g=chartGeometry(width);
  let lo=Math.min(...vals.map(r=>r.score)),hi=Math.max(...vals.map(r=>r.score));
  const pad=Math.max(25,(hi-lo)*.28||25); lo-=pad;hi+=pad;
  const x=order=>fixedX(order,g.W,g.L,g.R);
  const y=v=>g.T+(hi-v)*(g.H-g.T-g.B)/(hi-lo||1);
  let grid='';
  for(let i=0;i<=4;i++){
    const v=lo+(hi-lo)*i/4,yy=y(v);
    grid+=`<line class="chart-grid" x1="${g.L}" x2="${g.W-g.R}" y1="${yy}" y2="${yy}"/><text class="chart-label" x="${g.L-8}" y="${yy+4}" text-anchor="end" style="font-size:${g.labelSize}px">${Math.trunc(v)}</text>`;
  }
  const pts=vals.map(v=>`${x(v.ev.order)},${y(v.score)}`).join(' ');
  const line=vals.length>1?`<polyline points="${pts}" fill="none" stroke="#64267d" stroke-width="4.5" stroke-linecap="round" stroke-linejoin="round"/>`:'';
  return `<svg viewBox="0 0 ${g.W} ${g.H}" width="100%" role="img" aria-label="Comparación del puntaje global"><line class="chart-axis" x1="${g.L}" x2="${g.L}" y1="${g.T}" y2="${g.H-g.B}"/><line class="chart-axis" x1="${g.L}" x2="${g.W-g.R}" y1="${g.H-g.B}" y2="${g.H-g.B}"/>${grid}${line}${pointMarks(vals,x,y,v=>Math.trunc(v),g)}${axisLabels(g)}</svg>`;
}
function areaChartSvg(records,area,width){
  const vals=records.map(r=>({ev:r.ev,score:scoreOf(r.row,area)})).filter(x=>x.score!==null);
  if(!vals.length) return `<p class="method-note">No hay puntajes disponibles para esta área.</p>`;
  const g=chartGeometry(width);
  let lo=Math.min(...vals.map(v=>v.score)),hi=Math.max(...vals.map(v=>v.score));
  const pad=Math.max(.35,(hi-lo)*.32||.35); lo-=pad;hi+=pad;
  const x=order=>fixedX(order,g.W,g.L,g.R);
  const y=v=>g.T+(hi-v)*(g.H-g.T-g.B)/(hi-lo||1);
  let grid='';
  for(let i=0;i<=4;i++){
    const v=lo+(hi-lo)*i/4,yy=y(v);
    grid+=`<line class="chart-grid" x1="${g.L}" x2="${g.W-g.R}" y1="${yy}" y2="${yy}"/><text class="chart-label" x="${g.L-8}" y="${yy+4}" text-anchor="end" style="font-size:${g.labelSize}px">${v.toFixed(1)}</text>`;
  }
  const pts=vals.map(v=>`${x(v.ev.order)},${y(v.score)}`).join(' ');
  const line=vals.length>1?`<polyline points="${pts}" fill="none" stroke="#64267d" stroke-width="4.5" stroke-linecap="round" stroke-linejoin="round"/>`:'';
  return `<svg viewBox="0 0 ${g.W} ${g.H}" width="100%" role="img" aria-label="Comparación de ${esc(area)}"><line class="chart-axis" x1="${g.L}" x2="${g.L}" y1="${g.T}" y2="${g.H-g.B}"/><line class="chart-axis" x1="${g.L}" x2="${g.W-g.R}" y1="${g.H-g.B}" y2="${g.H-g.B}"/>${grid}${line}${pointMarks(vals,x,y,v=>v.toFixed(1),g)}${axisLabels(g)}</svg>`;
}
function drawGlobalChart(records){
  const target=$('#global-chart'); if(!target) return;
  target.innerHTML=globalChartSvg(records,target.clientWidth||target.parentElement?.clientWidth||760);
}
function drawAreaChart(records,area){
  const target=$('#area-chart'); if(!target) return;
  target.innerHTML=areaChartSvg(records,area,target.clientWidth||target.parentElement?.clientWidth||760);
}

function scoreHistory(records){
  return `<div class="table-wrap score-history"><table class="responsive-table"><thead><tr><th>Prueba</th><th>Puntaje global</th><th>Diferencia frente a la prueba anterior presentada</th></tr></thead><tbody>${records.map((r,i)=>{const d=i?r.score-records[i-1].score:null;return `<tr><td data-label="Prueba"><span class="table-eval-dot" style="background:${r.ev.color}"></span>${esc(r.ev.label)}</td><td data-label="Puntaje global" class="score-cell"><strong>${r.score}</strong></td><td data-label="Diferencia">${d===null?'—':deltaGlobalHTML(d)}</td></tr>`}).join('')}</tbody></table></div>`;
}
function deltaGlobalHTML(d){ const cls=d>0?'up':d<0?'down':'flat'; return `<span class="delta ${cls}">${d>0?'+':''}${Math.trunc(d)} puntos</span>`; }

function priorityTopicsForArea(area){
  const grouped=new Map();
  for(const rec of currentRecords){
    const topics=topicsFor(rec.data,currentDoc).filter(t=>String(t.area||'').trim()===area);
    for(const t of topics){
      const tema=String(t.tema||'').trim();
      if(!tema) continue;
      const hits=num(t.aciertos), total=num(t.total_preguntas);
      if(total===null || total<=0) continue;
      const key=tema.toLocaleLowerCase('es');
      if(!grouped.has(key)) grouped.set(key,{tema,hits:0,total:0});
      const g=grouped.get(key);
      g.hits += hits===null?0:hits;
      g.total += total;
    }
  }
  const rows=[...grouped.values()]
    .map(x=>({...x,pct:x.total?x.hits/x.total:1}))
    .sort((a,b)=>a.pct-b.pct || b.total-a.total || a.tema.localeCompare(b.tema,'es'));
  if(!rows.length) return '<p class="area-priority-empty">No hay información temática disponible para las pruebas presentadas.</p>';
  return `<div class="area-priority-plain"><h3>Temas que te recomendamos estudiar</h3><ul>${rows.slice(0,5).map(t=>`<li>${esc(t.tema)}</li>`).join('')}</ul></div>`;
}

function renderAreaProgress(area){
  document.querySelectorAll('.area-choice').forEach(b=>b.setAttribute('aria-pressed',b.dataset.area===area?'true':'false'));
  currentAreaProgress=area;
  const records=availableSeries();
  const vals=records.map(r=>({ev:r.ev,score:scoreOf(r.row,area)})).filter(x=>x.score!==null);
  const target=$('#area-progress-content'); if(!target) return;
  if(!vals.length){ target.innerHTML='<p class="method-note">No hay puntajes disponibles para esta área.</p>'; return; }
  const scores=EVALUATIONS.map(ev=>{
    const item=vals.find(v=>v.ev.order===ev.order);
    return `<div class="area-test-score" style="--eval-color:${ev.color}">
      <span>${esc(ev.label)}</span>
      <strong>${item?item.score.toFixed(1):'—'}</strong>
    </div>`;
  }).join('');
  target.innerHTML=`<div class="area-progress-stack">
    <div id="area-chart" class="chart-wrap responsive-chart"></div>
    <div class="area-test-scores">${scores}</div>
    ${priorityTopicsForArea(area)}
  </div>`;
  requestAnimationFrame(()=>drawAreaChart(records,area));
}

function showProgress(){
  const records=availableSeries(); if(!records.length) return;
  const last=records[records.length-1];
  const name=displayName(currentRecords);
  $('#progress-view').innerHTML=`
    ${toolbar('Progresión de tus resultados','')}
    ${studentOwnership(name)}
    <div class="progress-lead-grid">
      <article class="card progress-score"><span class="label">Puntaje más reciente</span><strong class="big-score">${last.score}</strong><p>${esc(last.ev.label)}</p></article>
      <a class="score-reference-card" href="${PUNTAJES_UNAL_URL}" target="_blank" rel="noopener">
        <span>Referencia para tu preparación</span>
        <strong>Revisa qué puntaje necesitas para pasar a la UNAL</strong>
      </a>
    </div>
    <section class="card wide"><div class="section-title"><div><h2>Comparación de tu puntaje total</h2><span>Las cinco pruebas aparecen en la gráfica y se muestran puntos únicamente donde tienes resultados.</span></div></div><p class="chart-intro"><strong>Cómo leer esta gráfica.</strong> Cada punto muestra tu puntaje en una prueba. Si no presentaste una prueba, verás su nombre sin un punto. La línea une los resultados disponibles.</p><div id="global-chart" class="chart-wrap responsive-chart"></div></section>
    <section class="card wide"><div class="section-title"><div><h2>Historial de puntajes</h2><span>Aquí aparecen las pruebas para las que tienes un resultado.</span></div></div>${scoreHistory(records)}</section>
    <section class="card wide"><div class="section-title"><div><h2>Explora la progresión de un área</h2><span>Elige un área para verla por separado y revisar los temas que conviene estudiar primero.</span></div></div><div class="area-selector">${AREA_CONFIG.map(a=>`<button class="area-choice" data-area="${a.name}" aria-pressed="false" onclick="renderAreaProgress('${a.name}')">${a.name}</button>`).join('')}</div><div id="area-progress-content"></div></section>
    <section class="preunal-cta"><div><span>Sigue construyendo tu preparación</span><strong>Prepárate para el examen de admisión UNAL con Proyectia</strong></div><a href="${PREUNAL_URL}" target="_blank" rel="noopener">Inscríbete en nuestro PREUNAL</a></section>`;
  showOnly('#progress-view');
  currentAreaProgress=AREA_CONFIG[0].name;
  renderAreaProgress(currentAreaProgress);
  requestAnimationFrame(()=>drawGlobalChart(records));
}


let chartResizeTimer=null;
window.addEventListener('resize',()=>{
  clearTimeout(chartResizeTimer);
  chartResizeTimer=setTimeout(()=>{
    if(!$('#progress-view')?.classList.contains('hidden')){
      const records=availableSeries();
      drawGlobalChart(records);
      drawAreaChart(records,currentAreaProgress);
    }
  },120);
});
