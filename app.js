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
const STORE = new Map();
let currentDoc = '';
let currentRecords = [];
let currentDetail = [];

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
  const floating=$('#floating-progress');
  if(floating) floating.classList.toggle('hidden',!currentRecords.length || id==='#progress-view');
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
    const floating=$('#floating-progress'); if(floating) floating.classList.add('hidden');
    return;
  }
  currentDoc=doc; currentRecords=found; estado.textContent=''; renderMenu(found);
}
function displayName(found){
  for(let i=found.length-1;i>=0;i--){ const n=getName(found[i].row); if(n) return n; }
  return '';
}

function studentOwnership(name){
  return name ? `<p class="student-ownership">Estos resultados pertenecen a <strong>${esc(name)}</strong>.</p>` : '';
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
        <button class="menu-progress-button" onclick="showProgress()"><span aria-hidden="true">↗</span> Ver mi progresión</button>
      </div>
    </div>`;
  showOnly('#menu-view');
}

function toolbar(title,subtitle,ev=null){
  return `<div class="report-toolbar" ${ev?`style="--eval-color:${ev.color}"`:''}><div>${ev?`<span class="evaluation-badge" style="--eval-color:${ev.color};--eval-text:${ev.text}"><i></i>${esc(ev.short)}</span>`:''}<h2>${esc(title)}</h2>${subtitle?`<p>${subtitle}</p>`:''}</div><div class="toolbar-actions"><button class="btn-small" onclick="renderMenu(currentRecords)">Volver a los resultados</button></div></div>`;
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
  return `<section class="card wide"><div class="section-title"><div><h2>Detalle por pregunta</h2><span>Desplázate dentro de la ventana para revisar las preguntas y la respuesta correcta.</span></div><label class="filter-label">Área <select id="area-filter" class="area-filter" onchange="renderDetail(this.value)"><option value="todas">Todas las áreas</option>${AREA_CONFIG.map(a=>`<option value="${a.name}">${a.name}</option>`).join('')}</select></label></div><div class="table-wrap detail-window"><table><thead><tr><th>#</th><th>Área</th><th>Tema</th><th>Tu respuesta</th><th>Resultado</th><th>Correcta</th></tr></thead><tbody id="detail-body"></tbody></table></div></section>`;
}
function renderDetail(area){
  const rows=currentDetail.filter(r=>area==='todas'||String(r.area)===area);
  const body=$('#detail-body'); if(!body) return;
  body.innerHTML=rows.map(r=>`<tr><td>${esc(r.pregunta)}</td><td>${esc(r.area)}</td><td>${esc(r.tema)}</td><td>${esc(r.respuesta_estudiante||'—')}</td><td><span class="pill ${String(r.resultado).toLowerCase()==='correcta'?'ok':'bad'}">${esc(r.resultado)}</span></td><td>${esc(r.respuesta_correcta||'—')}</td></tr>`).join('');
}

function availableSeries(){
  return currentRecords.slice().sort((a,b)=>a.ev.order-b.ev.order).map(x=>({ev:x.ev,row:x.row,data:x.data,score:globalValue(x.row.puntaje_global)})).filter(x=>x.score!==null);
}
function fixedX(order,W,L,R){ return L + order*(W-L-R)/(EVALUATIONS.length-1); }
function axisLabels(W,H,L,R,B){
  return EVALUATIONS.map(ev=>`<text class="chart-label chart-eval-label" x="${fixedX(ev.order,W,L,R)}" y="${H-B+30}" text-anchor="middle" style="fill:${ev.color}">${esc(ev.short)}</text>`).join('');
}
function pointMarks(vals,x,y,format){
  return vals.map(v=>`<circle cx="${x(v.ev.order)}" cy="${y(v.score)}" r="7" fill="${v.ev.color}" stroke="#351046" stroke-width="2.5"/><text class="chart-value" x="${x(v.ev.order)}" y="${y(v.score)-15}" text-anchor="middle">${format(v.score)}</text>`).join('');
}
function globalChart(records){
  const vals=records.filter(r=>r.score!==null);
  if(!vals.length) return `<p class="method-note">No hay puntajes disponibles.</p>`;
  const W=940,H=350,L=72,R=34,T=30,B=72; let lo=Math.min(...vals.map(r=>r.score)),hi=Math.max(...vals.map(r=>r.score)); const pad=Math.max(25,(hi-lo)*.28||25); lo-=pad;hi+=pad;
  const x=order=>fixedX(order,W,L,R); const y=v=>T+(hi-v)*(H-T-B)/(hi-lo||1);
  let grid=''; for(let i=0;i<=4;i++){const v=lo+(hi-lo)*i/4,yy=y(v);grid+=`<line class="chart-grid" x1="${L}" x2="${W-R}" y1="${yy}" y2="${yy}"/><text class="chart-label" x="${L-10}" y="${yy+4}" text-anchor="end">${Math.trunc(v)}</text>`;}
  const pts=vals.map(v=>`${x(v.ev.order)},${y(v.score)}`).join(' ');
  const line=vals.length>1?`<polyline points="${pts}" fill="none" stroke="#64267d" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"/>`:'';
  return `<div class="chart-wrap"><svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Comparación del puntaje global"><line class="chart-axis" x1="${L}" x2="${L}" y1="${T}" y2="${H-B}"/><line class="chart-axis" x1="${L}" x2="${W-R}" y1="${H-B}" y2="${H-B}"/>${grid}${line}${pointMarks(vals,x,y,v=>Math.trunc(v))}${axisLabels(W,H,L,R,B)}</svg></div>`;
}
function areaChart(records,area){
  const vals=records.map(r=>({ev:r.ev,score:scoreOf(r.row,area)})).filter(x=>x.score!==null);
  if(!vals.length) return `<p class="method-note">No hay puntajes disponibles para esta área.</p>`;
  const W=940,H=350,L=72,R=34,T=30,B=72; let lo=Math.min(...vals.map(v=>v.score)),hi=Math.max(...vals.map(v=>v.score)); const pad=Math.max(.35,(hi-lo)*.32||.35); lo-=pad;hi+=pad;
  const x=order=>fixedX(order,W,L,R); const y=v=>T+(hi-v)*(H-T-B)/(hi-lo||1);
  let grid=''; for(let i=0;i<=4;i++){const v=lo+(hi-lo)*i/4,yy=y(v);grid+=`<line class="chart-grid" x1="${L}" x2="${W-R}" y1="${yy}" y2="${yy}"/><text class="chart-label" x="${L-10}" y="${yy+4}" text-anchor="end">${v.toFixed(1)}</text>`;}
  const pts=vals.map(v=>`${x(v.ev.order)},${y(v.score)}`).join(' ');
  const line=vals.length>1?`<polyline points="${pts}" fill="none" stroke="#64267d" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"/>`:'';
  return `<div class="chart-wrap"><svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Comparación de ${esc(area)}"><line class="chart-axis" x1="${L}" x2="${L}" y1="${T}" y2="${H-B}"/><line class="chart-axis" x1="${L}" x2="${W-R}" y1="${H-B}" y2="${H-B}"/>${grid}${line}${pointMarks(vals,x,y,v=>v.toFixed(1))}${axisLabels(W,H,L,R,B)}</svg></div>`;
}
function scoreHistory(records){
  return `<div class="table-wrap score-history"><table><thead><tr><th>Evaluación</th><th>Puntaje global</th><th>Diferencia frente a la anterior presentada</th></tr></thead><tbody>${records.map((r,i)=>{const d=i?r.score-records[i-1].score:null;return `<tr><td><span class="table-eval-dot" style="background:${r.ev.color}"></span>${esc(r.ev.label)}</td><td><strong>${r.score}</strong></td><td>${d===null?'—':deltaGlobalHTML(d)}</td></tr>`}).join('')}</tbody></table></div>`;
}
function deltaGlobalHTML(d){ const cls=d>0?'up':d<0?'down':'flat'; return `<span class="delta ${cls}">${d>0?'+':''}${Math.trunc(d)} puntos</span>`; }

function areaReading(records,area){
  const vals=records.map(r=>({ev:r.ev,score:scoreOf(r.row,area)})).filter(x=>x.score!==null);
  if(!vals.length) return 'No hay puntajes disponibles para esta área.';
  const list=vals.map(v=>`<li><span class="table-eval-dot" style="background:${v.ev.color}"></span><strong>${esc(v.ev.label)}:</strong> ${v.score.toFixed(1)}</li>`).join('');
  return `<p>Estos son los puntajes registrados en ${esc(area)} para las evaluaciones que has presentado:</p><ul class="compact-score-list">${list}</ul>`;
}
function renderAreaProgress(area){
  document.querySelectorAll('.area-choice').forEach(b=>b.setAttribute('aria-pressed',b.dataset.area===area?'true':'false'));
  const records=availableSeries(); const vals=records.map(r=>({ev:r.ev,score:scoreOf(r.row,area)})).filter(x=>x.score!==null); const target=$('#area-progress-content'); if(!target) return;
  if(!vals.length){ target.innerHTML='<p class="method-note">No hay puntajes disponibles para esta área.</p>'; return; }
  const first=vals[0],last=vals[vals.length-1],max=vals.reduce((a,b)=>a.score>b.score?a:b),d=last.score-first.score;
  target.innerHTML=`<div class="area-progress-grid"><div><div class="area-mini-summary"><div class="mini-stat"><span>Puntaje más reciente</span><strong>${last.score.toFixed(1)}</strong></div><div class="mini-stat"><span>Diferencia entre el primero y el más reciente</span><strong>${vals.length>1?(d>0?'+':'')+d.toFixed(1):'—'}</strong></div><div class="mini-stat"><span>Mayor puntaje registrado</span><strong>${max.score.toFixed(1)}</strong></div></div><p class="chart-intro"><strong>Cómo leer esta gráfica.</strong> Verás un punto por cada evaluación en la que tengas un puntaje disponible para ${esc(area)}. La línea une únicamente esos resultados.</p>${areaChart(records,area)}</div><aside class="area-report"><h3>Resultados de ${esc(area)}</h3>${areaReading(records,area)}</aside></div>`;
}

function progressReading(records){
  if(!records.length) return '';
  if(records.length===1) return `<p class="friendly-reading">Tienes un puntaje global registrado: <strong>${records[0].score}</strong> en ${esc(records[0].ev.label)}. Las demás evaluaciones aparecerán cuando exista información asociada a tu documento.</p>`;
  const first=records[0],last=records[records.length-1],d=last.score-first.score,max=records.reduce((a,b)=>a.score>b.score?a:b);
  return `<p class="friendly-reading">El primer puntaje disponible es <strong>${first.score}</strong> en ${esc(first.ev.label)} y el más reciente es <strong>${last.score}</strong> en ${esc(last.ev.label)}. La diferencia entre ambos es de <strong>${d>0?'+':''}${d} puntos</strong>. El mayor puntaje registrado es <strong>${max.score}</strong> en ${esc(max.ev.label)}.</p>`;
}
function showProgress(){
  const records=availableSeries(); if(!records.length) return;
  const first=records[0],last=records[records.length-1],max=records.reduce((a,b)=>a.score>b.score?a:b),d=last.score-first.score; const name=displayName(currentRecords);
  $('#progress-view').innerHTML=`
    ${toolbar('Progresión de tus resultados','')}
    ${studentOwnership(name)}
    <div class="grid progress-summary">
      <article class="card progress-score"><span class="label">Puntaje más reciente</span><strong class="big-score">${last.score}</strong><p>${esc(last.ev.label)}</p></article>
      <article class="card metric-card"><span class="label">Mayor puntaje registrado</span><strong>${max.score}</strong><p>${esc(max.ev.label)}</p></article>
      <article class="card metric-card"><span class="label">Diferencia entre el primero y el más reciente</span><strong>${records.length>1?(d>0?'+':'')+d:'—'}</strong><p>${records.length>1?'Puntos de diferencia':'Aparecerá con otro resultado'}</p></article>
    </div>
    <section class="card wide"><div class="section-title"><div><h2>Comparación de tu puntaje total</h2><span>Las cinco evaluaciones aparecen en la gráfica y se muestran puntos únicamente donde tienes resultados.</span></div></div><p class="chart-intro"><strong>Cómo leer esta gráfica.</strong> Cada punto muestra tu puntaje en una prueba. Si no presentaste una evaluación, verás su nombre sin un punto. La línea une los resultados disponibles.</p>${globalChart(records)}</section>
    <section class="card wide"><div class="section-title"><div><h2>Historial de puntajes</h2><span>Aquí aparecen las evaluaciones para las que tienes un resultado.</span></div></div>${scoreHistory(records)}</section>
    <section class="card wide"><div class="section-title"><div><h2>Explora la progresión de un área</h2><span>Elige un área para verla por separado.</span></div></div><p class="chart-intro"><strong>Cómo usar esta sección.</strong> Selecciona un área y verás sus puntajes en las evaluaciones que hayas presentado.</p><div class="area-selector">${AREA_CONFIG.map(a=>`<button class="area-choice" data-area="${a.name}" aria-pressed="false" onclick="renderAreaProgress('${a.name}')">${a.name}</button>`).join('')}</div><div id="area-progress-content"></div></section>
    <section class="card wide"><h2>Resumen de tus puntajes</h2>${progressReading(records)}</section>
    <section class="preunal-cta"><div><span>Sigue construyendo tu preparación</span><strong>Prepárate para el examen de admisión UNAL con Proyectia</strong></div><a href="${PREUNAL_URL}" target="_blank" rel="noopener">Inscríbete en nuestro PREUNAL</a></section>`;
  showOnly('#progress-view');
  renderAreaProgress(AREA_CONFIG[0].name);
}

