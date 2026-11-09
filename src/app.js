import { analyze, prepare, run, exportNew, exportInject } from './engine.js';
import { numToCol } from './xlsx-reader.js';

const $ = s => document.querySelector(s);
const JSZip = window.JSZip;
const ExcelJS = window.ExcelJS;
const fmt = new Intl.NumberFormat('es-CO', { maximumFractionDigits: 0 });
const money = v => (typeof v === 'number' ? '$ ' + fmt.format(v) : '—');
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const tick = () => new Promise(r => setTimeout(r, 0));

const ROLE_NAMES = [
  ['item', 'Ítem'], ['apu', 'Código APU'], ['desc', 'Descripción'], ['und', 'Unidad'],
  ['cant', 'Cantidad'], ['vu', 'Valor unitario'], ['vt', 'Valor total'],
];

const st = { data: null, fileName: '', book: null, candidates: [], sheet: null, full: new Map(), prep: null,
  roles: {}, headerRow: null, extras: new Map(), result: null };

function setStatus(el, msg, err = false) { el.textContent = msg; el.classList.toggle('err', err); }

// ---------------------------------------------------------------- carga
const drop = $('#drop');
const input = $('#file');
drop.addEventListener('click', () => input.click());
drop.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('over'); });
drop.addEventListener('dragleave', () => drop.classList.remove('over'));
drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('over'); if (e.dataTransfer.files[0]) load(e.dataTransfer.files[0]); });
input.addEventListener('change', () => { if (input.files[0]) load(input.files[0]); input.value = ''; });

async function load(file) {
  const status = $('#status');
  if (/\.xls$/i.test(file.name)) { setStatus(status, 'Ese es un .xls antiguo. Ábrelo en Excel y guárdalo como .xlsx para poder procesarlo.', true); return; }
  if (!/\.xls[xm]$/i.test(file.name)) { setStatus(status, 'El archivo debe ser .xlsx o .xlsm.', true); return; }
  try {
    setStatus(status, `Leyendo ${file.name} (${(file.size / 1048576).toFixed(1)} MB)…`);
    await tick();
    st.data = await file.arrayBuffer();
    st.fileName = file.name;
    const t0 = performance.now();
    const { book, candidates } = await analyze(st.data, { JSZip, onProgress: (i, n, name) => setStatus(status, `Buscando el presupuesto detallado: hoja ${i} de ${n} (${name})…`) });
    if (!candidates.length) throw new Error('No se encontró ninguna hoja con encabezado de presupuesto (DESCRIPCIÓN, UND, CANT, VR TOTAL).');
    Object.assign(st, { book, candidates, full: new Map(), roles: {}, headerRow: null, extras: new Map() });
    for (const c of candidates) if (c.full) st.full.set(c.sheet.name, c.full);
    st.sheet = candidates[0].sheet;
    fillSheetSelect();
    await reprepare();
    drop.classList.add('loaded');
    drop.querySelector('.drop-main').innerHTML = `${esc(file.name)} <u>cambiar archivo</u>`;
    setStatus(status, `Listo en ${((performance.now() - t0) / 1000).toFixed(1)} s.`);
    $('#work').hidden = false;
  } catch (e) {
    console.error(e);
    setStatus(status, e.message || String(e), true);
  }
}

function fillSheetSelect() {
  const sel = $('#sheet');
  const ranked = new Set(st.candidates.map(c => c.sheet.name));
  const opts = st.candidates.map((c, i) => `<option value="${esc(c.sheet.name)}">${esc(c.sheet.name)}${i === 0 ? ' (detectada)' : ''}${c.stats ? ` — ${c.stats.ues} UE, ${c.stats.acts} actividades` : ''}</option>`);
  const others = st.book.sheets.filter(s => !ranked.has(s.name)).map(s => `<option value="${esc(s.name)}">${esc(s.name)}</option>`);
  sel.innerHTML = opts.join('') + (others.length ? `<optgroup label="Otras hojas">${others.join('')}</optgroup>` : '');
  sel.value = st.sheet.name;
  const n = st.candidates.length;
  $('#sheet-note').textContent = n > 1 ? 'Se eligió la hoja con más estructura de presupuesto. Puedes cambiarla aquí.' : 'Única hoja con estructura de presupuesto.';
}

$('#sheet').addEventListener('change', async e => {
  st.sheet = st.book.sheets.find(s => s.name === e.target.value);
  st.roles = {}; st.headerRow = null; st.extras = new Map();
  await reprepare();
});
$('#hdr').addEventListener('change', async e => {
  const n = parseInt(e.target.value, 10);
  if (!(n > 0)) return;
  st.headerRow = n; st.roles = {}; st.extras = new Map();
  await reprepare();
});

async function reprepare() {
  const status = $('#status');
  try {
    let full = st.full.get(st.sheet.name);
    st.prep = await prepare(st.book, st.sheet, { roles: st.roles, headerRow: st.headerRow, full });
    st.full.set(st.sheet.name, st.prep.full);
    $('#hdr').value = st.prep.header.row;
    renderRoles();
    renderExtras();
    compute();
    setStatus(status, '');
  } catch (e) {
    console.error(e);
    setStatus(status, e.message || String(e), true);
    renderRoles(true);
  }
}

// ---------------------------------------------------------------- columnas
function colOptions(selected) {
  const labels = st.prep?.header.labels || new Map();
  const groups = st.prep?.groups || new Map();
  let html = '<option value="0">Ninguna</option>';
  for (const [c, l] of [...labels].sort((a, b) => a[0] - b[0])) {
    const g = groups.get(c);
    html += `<option value="${c}"${c === selected ? ' selected' : ''}>${numToCol(c)}: ${esc(l)}${g ? ` (${esc(g)})` : ''}</option>`;
  }
  return html;
}
function renderRoles(failed = false) {
  const box = $('#roles');
  const roles = st.prep?.header.roles || {};
  box.innerHTML = ROLE_NAMES.map(([k, name]) => `<label class="field">${name}<select data-role="${k}">${colOptions(roles[k])}</select></label>`).join('');
  box.querySelectorAll('select').forEach(sel => sel.addEventListener('change', async () => {
    st.roles = { ...(st.prep?.header.roles || {}), [sel.dataset.role]: parseInt(sel.value, 10) || 0 };
    await reprepare();
  }));
  if (failed) box.querySelectorAll('select').forEach(s => { s.disabled = false; });
}
function renderExtras() {
  const list = st.prep.available.filter(a => !a.role);
  $('#extras-count').textContent = `(${[...st.extras.values()].filter(x => x.on).length} de ${list.length})`;
  $('#extras').innerHTML = list.map(a => {
    const sel = st.extras.get(a.col) || { on: false, sum: a.sumDefault };
    return `<div class="xrow">
      <input type="checkbox" id="x${a.col}" data-col="${a.col}" class="x-on"${sel.on ? ' checked' : ''}>
      <span class="col">Col. ${a.letter}</span>
      <label for="x${a.col}">${esc(a.label)}${a.group ? `<span class="grp">${esc(a.group)}</span>` : ''}</label>
      <label class="sum"><input type="checkbox" data-col="${a.col}" class="x-sum"${sel.sum ? ' checked' : ''}${['money', 'qty'].includes(a.kind) || a.sumDefault ? '' : ''}> sumar</label>
    </div>`;
  }).join('') || '<p class="hint">No hay más columnas con título en el encabezado.</p>';
  $('#extras').querySelectorAll('input').forEach(i => i.addEventListener('change', () => {
    const col = +i.dataset.col;
    const a = st.prep.available.find(x => x.col === col);
    const cur = st.extras.get(col) || { on: false, sum: a.sumDefault };
    if (i.classList.contains('x-on')) cur.on = i.checked; else { cur.sum = i.checked; if (i.checked) cur.on = true; }
    st.extras.set(col, cur);
    renderExtras();
    compute();
  }));
}

// ---------------------------------------------------------------- cálculo y vista previa
function compute() {
  const extras = [...st.extras].filter(([, v]) => v.on).map(([col, v]) => ({ ...st.prep.available.find(a => a.col === col), sum: v.sum }));
  st.result = run(st.prep, { extras, fileName: st.fileName });
  renderSummary();
}

function renderSummary() {
  const { layout } = st.result;
  const s = layout.summary;
  const ok = s.closesAll;
  const badUE = s.ueChecks.filter(u => u.ok === false);
  $('#titleblock').innerHTML = `
    <div class="tb wide"><span class="k">Proyecto</span><span class="v">${esc((s.project || '').replace(/^PROYECTO:?\s*/i, '') || 'Sin nombre en la hoja')}</span></div>
    <div class="tb wide"><span class="k">Archivo</span><span class="v">${esc(st.fileName)}</span></div>
    <div class="tb"><span class="k">Hoja</span><span class="v">${esc(st.sheet.name)}</span></div>
    <div class="tb"><span class="k">Encabezado</span><span class="v">Fila ${st.prep.header.row}</span></div>
    <div class="tb num"><span class="k">Unidades estructurales</span><span class="v">${s.ues}</span></div>
    <div class="tb num"><span class="k">Capítulos</span><span class="v">${s.chapters}</span></div>
    <div class="tb num"><span class="k">Actividades</span><span class="v">${fmt.format(s.acts)}</span></div>
    <div class="tb num"><span class="k">Alertas</span><span class="v">${s.alerts.length}</span></div>
    <div class="tb close ${ok ? 'ok' : 'bad'}">
      <span class="state">${ok ? 'Cierra exacto' : 'No cierra'}</span>
      <span class="pair">Costos directos: original ${money(s.cd.orig)}, reorganizado ${money(s.cd.reorg)}</span>
      <span class="pair">Total presupuesto: original ${money(s.total.orig)}, reorganizado ${money(s.total.reorg)}</span>
    </div>`;
  const alerts = [...s.alerts];
  for (const u of badUE) alerts.unshift({ msg: `La unidad "${u.ue.label}" no cuadra contra su SUB-TOTAL: original ${money(u.orig)}, actividades ${money(u.reorg)}.` });
  $('#alerts').innerHTML = alerts.length ? `<details class="alerts"${badUE.length ? ' open' : ''}><summary>${alerts.length} ${alerts.length === 1 ? 'alerta' : 'alertas'} de estructura (revisa antes de usar)</summary><ul>${alerts.map(a => `<li>${esc(a.msg)}</li>`).join('')}</ul></details>` : '';

  // árbol desde la hoja de salida
  const rows = layout.main.rows;
  const caps = [];
  let cur = null; let ue = null;
  for (const r of rows) {
    if (r.type === 'cap') { cur = { code: r.cells.item?.v || '', name: r.cells.desc?.v || '', v: r.cells.vt?.v, ues: [] }; caps.push(cur); }
    else if (r.type === 'ue' && cur) { ue = { code: r.cells.item?.v || '', name: r.cells.desc?.v || '', v: r.cells.vt?.v, n: 0 }; cur.ues.push(ue); }
    else if (r.type === 'act' && ue) ue.n++;
  }
  $('#tree').innerHTML = caps.map(c => `<details class="cap"><summary><span class="code">${esc(c.code)}</span><span>${esc(c.name)}</span><span class="money">${money(c.v)}</span></summary>
    ${c.ues.map(u => `<div class="ue"><span class="code"></span><span class="lbl">${esc([u.code, u.name].filter(Boolean).join(' '))}</span><span class="n">${u.n} act.</span><span class="money">${money(u.v)}</span></div>`).join('')}
  </details>`).join('') + `<div class="tree-foot"><span>Total presupuesto</span><span class="money">${money(s.total.reorg)}</span></div>`;
}

// ---------------------------------------------------------------- descarga
$('#go').addEventListener('click', async () => {
  const btn = $('#go');
  const status = $('#go-status');
  const mode = document.querySelector('input[name=mode]:checked').value;
  btn.disabled = true;
  try {
    setStatus(status, mode === 'inject' ? 'Armando el libro completo con las hojas vinculadas…' : 'Armando el libro nuevo…');
    await tick();
    const base = st.fileName.replace(/\.(xlsx|xlsm)$/i, '');
    const ext = (/\.xlsm$/i.test(st.fileName) && mode === 'inject') ? '.xlsm' : '.xlsx';
    let bytes;
    if (mode === 'inject') bytes = (await exportInject(st.data, st.result, st.sheet.name, { JSZip })).data;
    else bytes = await exportNew(st.result, { ExcelJS });
    const name = `${base}_por_capitulos${mode === 'inject' ? '_vinculado' : ''}${ext}`;
    const type = ext === '.xlsm' ? 'application/vnd.ms-excel.sheet.macroEnabled.12' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const url = URL.createObjectURL(new Blob([bytes], { type }));
    const a = Object.assign(document.createElement('a'), { href: url, download: name });
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    setStatus(status, `Descargado: ${name}`);
  } catch (e) {
    console.error(e);
    setStatus(status, 'No se pudo generar el archivo: ' + (e.message || e), true);
  } finally {
    btn.disabled = false;
  }
});
