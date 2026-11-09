// Construye la estructura de salida (independiente del formato de escritura):
// hoja principal CAPÍTULO > UE > ACTIVIDAD, hoja de verificación y matriz CAP x UE.
import { numToCol } from './xlsx-reader.js';
import { norm, isNum, isErr, fmtCode } from './classify.js';

export const SHEET_MAIN = 'PPTO x CAPITULO';
export const SHEET_CHECK = 'VERIFICACION';
export const SHEET_MATRIX = 'MATRIZ CAP x UE';

const num = v => (isNum(v) ? v : 0);

/** Columnas disponibles en el encabezado detectado (para la interfaz). */
export function availableColumns(header, groups) {
  const R = header.roles;
  const roleOf = new Map(Object.entries(R).map(([k, c]) => [c, k]));
  const out = [];
  for (const [c, label] of [...header.labels].sort((a, b) => a[0] - b[0])) {
    const t = norm(label);
    const kind = /UNIT|PRECIO/.test(t) ? 'price' : /CANT/.test(t) ? 'qty' : /%|INC|PORC/.test(t) ? 'pct'
      : /TOTAL|VALOR|PARCIAL|VR\b|M2/.test(t) ? 'money' : 'text';
    out.push({ col: c, letter: numToCol(c), label, group: (c > (R.und || R.desc) ? groups.get(c) : '') || '', role: roleOf.get(c) || null, kind,
      sumDefault: kind === 'money' && !/M2/.test(t) });
  }
  return out;
}

/** Columnas de salida por defecto: núcleo + extras elegidas. */
export function buildColumns(header, available, extras = []) {
  const R = header.roles;
  const lab = c => available.find(a => a.col === c);
  const cols = [];
  const add = (key, label, src, kind, sum = false) => cols.push({ key, label, src, kind, sum, group: src ? (lab(src)?.group || '') : '' });
  add('item', 'ÍTEM', R.item || null, 'code');
  if (R.apu) add('apu', 'APU', R.apu, 'code');
  add('desc', 'DESCRIPCIÓN', R.desc, 'text');
  add('und', 'UND', R.und || null, 'unit');
  if (R.cant) add('cant', lab(R.cant)?.label || 'CANT', R.cant, 'qty');
  if (R.vu) add('vu', lab(R.vu)?.label || 'VR. UNITARIO', R.vu, 'price');
  add('vt', lab(R.vt)?.label || 'VR. TOTAL', R.vt, 'money', true);
  for (const e of extras) {
    if (Object.values(R).includes(e.col)) continue;
    add('x' + e.col, e.label, e.col, e.kind === 'text' ? 'text' : e.kind, !!e.sum);
  }
  add('ue', 'UNIDAD ESTRUCTURAL', null, 'text');
  add('cap', 'CAPÍTULO', null, 'text');
  add('grupo', 'SUBGRUPO', null, 'text');
  add('origen', 'FILA ORIGEN', null, 'int');
  cols.forEach((c, i) => { c.idx = i + 1; c.letter = numToCol(i + 1); });
  return cols;
}

function projectInfo(rows, headerRow) {
  const lines = [];
  for (const [r, row] of rows) {
    if (r >= headerRow) break;
    const parts = [...row.cells.values()].sort((a, b) => a.col - b.col)
      .map(c => c.v).filter(v => typeof v === 'string' && v.trim()).map(v => v.trim());
    if (parts.length) lines.push(parts.join(' '));
  }
  const proj = lines.find(l => /PROYECTO|OBRA/i.test(l)) || lines[1] || lines[0] || '';
  return { company: lines[0] || '', project: proj, lines };
}

/**
 * Devuelve { sheets:[{name, rows:[{type, level, cells:{key:{v,f,link}}}], cols, ...}], checks }
 * cell: { v } valor fijo | { link:{row,col}, v } vínculo a la hoja origen | { f, v } fórmula propia
 */
export function buildLayout(model, pv, cols, { sourceName, fileName } = {}) {
  const src = model.rows;
  const get = (r, c) => (c ? src.get(r)?.cells.get(c) : null);
  const C = Object.fromEntries(cols.map(c => [c.key, c]));
  const sumCols = cols.filter(c => c.sum);
  const info = projectInfo(src, model.header.row);
  const rows = [];
  const push = (type, level, cells = {}) => { const r = { type, level, cells, n: rows.length + 1 }; rows.push(r); return r; };

  // ---------------- encabezado
  push('title', 0, { [cols[0].key]: { v: 'PRESUPUESTO DETALLADO POR CAPÍTULOS' } });
  push('subtitle', 0, { [cols[0].key]: { v: info.project || info.company } });
  push('subtitle', 0, { [cols[0].key]: { v: `Fuente: hoja "${sourceName}"${fileName ? ` de ${fileName}` : ''} · Estructura CAPÍTULO › UNIDAD ESTRUCTURAL › ACTIVIDAD` } });
  push('blank', 0);
  if (cols.some(c => c.group)) push('groupHeader', 0, Object.fromEntries(cols.filter(c => c.group).map(c => [c.key, { v: c.group }])));
  const hdr = push('header', 0, Object.fromEntries(cols.map(c => [c.key, { v: c.label }])));

  const actCells = (a, ueLabel, capLabel) => {
    const cells = {};
    for (const c of cols) {
      if (c.src) {
        const cell = get(a.row, c.src);
        if (cell && cell.v != null && cell.v !== '') {
          cells[c.key] = { link: { row: a.row, col: c.src }, v: isErr(cell.v) ? 0 : cell.v, s: cell.s, err: isErr(cell.v) ? cell.v.error : null };
          if (c.kind === 'code' && typeof cell.v === 'number') cells[c.key].text = fmtCode(cell.v);
        }
      }
    }
    cells.ue = { v: ueLabel };
    cells.cap = { v: capLabel };
    if (a.group) cells.grupo = { v: a.group };
    cells.origen = { v: a.row, origin: a.row };
    return cells;
  };
  const sumOver = (rowsList, colKey) => rowsList.reduce((s, r) => s + num(r.cells[colKey]?.v), 0);
  const setRangeSum = (target, first, last, members) => {
    for (const c of sumCols) target.cells[c.key] = { f: { range: [first, last] }, v: sumOver(members, c.key) };
  };
  const setListSum = (target, members) => {
    for (const c of sumCols) target.cells[c.key] = { f: { list: members.map(m => m.n) }, v: sumOver(members, c.key) };
  };

  // ---------------- costos directos
  const capRows = [];
  const capOut = [];
  const skippedEmpty = [];
  for (const ch of pv.chapters) {
    if (!ch.ueList.length) { skippedEmpty.push(ch.code || ch.name); continue; }
    const capLabel = [ch.code, ch.name].filter(Boolean).join(' - ');
    const capRow = push('cap', 0, { item: { v: ch.code || '' }, desc: { v: ch.name } });
    const ueRows = [];
    for (const { ue, acts } of ch.ueList) {
      const ueRow = push('ue', 1, { item: { v: ue.code }, desc: { v: ue.name } });
      const members = [];
      let lastGroup = null;
      for (const a of acts) {
        if (a.group && a.group !== lastGroup) push('sub', 2, { desc: { v: a.group } });
        lastGroup = a.group;
        members.push(push('act', 2, actCells(a, ue.label, capLabel)));
      }
      setRangeSum(ueRow, members[0].n, members[members.length - 1].n, members);
      ueRow.ueRef = ue; ueRows.push(ueRow);
    }
    setListSum(capRow, ueRows);
    capRow.chapter = ch; capRow.capLabel = capLabel;
    capRows.push(capRow); capOut.push({ ch, row: capRow, label: capLabel });
  }
  const cdRow = push('total', 0, { desc: { v: 'SUB-TOTAL COSTOS DIRECTOS' } });
  setListSum(cdRow, capRows);

  // ---------------- costos indirectos (se conservan en su orden original)
  let indRow = null;
  const indActs = [];
  if (model.indirect.some(e => e.kind === 'act')) {
    push('blank', 0);
    indRow = push('cap', 0, { desc: { v: 'COSTOS INDIRECTOS (ADMINISTRACIÓN, GASTOS GENERALES)' } });
    const ueBlocks = [];
    let block = null;
    let lastCh = null;
    for (const e of model.indirect) {
      if (e.kind === 'ue') {
        block = { row: push('ue', 1, { item: { v: e.ue.code }, desc: { v: e.ue.name } }), members: [] };
        ueBlocks.push(block); lastCh = null;
      } else if (e.kind === 'chapter') {
        lastCh = [fmtCode(e.occ.rawCode), e.occ.name].filter(Boolean).join(' - ');
        push('sub', 2, { item: { v: fmtCode(e.occ.rawCode) }, desc: { v: e.occ.name } });
      } else if (e.kind === 'sub') {
        push('sub', 2, { desc: { v: e.label } });
      } else if (e.kind === 'act') {
        if (!block) { block = { row: push('ue', 1, { desc: { v: 'COSTOS INDIRECTOS' } }), members: [] }; ueBlocks.push(block); }
        const r = push('act', 2, actCells(e.act, e.act.ue.label, lastCh || 'COSTOS INDIRECTOS'));
        block.members.push(r); indActs.push(r);
      }
    }
    const withActs = ueBlocks.filter(b => b.members.length);
    for (const b of withActs) setRangeSum(b.row, b.members[0].n, b.members[b.members.length - 1].n, b.members);
    for (const b of ueBlocks.filter(b => !b.members.length)) b.row.type = 'sub';
    setListSum(indRow, withActs.map(b => b.row));
  }
  push('blank', 0);
  const totRow = push('grand', 0, { desc: { v: 'TOTAL PRESUPUESTO' } });
  setListSum(totRow, indRow ? [cdRow, indRow] : [cdRow]);

  const main = { name: SHEET_MAIN, cols, rows, headerRow: hdr.n, cdRow, indRow, totRow, capOut };

  // ---------------- verificación
  const vcols = [{ key: 'a', w: 46 }];
  const vrows = [];
  const vpush = (type, cells = {}) => { const r = { type, cells, n: vrows.length + 1 }; vrows.push(r); return r; };
  const ref = (row, col) => ({ sheet: SHEET_MAIN, row, col });
  vpush('title', { a: { v: 'VERIFICACIÓN DE CIERRE' } });
  vpush('subtitle', { a: { v: 'Original = valores de la hoja fuente · Reorganizado = hoja "' + SHEET_MAIN + '" · tolerancia 0,01' } });
  vpush('blank');
  // cierre general, una terna de columnas por cada columna sumada
  const triple = [];
  sumCols.forEach((c, i) => {
    const base = 2 + i * 3;
    triple.push({ c, o: numToCol(base), r: numToCol(base + 1), d: numToCol(base + 2) });
    vcols.push({ key: 'o' + i, w: 20 }, { key: 'r' + i, w: 20 }, { key: 'd' + i, w: 14 });
  });
  const estadoCol = numToCol(2 + sumCols.length * 3);
  vcols.push({ key: 'est', w: 12 });
  vcols.forEach((c, i) => { c.letter = numToCol(i + 1); });
  const vk = (i, t) => (t === 'o' ? 'o' : t === 'r' ? 'r' : 'd') + i;

  const hg = vpush('groupHeader', { a: { v: 'CIERRE GENERAL' } });
  sumCols.forEach((c, i) => { hg.cells[vk(i, 'o')] = { v: [c.group, c.label].filter(Boolean).join(' · ') }; });
  const hh = vpush('header', { a: { v: 'CONCEPTO' }, est: { v: 'ESTADO' } });
  sumCols.forEach((c, i) => { hh.cells[vk(i, 'o')] = { v: 'ORIGINAL' }; hh.cells[vk(i, 'r')] = { v: 'REORGANIZADO' }; hh.cells[vk(i, 'd')] = { v: 'DIFERENCIA' }; });

  const origVal = (r, c) => num(get(r, c.src)?.v);
  const directUEs = model.ues.filter(u => u.section === 'direct');
  const indirectUEs = model.ues.filter(u => u.section === 'indirect');
  const addCheck = (label, origSpec, reorgRow) => {
    const r = vpush('row', { a: { v: label } });
    sumCols.forEach((c, i) => {
      const o = origSpec(c);
      r.cells[vk(i, 'o')] = o;
      r.cells[vk(i, 'r')] = { f: { ref: ref(reorgRow.n, c.idx) }, v: num(reorgRow.cells[c.key]?.v) };
      r.cells[vk(i, 'd')] = { f: { diff: [triple[i].r, triple[i].o] }, v: num(reorgRow.cells[c.key]?.v) - num(o.v) };
    });
    r.cells.est = { f: { status: triple.map(t => t.d) }, v: sumCols.every((c, i) => Math.abs(r.cells[vk(i, 'd')].v) < 0.01) ? 'OK' : 'REVISAR' };
    return r;
  };
  const linkOrSum = (rowNum, fallbackUEs) => c => {
    if (rowNum) return { link: { row: rowNum, col: c.src }, v: origVal(rowNum, c) };
    const parts = fallbackUEs.filter(u => u.subtotalRow).map(u => ({ row: u.subtotalRow, col: c.src }));
    return { links: parts, v: parts.reduce((s, p) => s + origVal(p.row, c), 0) };
  };
  const checkCD = addCheck('Costos directos', linkOrSum(model.cdRow, directUEs), cdRow);
  if (indRow) addCheck('Costos indirectos', c => {
    const parts = indirectUEs.filter(u => u.subtotalRow).map(u => ({ row: u.subtotalRow, col: c.src }));
    return { links: parts, v: parts.reduce((s, p) => s + origVal(p.row, c), 0) };
  }, indRow);
  const checkTot = addCheck('Total presupuesto', linkOrSum(model.totalRow, model.ues), totRow);

  // por UE (columna principal)
  const vt = cols.find(c => c.key === 'vt');
  vpush('blank');
  vpush('groupHeader', { a: { v: 'POR UNIDAD ESTRUCTURAL · ' + vt.label } });
  vpush('header', { a: { v: 'UNIDAD ESTRUCTURAL' }, o0: { v: 'SUB-TOTAL ORIGINAL' }, r0: { v: 'REORGANIZADO' }, d0: { v: 'DIFERENCIA' }, est: { v: 'ESTADO' } });
  const ueChecks = [];
  for (const u of model.ues) {
    const reorg = u.acts.reduce((s, a) => s + num(get(a.row, vt.src)?.v), 0);
    const orig = u.subtotalRow ? { link: { row: u.subtotalRow, col: vt.src }, v: origVal(u.subtotalRow, vt) } : { v: 'sin sub-total' };
    const r = vpush('row', {
      a: { v: `${u.label}  (${u.acts.length} act.)` },
      o0: orig,
      r0: { f: { sumifs: { sumCol: vt.letter, critCol: C.ue.letter, crit: u.label } }, v: reorg },
    });
    if (u.subtotalRow) {
      r.cells.d0 = { f: { diff: ['C', 'B'] }, v: reorg - orig.v };
      r.cells.est = { f: { status: ['D'] }, v: Math.abs(reorg - orig.v) < 0.01 ? 'OK' : 'REVISAR' };
    }
    ueChecks.push({ ue: u, orig: orig.v, reorg, ok: u.subtotalRow ? Math.abs(reorg - orig.v) < 0.01 : null });
  }
  // por capítulo
  vpush('blank');
  vpush('groupHeader', { a: { v: 'POR CAPÍTULO · ' + vt.label + ' (original = suma de las filas de capítulo de cada UE)' } });
  vpush('header', { a: { v: 'CAPÍTULO' }, o0: { v: 'ORIGINAL' }, r0: { v: 'REORGANIZADO' }, d0: { v: 'DIFERENCIA' }, est: { v: 'ESTADO' } });
  const capChecks = [];
  for (const co of capOut) {
    const occ = co.ch.occ.filter(o => o.row != null && co.ch.ueList.some(x => x.ue.id === o.ue?.id));
    const vals = occ.map(o => get(o.row, vt.src)?.v);
    const complete = occ.length > 0 && vals.every(isNum);
    const reorg = num(co.row.cells.vt?.v);
    const r = vpush('row', { a: { v: co.label }, r0: { f: { ref: ref(co.row.n, vt.idx) }, v: reorg } });
    if (complete) {
      const o = vals.reduce((s, v) => s + v, 0);
      r.cells.o0 = { links: occ.map(x => ({ row: x.row, col: vt.src })), v: o };
      r.cells.d0 = { f: { diff: ['C', 'B'] }, v: reorg - o };
      r.cells.est = { f: { status: ['D'] }, v: Math.abs(reorg - o) < 0.01 ? 'OK' : 'REVISAR' };
      capChecks.push({ label: co.label, orig: o, reorg, ok: Math.abs(reorg - o) < 0.01 });
    } else {
      r.cells.o0 = { v: 'sin valor en filas de capítulo' };
      capChecks.push({ label: co.label, orig: null, reorg, ok: null });
    }
  }
  // alertas
  const alerts = [...pv.alerts];
  if (skippedEmpty.length) alerts.push({ type: 'cap_vacios', msg: `Capítulos sin actividades (no se listan): ${skippedEmpty.join(', ')}.` });
  vpush('blank');
  vpush('groupHeader', { a: { v: `ALERTAS (${alerts.length})` } });
  for (const al of alerts) vpush('note', { a: { v: al.msg } });
  if (!alerts.length) vpush('note', { a: { v: 'Sin alertas.' } });
  const check = { name: SHEET_CHECK, cols: vcols, rows: vrows, headerRow: hh.n };

  // ---------------- matriz CAP x UE (columna principal)
  const ueList = directUEs.filter(u => u.acts.length);
  const mcols = [{ key: 'a', w: 42, letter: 'A' }];
  ueList.forEach((u, i) => mcols.push({ key: 'u' + i, w: 18, letter: numToCol(i + 2), ue: u }));
  const totLetter = numToCol(ueList.length + 2);
  mcols.push({ key: 't', w: 20, letter: totLetter });
  const mrows = [];
  const mpush = (type, cells = {}) => { const r = { type, cells, n: mrows.length + 1 }; mrows.push(r); return r; };
  mpush('title', { a: { v: 'MATRIZ CAPÍTULO x UNIDAD ESTRUCTURAL · ' + vt.label } });
  mpush('subtitle', { a: { v: 'Costos directos. Cada celda suma las actividades de la hoja "' + SHEET_MAIN + '".' } });
  mpush('blank');
  const mh = mpush('header', { a: { v: 'CAPÍTULO' }, t: { v: 'TOTAL CAPÍTULO' } });
  ueList.forEach((u, i) => { mh.cells['u' + i] = { v: u.label }; });
  const firstData = mrows.length + 1;
  for (const co of capOut) {
    const r = mpush('row', { a: { v: co.label } });
    let tot = 0;
    ueList.forEach((u, i) => {
      const entry = co.ch.ueList.find(x => x.ue.id === u.id);
      const v = entry ? entry.acts.reduce((s, a) => s + num(get(a.row, vt.src)?.v), 0) : 0;
      tot += v;
      r.cells['u' + i] = { f: { sumifs2: { sumCol: vt.letter, c1: C.cap.letter, k1: co.label, c2: C.ue.letter, k2: u.label } }, v };
    });
    r.cells.t = { f: { rowSum: ['B', numToCol(ueList.length + 1)] }, v: tot };
  }
  const lastData = mrows.length;
  const mt = mpush('total', { a: { v: 'TOTAL UE' } });
  for (const c of mcols.slice(1)) {
    const v = mrows.slice(firstData - 1, lastData).reduce((s, r) => s + num(r.cells[c.key]?.v), 0);
    mt.cells[c.key] = { f: { colSum: [c.letter, firstData, lastData] }, v };
  }
  const mchk = mpush('row', { a: { v: 'Control vs. SUB-TOTAL COSTOS DIRECTOS (debe ser 0)' } });
  mchk.cells.t = { f: { diffRef: [totLetter + mt.n, ref(cdRow.n, vt.idx)] }, v: Math.round((mt.cells.t.v - num(cdRow.cells.vt?.v)) * 100) / 100 };
  const matrix = { name: SHEET_MATRIX, cols: mcols, rows: mrows, headerRow: mh.n, freezeCol: 1 };

  const summary = {
    project: info.project, ues: model.ues.length, acts: model.acts.length,
    chapters: capOut.length, cd: { orig: checkCD.cells.o0.v, reorg: num(cdRow.cells.vt?.v) },
    total: { orig: checkTot.cells.o0.v, reorg: num(totRow.cells.vt?.v) },
    closesAll: [checkCD, checkTot].every(r => r.cells.est.v === 'OK') && ueChecks.every(u => u.ok !== false),
    ueChecks, capChecks, alerts,
  };
  return { sheets: [main, check, matrix], main, summary };
}

// ---------------------------------------------------------------- fórmulas a texto
export function quoteSheet(name) { return `'${String(name).replace(/'/g, "''")}'`; }

/** Traduce la especificación de fórmula a texto de Excel (sin '='). */
export function formulaText(spec, { col, row, names = {} }) {
  const L = col;
  const nm = n => names[n] || n;
  const main = nm(SHEET_MAIN);
  if (spec.range) return `SUM(${L}${spec.range[0]}:${L}${spec.range[1]})`;
  if (spec.list) {
    if (!spec.list.length) return '0';
    // agrupa en bloques para no exceder 255 argumentos
    const parts = spec.list.map(n => `${L}${n}`);
    const chunks = [];
    for (let i = 0; i < parts.length; i += 200) chunks.push(`SUM(${parts.slice(i, i + 200).join(',')})`);
    return chunks.join('+');
  }
  if (spec.ref) return `${quoteSheet(nm(spec.ref.sheet))}!${numToCol(spec.ref.col)}${spec.ref.row}`;
  if (spec.diff) return `${spec.diff[0]}${row}-${spec.diff[1]}${row}`;
  if (spec.status) return `IF(AND(${spec.status.map(c => `ABS(${c}${row})<0.01`).join(',')}),"OK","REVISAR")`;
  if (spec.sumifs) {
    const s = quoteSheet(main);
    return `SUMIFS(${s}!$${spec.sumifs.sumCol}:$${spec.sumifs.sumCol},${s}!$${spec.sumifs.critCol}:$${spec.sumifs.critCol},${strLit(spec.sumifs.crit)})`;
  }
  if (spec.sumifs2) {
    const s = quoteSheet(main); const q = spec.sumifs2;
    return `SUMIFS(${s}!$${q.sumCol}:$${q.sumCol},${s}!$${q.c1}:$${q.c1},${strLit(q.k1)},${s}!$${q.c2}:$${q.c2},${strLit(q.k2)})`;
  }
  if (spec.rowSum) return `SUM(${spec.rowSum[0]}${row}:${spec.rowSum[1]}${row})`;
  if (spec.colSum) return `SUM(${spec.colSum[0]}${spec.colSum[1]}:${spec.colSum[0]}${spec.colSum[2]})`;
  if (spec.diffRef) return `ROUND(${spec.diffRef[0]}-${quoteSheet(nm(spec.diffRef[1].sheet))}!${numToCol(spec.diffRef[1].col)}${spec.diffRef[1].row},2)`;
  throw new Error('fórmula desconocida ' + JSON.stringify(spec));
}

// Los criterios de SUMIFS tratan * ? ~ como comodines: se escapan con ~
function strLit(s) {
  const esc = String(s).replace(/[~*?]/g, m => '~' + m).replace(/"/g, '""');
  return `"=${esc}"`; // "=" fuerza coincidencia exacta
}
export function srcRef(sourceName, row, col) { return `${quoteSheet(sourceName)}!${numToCol(col)}${row}`; }
