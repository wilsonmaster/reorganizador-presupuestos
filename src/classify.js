// Detección de la hoja de presupuesto detallado, del encabezado y clasificación
// de filas en: unidad estructural (UE), capítulo, subgrupo, actividad, subtotal.
import { readHead, parseSheetXml, loadSheet, numToCol } from './xlsx-reader.js';

export function norm(s) {
  if (s == null) return '';
  return String(s).normalize('NFKD').replace(/[̀-ͯ]/g, '').toUpperCase()
    .replace(/\s+/g, ' ').trim();
}

export const isNum = v => typeof v === 'number' && Number.isFinite(v);
export const isErr = v => v != null && typeof v === 'object' && 'error' in v;

export function fmtCode(v) {
  if (v == null) return '';
  if (typeof v === 'number') return String(Number(v.toPrecision(10)));
  if (isErr(v)) return v.error;
  return String(v).trim();
}

const val = (row, col) => (col ? row?.cells.get(col)?.v ?? null : null);
const txt = (row, col) => { const v = val(row, col); return v == null || isErr(v) ? '' : (typeof v === 'string' ? v.trim() : fmtCode(v)); };

// ---------------------------------------------------------------- encabezado
const ROLE_RULES = [
  ['apu', t => /^(APU|COD\.? ?APU|CODIGO APU|APU N[O°º]?\.?)$/.test(t)],
  ['item', t => /^(ITEM|ITEMS|IT\.?|N[O°º]\.?|NUMERAL|COD\.?|CODIGO|COD\.? ?ITEM|CAP\.?\/ITEM)$/.test(t)],
  ['desc', t => /DESCRIP|^ACTIVIDAD(ES)?$|^CONCEPTO$|^NOMBRE( ACTIVIDAD)?$/.test(t)],
  ['und', t => /^(UND|UN|UNID|UNIDAD|UNIDADES|U|UM|U\.M\.)\.?$/.test(t)],
  ['cant', t => /^CANT/.test(t)],
  ['vu', t => /UNIT|^PRECIO( UNITARIO)?\b|^P\.? ?U\.?$|^V\.? ?U\.?$/.test(t)],
  ['vt', t => /TOTAL|^PARCIAL|^VALOR$|^IMPORTE/.test(t)],
];

function headerCandidates(rows, maxRow = 80, minScore = 3, only = null) {
  const out = [];
  for (const [r, row] of rows) {
    if (r > maxRow) break;
    if (only != null && r !== only) continue;
    const roles = {};
    const labels = new Map();
    for (const [c, cell] of [...row.cells].sort((a, b) => a[0] - b[0])) {
      if (typeof cell.v !== 'string') continue;
      const t = norm(cell.v);
      if (!t) continue;
      labels.set(c, cell.v.trim());
      for (const [role, test] of ROLE_RULES) {
        if (roles[role] == null && test(t)) {
          // vu / vt / und / cant deben estar a la derecha de la descripción
          if (['und', 'cant', 'vu', 'vt'].includes(role) && roles.desc != null && c < roles.desc) continue;
          roles[role] = c;
          break;
        }
      }
    }
    if (roles.desc == null) continue;
    const score = ['desc', 'und', 'cant', 'vu', 'vt'].filter(k => roles[k] != null).length + (roles.item != null || roles.apu != null ? 0.5 : 0);
    if (score >= minScore) out.push({ row: r, roles, labels, score });
  }
  return out.sort((a, b) => b.score - a.score || a.row - b.row);
}

export function headerAtRow(rows, r) {
  const h = headerCandidates(rows, r, 0, r)[0];
  if (h) return h;
  const labels = new Map();
  for (const [c, cell] of rows.get(r)?.cells || []) if (typeof cell.v === 'string' && cell.v.trim()) labels.set(c, cell.v.trim());
  return { row: r, roles: {}, labels, score: 0 };
}

export function findHeader(rows) {
  return headerCandidates(rows)[0] || null;
}

/** Etiquetas de grupo encima del encabezado (p. ej. "CARTILLA PALMERI ET.1 Junio"). */
export function groupLabels(rows, merges, headerRow) {
  const out = new Map();
  const lookup = (r, c) => {
    const direct = rows.get(r)?.cells.get(c)?.v;
    if (typeof direct === 'string' && direct.trim()) return direct.trim();
    for (const m of merges) {
      if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) {
        const v = rows.get(m.r1)?.cells.get(m.c1)?.v;
        if (typeof v === 'string' && v.trim()) return v.trim();
      }
    }
    return null;
  };
  const cols = new Set();
  for (const [r, row] of rows) if (r >= headerRow - 4 && r <= headerRow) for (const c of row.cells.keys()) cols.add(c);
  for (const m of merges) if (m.r2 >= headerRow - 4 && m.r1 < headerRow) for (let c = m.c1; c <= m.c2; c++) cols.add(c);
  for (const c of cols) {
    for (let r = headerRow - 1; r >= Math.max(1, headerRow - 4); r--) {
      const g = lookup(r, c);
      if (g && /^(FECHA|PROYECTO|OBRA|ELABOR|CONSTRUCCIONES|PRESUPUESTO DETALLADO)/i.test(norm(g))) break;
      if (g) { out.set(c, g.replace(/\s+/g, ' ')); break; }
    }
  }
  return out;
}

// ---------------------------------------------------------------- detección de hoja
export async function detectSheets(book, { onProgress } = {}) {
  const res = [];
  let i = 0;
  for (const sh of book.sheets) {
    onProgress?.(++i, book.sheets.length, sh.name);
    let head;
    try { head = await readHead(book.zip, sh.path); } catch { continue; }
    const rows = parseSheetXml(head, book.sst, { maxRows: 80 });
    const h = findHeader(rows);
    if (!h) continue;
    const n = norm(sh.name);
    let score = h.score * 10;
    if (/DETALL/.test(n)) score += 15;
    if (/PPTO|PRESUP/.test(n)) score += 6;
    if (/\bIMP\b|IMPRES|RESUM|REPORTE|APU|INSUMO|MEMORIA|MEM\b|CUANT|COMPARA/.test(n)) score -= 8;
    if (sh.state !== 'visible') score -= 12;
    res.push({ sheet: sh, header: h, score });
  }
  res.sort((a, b) => b.score - a.score);
  // desempate con estructura real: contar actividades en los mejores candidatos
  for (const cand of res.slice(0, 3)) {
    try {
      const full = await loadSheet(book, cand.sheet);
      cand.full = full;
      const model = classify(full.rows, cand.header);
      cand.stats = { ues: model.ues.length, acts: model.acts.length };
      cand.score += Math.min(25, model.acts.length / 20) + Math.min(10, model.ues.length);
      if (model.ues.length === 0) cand.score -= 20;
    } catch (e) { cand.error = String(e); }
  }
  res.sort((a, b) => b.score - a.score);
  return res;
}

// ---------------------------------------------------------------- clasificación
const RE_SUB = /^SUB ?-? ?TOTAL(ES)?\b/;
const RE_GRAND = /^TOTAL (DEL |DE )?(PRESUPUESTO|GENERAL|PROYECTO|OBRA|COSTO TOTAL)\b|^GRAN TOTAL\b|^VALOR TOTAL (DEL )?(PRESUPUESTO|PROYECTO)\b/;
const RE_CD = /^SUB ?-? ?TOTAL (DE )?(LOS )?(COSTOS? DIRECTOS?|C\.? ?D\.?)(\b|$)/;

function looksLikeUnit(v) {
  if (typeof v !== 'string') return false;
  const t = v.trim();
  return t.length > 0 && t.length <= 14 && !/^SUB|TOTAL/i.test(t);
}
function codeLike(v) {
  if (v == null || isErr(v)) return false;
  if (typeof v === 'number') return true;
  const t = String(v).trim();
  return t.length > 0 && t.length <= 14 && !/\s/.test(t);
}
// código de actividad: 1.002, 3,011-1, 4.03200000000001 (no un entero de capítulo)
function decimalCode(v) {
  if (typeof v === 'number') return !Number.isInteger(v);
  return typeof v === 'string' && /^\s*\d{1,3}[.,]\d+/.test(v);
}
export function chapterNumber(code) {
  if (typeof code === 'number' && Number.isInteger(code)) return String(code);
  if (typeof code === 'string' && /^\s*\d{1,3}\s*$/.test(code)) return String(parseInt(code, 10));
  return null;
}
function similar(a, b) {
  const x = norm(a); const y = norm(b);
  if (!x || !y) return false;
  return x === y || (x.length >= 6 && y.startsWith(x)) || (y.length >= 6 && x.startsWith(y));
}

/**
 * Clasifica las filas a partir del encabezado.
 * Devuelve { ues, acts, chOcc, subs, indirect, cdRow, totalRow, alerts, ... }
 */
export function classify(rows, header, opts = {}) {
  const R = header.roles;
  const sumCols = opts.sumCols || [R.vt].filter(Boolean);
  const ues = []; const acts = []; const chOcc = []; const indirect = [];
  const alerts = [];
  let cur = { ue: null, ch: null, sub: null };
  let expectUE = true;
  let section = 'direct';
  let cdRow = null; let totalRow = null;
  const rowNums = [...rows.keys()].filter(r => r > header.row).sort((a, b) => a - b);

  const newUE = (r, code, name) => {
    const ue = { id: ues.length, code: fmtCode(code), rawCode: code, name: String(name ?? '').trim(), row: r, section, subtotalRow: null, acts: [] };
    ues.push(ue);
    cur = { ue, ch: null, sub: null };
    if (section === 'indirect') indirect.push({ kind: 'ue', row: r, ue });
    return ue;
  };

  for (let k = 0; k < rowNums.length; k++) {
    const r = rowNums[k];
    const row = rows.get(r);
    const item = val(row, R.item);
    const apu = val(row, R.apu);
    const desc = txt(row, R.desc);
    const und = val(row, R.und);
    const vt = val(row, R.vt);
    const itemTxt = txt(row, R.item);
    // la etiqueta puede venir en la columna de ítem (p. ej. "SUB-TOTAL CD")
    const label = desc || (!codeLike(item) ? itemTxt : '') || (R.apu && typeof apu === 'string' && !codeLike(apu) ? apu.trim() : '');
    const nl = norm(label);
    const hasAnyKey = label || itemTxt || txt(row, R.apu) || (und != null && und !== '') || vt != null;
    if (!hasAnyKey) continue;

    // ---------------- totales
    if (nl && (RE_SUB.test(nl) || /^TOTAL\b/.test(nl)) && !looksLikeUnit(und)) {
      if (RE_GRAND.test(nl) && !RE_SUB.test(nl)) { totalRow = r; break; }
      if (RE_CD.test(nl) && !/GG|GASTOS|ADMIN/.test(nl)) {
        cdRow = r; section = 'indirect'; expectUE = true;
        if (cur.ue && cur.ue.subtotalRow == null) cur.ue.subtotalRow = r;
        cur = { ue: null, ch: null, sub: null };
        continue;
      }
      const name = nl.replace(RE_SUB, '').replace(/^TOTAL\b/, '').trim();
      const isChapterSub = cur.ch && similar(name, cur.ch.name) && !(cur.ue && similar(name, cur.ue.name));
      if (isChapterSub) { cur.sub = null; continue; }
      if (cur.ue && cur.ue.subtotalRow == null) cur.ue.subtotalRow = r;
      if (section === 'indirect') indirect.push({ kind: 'subtotal', row: r });
      expectUE = true;
      continue;
    }

    // ---------------- actividad
    const byCode = !looksLikeUnit(und) && desc && decimalCode(item);
    const isAct = looksLikeUnit(und) || (isNum(val(row, R.cant)) && isNum(val(row, R.vu)) && (desc || itemTxt)) || byCode;
    if (isAct) {
      if (byCode) alerts.push({ type: 'act_sin_unidad', row: r, msg: `Fila ${r}: "${itemTxt} ${desc}" no tiene unidad; se toma como actividad por su código.` });
      if (!cur.ue) {
        newUE(r, '', section === 'direct' ? '(SIN UNIDAD ESTRUCTURAL)' : 'COSTOS INDIRECTOS');
        if (section === 'direct') alerts.push({ type: 'sin_ue', row: r, msg: `Fila ${r}: actividad antes de cualquier unidad estructural.` });
      }
      if (!cur.ch && section === 'direct') {
        const occ = { id: chOcc.length, ue: cur.ue, code: '', rawCode: null, name: '(SIN CAPÍTULO)', row: null, value: null, synthetic: true };
        chOcc.push(occ); cur.ch = occ;
        alerts.push({ type: 'sin_capitulo', row: r, msg: `Fila ${r}: actividad sin capítulo dentro de "${cur.ue.name}".` });
      }
      expectUE = false;
      const a = { row: r, ue: cur.ue, ch: cur.ch, sub: cur.sub, section, item: fmtCode(item), desc };
      for (const c of sumCols) { const v = val(row, c); if (isErr(v)) alerts.push({ type: 'error_valor', row: r, msg: `Fila ${r}: valor con error ${v.error} en la columna ${numToCol(c)}.` }); }
      acts.push(a); cur.ue.acts.push(a);
      if (section === 'indirect') indirect.push({ kind: 'act', row: r, act: a });
      continue;
    }

    // ---------------- filas de grupo (UE / capítulo / subgrupo)
    const looksUE = typeof item === 'string' && (/^0\d{1,2}$/.test(item.trim()) || /^[A-Z]{1,3}-\d{1,3}$/i.test(item.trim()));
    if (expectUE || (looksUE && cur.ch && chapterNumber(cur.ch.rawCode) != null)) {
      newUE(r, item, label || itemTxt);
      expectUE = false;
      continue;
    }
    if (codeLike(item) || isNum(vt)) {
      const occ = { id: chOcc.length, ue: cur.ue, code: fmtCode(item), rawCode: item, name: label, row: r, value: isNum(vt) ? vt : null, section };
      if (section === 'direct') chOcc.push(occ);
      else indirect.push({ kind: 'chapter', row: r, occ });
      cur.ch = occ; cur.sub = null;
      continue;
    }
    if (label) {
      cur.sub = label;
      if (section === 'indirect') indirect.push({ kind: 'sub', row: r, label });
    }
  }

  return { header, rows, ues, acts, chOcc, indirect, cdRow, totalRow, alerts, sumCols };
}

// ---------------------------------------------------------------- pivot CAP > UE > ACT
export function pivot(model) {
  const alerts = [...model.alerts];
  const directUEs = model.ues.filter(u => u.section === 'direct');
  // etiquetas únicas de UE
  const seen = new Map();
  for (const u of model.ues) {
    const base = [u.code, u.name].filter(Boolean).join(' - ') || `UE fila ${u.row}`;
    const n = (seen.get(base) || 0) + 1; seen.set(base, n);
    u.label = n > 1 ? `${base} (${n})` : base;
    if (n === 2) alerts.push({ type: 'ue_repetida', row: u.row, msg: `La unidad estructural "${base}" aparece más de una vez; se separa como "${u.label}".` });
  }
  const codeCount = new Map();
  for (const u of directUEs) if (u.code) codeCount.set(u.code, [...(codeCount.get(u.code) || []), u.name]);
  for (const [code, names] of codeCount) if (new Set(names.map(norm)).size > 1)
    alerts.push({ type: 'ue_codigo', msg: `El código de UE "${code}" se usa para unidades distintas: ${[...new Set(names)].join(' | ')}. Se identifican por código + nombre.` });

  // nombre -> número de capítulo (para capítulos sin código numérico)
  const byName = new Map();
  for (const o of model.chOcc) { const n = chapterNumber(o.rawCode); if (n && !byName.has(norm(o.name))) byName.set(norm(o.name), n); }
  const resolve = o => {
    const n = chapterNumber(o.rawCode);
    if (n) return n;
    if (o.synthetic) return 'SIN';
    const exact = byName.get(norm(o.name));
    let found = exact;
    if (!found) for (const [nm, num] of byName) if (similar(nm, o.name)) { found = num; break; }
    if (found) {
      alerts.push({ type: 'capitulo_por_nombre', row: o.row, msg: `Fila ${o.row}: capítulo "${[o.code, o.name].filter(Boolean).join(' ')}" (en ${o.ue?.name}) sin código numérico; se asigna al capítulo ${found} por nombre.` });
      return found;
    }
    alerts.push({ type: 'capitulo_sin_codigo', row: o.row, msg: `Fila ${o.row}: capítulo "${[o.code, o.name].filter(Boolean).join(' ')}" (en ${o.ue?.name}) sin código; queda como capítulo propio.` });
    return 'N:' + norm(o.name);
  };

  const chapters = new Map();
  for (const o of model.chOcc) {
    const key = resolve(o);
    o.key = key;
    if (!chapters.has(key)) chapters.set(key, { key, code: /^\d+$/.test(key) ? key : '', names: new Map(), occ: [], ues: new Map(), first: o.row ?? Infinity });
    const ch = chapters.get(key);
    ch.occ.push(o);
    ch.names.set(o.name, (ch.names.get(o.name) || 0) + 1);
  }
  for (const ch of chapters.values()) {
    const sorted = [...ch.names].sort((a, b) => b[1] - a[1]);
    ch.name = sorted[0]?.[0] || '';
    if (new Set(sorted.map(([n]) => norm(n))).size > 1)
      alerts.push({ type: 'conflicto_nombre', msg: `Capítulo ${ch.code || ch.key}: aparece con varios nombres (${sorted.map(([n, c]) => `${n} ×${c}`).join(' | ')}). Se unifican como "${ch.name}".` });
  }
  for (const a of model.acts) {
    if (a.section !== 'direct') continue;
    const ch = chapters.get(a.ch.key);
    if (!ch.ues.has(a.ue.id)) ch.ues.set(a.ue.id, { ue: a.ue, acts: [] });
    // subgrupo: etiqueta explícita o nombre original del capítulo si difiere del unificado
    a.group = a.sub || (norm(a.ch.name) !== norm(ch.name) && !a.ch.synthetic ? a.ch.name : null);
    ch.ues.get(a.ue.id).acts.push(a);
  }
  // capítulos sin actividades (vacíos) también se listan para no perder estructura
  const order = [...chapters.values()].sort((a, b) => {
    const na = /^\d+$/.test(a.key) ? +a.key : Infinity; const nb = /^\d+$/.test(b.key) ? +b.key : Infinity;
    return na - nb || a.first - b.first;
  });
  for (const ch of order) ch.ueList = [...ch.ues.values()].sort((x, y) => x.ue.id - y.ue.id);
  return { chapters: order, alerts };
}
