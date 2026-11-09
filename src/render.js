// Resuelve cada celda abstracta del layout a {f: fórmula|null, v: valor} y su estilo.
import { formulaText, srcRef, quoteSheet } from './layout.js';
import { numToCol } from './xlsx-reader.js';

const FULL_ROW_TYPES = new Set(['header', 'groupHeader', 'cap', 'ue', 'sub', 'act', 'total', 'grand', 'row']);

export function cellKind(sheetIdx, col) {
  if (sheetIdx === 0) {
    if (col.key === 'origen') return 'int';
    if (col.key === 'ue' || col.key === 'cap' || col.key === 'grupo') return 'text';
    return col.kind === 'text' ? 'text' : col.kind;
  }
  if (sheetIdx === 1) {
    if (col.key === 'a') return 'text';
    if (col.key === 'est') return 'center';
    if (col.key.startsWith('d')) return 'diff';
    return 'money';
  }
  return col.key === 'a' ? 'text' : 'money';
}

/** Itera celdas a escribir: callback(rowObj, col, colIdx, resolved, styleType, kind) */
export function eachCell(sheet, sheetIdx, ctx, cb) {
  for (const r of sheet.rows) {
    const full = FULL_ROW_TYPES.has(r.type);
    for (const col of sheet.cols) {
      const cell = r.cells[col.key];
      if (!cell && !full) continue;
      const kind = cellKind(sheetIdx, col);
      let styleType = r.type;
      if (r.type === 'row' && col.key === 'est') styleType = 'ok';
      const resolved = cell ? resolve(cell, { ...ctx, col: col.letter, row: r.n }) : null;
      cb(r, col, resolved, styleType, kind, cell);
    }
  }
}

export function resolve(cell, ctx) {
  const { mode, sourceName } = ctx;
  if (cell.link) {
    if (mode === 'inject') return { f: srcRef(sourceName, cell.link.row, cell.link.col), v: cell.v };
    return { f: null, v: cell.text ?? cell.v };
  }
  if (cell.links) {
    if (mode === 'inject') {
      if (!cell.links.length) return { f: null, v: 0 };
      const refs = cell.links.map(l => srcRef(sourceName, l.row, l.col));
      return { f: refs.length > 20 ? `SUM(${refs.join(',')})` : refs.join('+'), v: cell.v };
    }
    return { f: null, v: cell.v };
  }
  if (cell.origin != null) {
    if (mode === 'inject') return { f: `HYPERLINK("#${quoteSheet(sourceName).replace(/"/g, '""')}!A${cell.origin}",${cell.origin})`, v: cell.origin };
    return { f: null, v: cell.origin };
  }
  if (cell.f) return { f: formulaText(cell.f, ctx), v: cell.v };
  return { f: null, v: cell.v };
}

export function colWidth(sheetIdx, col) {
  if (col.w) return col.w;
  if (sheetIdx !== 0) return 16;
  return { item: 11, apu: 10, desc: 58, und: 7, cant: 13, vu: 16, vt: 19, ue: 34, cap: 34, grupo: 26, origen: 9 }[col.key]
    || (col.kind === 'money' ? 19 : col.kind === 'qty' ? 13 : 16);
}

export function uniqueSheetNames(existing) {
  const lower = new Set(existing.map(n => n.toLowerCase()));
  return base => {
    let name = base; let i = 2;
    while (lower.has(name.toLowerCase())) name = `${base.slice(0, 27)} (${i++})`;
    lower.add(name.toLowerCase());
    return name;
  };
}

export { numToCol };
