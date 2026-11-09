// Fachada del motor: analizar → preparar → construir → exportar.
import { openWorkbook, loadSheet } from './xlsx-reader.js';
import { detectSheets, classify, pivot, groupLabels, findHeader, headerAtRow } from './classify.js';
import { availableColumns, buildColumns, buildLayout } from './layout.js';
import { writeNew } from './writer-new.js';
import { writeInject } from './writer-inject.js';

export async function analyze(data, { JSZip, onProgress } = {}) {
  const book = await openWorkbook(data, JSZip);
  const candidates = await detectSheets(book, { onProgress });
  return { book, candidates };
}

/** Carga la hoja elegida y propone columnas. roles: sobrescritura opcional {desc: 4, vt: 19, ...} */
export async function prepare(book, sheet, { roles, headerRow, full: cached } = {}) {
  const full = cached || await loadSheet(book, sheet);
  let header = headerRow ? headerAtRow(full.rows, headerRow) : findHeader(full.rows);
  if (!header) throw new Error(`No se encontró en "${sheet.name}" una fila de encabezado con DESCRIPCIÓN / UND / CANT / VR TOTAL.`);
  if (roles) header = { ...header, roles: { ...header.roles, ...Object.fromEntries(Object.entries(roles).filter(([, v]) => v !== undefined)) } };
  for (const k of Object.keys(header.roles)) if (!header.roles[k]) delete header.roles[k];
  if (!header.roles.desc || !header.roles.vt) throw new Error('Debe definirse al menos la columna de DESCRIPCIÓN y la de VR. TOTAL.');
  const groups = groupLabels(full.rows, full.merges, header.row);
  const available = availableColumns(header, groups);
  return { sheet, full, header, groups, available };
}

export function run(prep, { extras = [], fileName } = {}) {
  const cols = buildColumns(prep.header, prep.available, extras);
  const sumCols = cols.filter(c => c.sum && c.src).map(c => c.src);
  const model = classify(prep.full.rows, prep.header, { sumCols });
  const pv = pivot(model);
  const layout = buildLayout(model, pv, cols, { sourceName: prep.sheet.name, fileName });
  return { model, pv, cols, layout };
}

export async function exportNew(result, { ExcelJS }) {
  return writeNew(result.layout, { ExcelJS });
}

/** Re-abre el libro original (para no mutar el que está en memoria) e inyecta las hojas. */
export async function exportInject(originalData, result, sourceName, { JSZip, keepSourceStyle = true } = {}) {
  const book = await openWorkbook(originalData, JSZip);
  const { zip, names } = await writeInject(book, result.layout, { sourceName, keepSourceStyle });
  const out = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  return { data: out, names };
}
