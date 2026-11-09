// Salida A: libro nuevo y liviano con ExcelJS (actividades como valores, sumas como fórmulas).
import { styleFor, styleKey } from './styles.js';
import { eachCell, colWidth } from './render.js';

function toExcelJsStyle(s) {
  const out = {};
  const f = s.font || {};
  out.font = { name: 'Calibri', size: f.sz || 10, bold: !!f.b, italic: !!f.i, color: { argb: 'FF' + (f.color || '000000') } };
  if (s.fill) out.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + s.fill } };
  if (s.border) {
    const c = { argb: 'FFA6A6A6' };
    if (s.border === 'thin') out.border = { top: { style: 'thin', color: c }, left: { style: 'thin', color: c }, bottom: { style: 'thin', color: c }, right: { style: 'thin', color: c } };
    else if (s.border === 'hair') out.border = { bottom: { style: 'hair', color: c } };
    else out.border = { top: { style: 'medium', color: c }, bottom: { style: 'medium', color: c } };
  }
  if (s.numFmt) out.numFmt = s.numFmt;
  const a = s.align || {};
  out.alignment = { vertical: a.v || 'middle', horizontal: a.h, wrapText: !!a.wrap, indent: a.indent || 0 };
  if (out.alignment.vertical === 'center') out.alignment.vertical = 'middle';
  return out;
}

export async function writeNew(layout, { ExcelJS, sourceName, project } = {}) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Reorganizador de presupuestos';
  wb.created = new Date();
  wb.calcProperties.fullCalcOnLoad = true;
  const cache = new Map();
  const st = (type, kind) => {
    const k = styleKey(type, kind);
    if (!cache.has(k)) cache.set(k, toExcelJsStyle(styleFor(type, kind)));
    return cache.get(k);
  };
  layout.sheets.forEach((sheet, i) => {
    const ws = wb.addWorksheet(sheet.name, {
      properties: { tabColor: { argb: 'FF1F3864' }, outlineProperties: { summaryBelow: false, summaryRight: false } },
      views: [{ state: 'frozen', ySplit: sheet.headerRow, xSplit: sheet.freezeCol || 0, showGridLines: false, zoomScale: 90 }],
    });
    ws.columns = sheet.cols.map(c => ({ key: c.key, width: colWidth(i, c) }));
    eachCell(sheet, i, { mode: 'new' }, (r, col, res, styleType, kind) => {
      const cell = ws.getCell(col.letter + r.n);
      if (res) {
        if (res.f) cell.value = { formula: res.f, result: res.v };
        else if (res.v != null && res.v !== '') cell.value = res.v;
      }
      cell.style = st(styleType, kind);
    });
    for (const r of sheet.rows) {
      const row = ws.getRow(r.n);
      if (r.level) row.outlineLevel = r.level;
      if (r.type === 'header') row.height = 32;
      else if (r.type === 'title') row.height = 21;
      else if (r.type === 'cap' || r.type === 'grand') row.height = 18;
    }
    if (i === 0) ws.autoFilter = { from: { row: sheet.headerRow, column: 1 }, to: { row: sheet.rows.length, column: sheet.cols.length } };
  });
  return wb.xlsx.writeBuffer();
}
