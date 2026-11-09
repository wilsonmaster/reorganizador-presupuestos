// Especificación única de estilos, traducida luego a ExcelJS o a styles.xml.
export const COLORS = {
  navy: '1F3864', navyText: 'FFFFFF', ue: 'D9E1F2', sub: 'F2F2F2', total: 'FFF2CC', grand: 'FFD966',
  text: '1F1F1F', muted: '595959', border: 'A6A6A6',
};

export const NUMFMT = {
  money: '#,##0', price: '#,##0', qty: '#,##0.00', pct: '0.00%', int: '0', diff: '#,##0.00',
};

const base = { font: { sz: 10, color: COLORS.text }, fill: null, border: null, numFmt: null, align: { v: 'center' } };

export function styleFor(type, kind) {
  const s = JSON.parse(JSON.stringify(base));
  const numeric = ['money', 'price', 'qty', 'pct', 'int', 'diff'].includes(kind);
  if (numeric) { s.numFmt = NUMFMT[kind]; s.align.h = 'right'; }
  if (kind === 'unit' || kind === 'center') s.align.h = 'center';
  if (kind === 'code') s.align.h = 'left';
  switch (type) {
    case 'title': s.font = { sz: 14, b: true, color: COLORS.navy }; s.align = { v: 'center' }; break;
    case 'subtitle': s.font = { sz: 10, i: true, color: COLORS.muted }; s.align = { v: 'center' }; break;
    case 'header': s.font = { sz: 10, b: true, color: COLORS.navyText }; s.fill = COLORS.navy; s.border = 'thin'; s.align = { h: 'center', v: 'center', wrap: true }; s.numFmt = null; break;
    case 'groupHeader': s.font = { sz: 9, b: true, color: COLORS.navy }; s.fill = COLORS.ue; s.border = 'thin'; s.align = { h: 'center', v: 'center', wrap: true }; s.numFmt = null; break;
    case 'cap': s.font = { sz: 11, b: true, color: COLORS.navyText }; s.fill = COLORS.navy; s.border = 'thin'; break;
    case 'ue': s.font = { sz: 10, b: true, color: COLORS.navy }; s.fill = COLORS.ue; s.border = 'thin'; if (kind === 'text') s.align.indent = 1; break;
    case 'sub': s.font = { sz: 10, i: true, b: true, color: COLORS.muted }; s.fill = COLORS.sub; if (kind === 'text') s.align.indent = 2; break;
    case 'act': s.border = 'hair'; if (kind === 'text') s.align.indent = 3; if (kind === 'text') s.align.wrap = false; break;
    case 'total': s.font = { sz: 11, b: true, color: COLORS.text }; s.fill = COLORS.total; s.border = 'medium'; break;
    case 'grand': s.font = { sz: 12, b: true, color: COLORS.text }; s.fill = COLORS.grand; s.border = 'medium'; break;
    case 'note': s.font = { sz: 10, color: COLORS.text }; s.align = { v: 'top', wrap: false }; break;
    case 'row': s.border = 'hair'; break;
    case 'ok': s.font = { sz: 10, b: true, color: COLORS.text }; s.align = { h: 'center', v: 'center' }; s.border = 'hair'; break;
    default: break;
  }
  return s;
}

export function styleKey(type, kind) { return type + '|' + kind; }
