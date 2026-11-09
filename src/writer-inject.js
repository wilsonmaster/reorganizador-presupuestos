// Salida B: agrega las hojas nuevas DENTRO del libro original, sin tocar las demás partes.
// Las actividades quedan como fórmulas vinculadas a la hoja fuente (='Ppto Detallado'!I33).
import { escapeXml } from './xlsx-reader.js';
import { styleFor, styleKey } from './styles.js';
import { eachCell, colWidth, uniqueSheetNames } from './render.js';
import { SHEET_MAIN, SHEET_CHECK, SHEET_MATRIX } from './layout.js';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WS_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const WS_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';

const BUILTIN_FMT = { '#,##0': 3, '#,##0.00': 4, '0': 1, '0.00%': 10 };

// ---------------------------------------------------------------- styles.xml
class StyleAppender {
  constructor(xml) {
    this.xml = xml;
    this.ok = /<styleSheet\b/.test(xml) && /<fonts\b[\s\S]*?<\/fonts>/.test(xml) && /<fills\b[\s\S]*?<\/fills>/.test(xml)
      && /<borders\b[\s\S]*?<\/borders>/.test(xml) && /<cellXfs\b[\s\S]*?<\/cellXfs>/.test(xml);
    if (!this.ok) return;
    const sec = name => new RegExp(`<${name}\\b[^>]*>([\\s\\S]*?)</${name}>`).exec(xml)[1];
    this.counts = {
      fonts: (sec('fonts').match(/<font[\s>/]/g) || []).length,
      fills: (sec('fills').match(/<fill[\s>/]/g) || []).length,
      borders: (sec('borders').match(/<border[\s>/]/g) || []).length,
      cellXfs: (sec('cellXfs').match(/<xf[\s>/]/g) || []).length,
    };
    const fn = /<name\s+val="([^"]+)"/.exec(sec('fonts'));
    this.fontName = fn ? fn[1] : 'Calibri';
    this.add = { fonts: [], fills: [], borders: [], cellXfs: [] };
    this.cache = new Map();
    this.sub = { font: new Map(), fill: new Map(), border: new Map() };
  }
  _idx(kind, mapKey, xmlFrag) {
    const m = this.sub[kind];
    if (m.has(mapKey)) return m.get(mapKey);
    const list = this.add[kind + 's'];
    const idx = this.counts[kind + 's'] + list.length;
    list.push(xmlFrag); m.set(mapKey, idx);
    return idx;
  }
  index(type, kind) {
    if (!this.ok) return 0;
    const key = styleKey(type, kind);
    if (this.cache.has(key)) return this.cache.get(key);
    const s = styleFor(type, kind);
    const f = s.font || {};
    const fontXml = `<font>${f.b ? '<b/>' : ''}${f.i ? '<i/>' : ''}<sz val="${f.sz || 10}"/><color rgb="FF${f.color || '000000'}"/><name val="${escapeXml(this.fontName)}"/><family val="2"/></font>`;
    const fontId = this._idx('font', fontXml, fontXml);
    let fillId = 0;
    if (s.fill) { const x = `<fill><patternFill patternType="solid"><fgColor rgb="FF${s.fill}"/><bgColor indexed="64"/></patternFill></fill>`; fillId = this._idx('fill', x, x); }
    let borderId = 0;
    if (s.border) {
      const side = (n, st) => (st ? `<${n} style="${st}"><color rgb="FFA6A6A6"/></${n}>` : `<${n}/>`);
      const b = s.border;
      const x = b === 'thin' ? `<border>${side('left', 'thin')}${side('right', 'thin')}${side('top', 'thin')}${side('bottom', 'thin')}<diagonal/></border>`
        : b === 'hair' ? `<border>${side('left')}${side('right')}${side('top')}${side('bottom', 'hair')}<diagonal/></border>`
          : `<border>${side('left')}${side('right')}${side('top', 'medium')}${side('bottom', 'medium')}<diagonal/></border>`;
      borderId = this._idx('border', x, x);
    }
    const numFmtId = s.numFmt ? (BUILTIN_FMT[s.numFmt] ?? 0) : 0;
    const a = s.align || {};
    const align = `<alignment${a.h ? ` horizontal="${a.h}"` : ''}${a.v ? ` vertical="${a.v}"` : ''}${a.wrap ? ' wrapText="1"' : ''}${a.indent ? ` indent="${a.indent}"` : ''}/>`;
    const xf = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">${align}</xf>`;
    const idx = this.counts.cellXfs + this.add.cellXfs.length;
    this.add.cellXfs.push(xf);
    this.cache.set(key, idx);
    return idx;
  }
  result() {
    if (!this.ok) return this.xml;
    let xml = this.xml;
    for (const name of ['fonts', 'fills', 'borders', 'cellXfs']) {
      const extra = this.add[name];
      if (!extra.length) continue;
      const total = this.counts[name] + extra.length;
      xml = xml.replace(new RegExp(`<${name}\\b([^>]*)>([\\s\\S]*?)</${name}>`), (m, at, body) => {
        const at2 = /\bcount="/.test(at) ? at.replace(/\bcount="\d*"/, `count="${total}"`) : `${at} count="${total}"`;
        return `<${name}${at2}>${body}${extra.join('')}</${name}>`;
      });
    }
    return xml;
  }
}

// ---------------------------------------------------------------- hoja
function cellXml(ref, s, res) {
  const sAttr = ` s="${s}"`;
  if (!res) return `<c r="${ref}"${sAttr}/>`;
  const { f, v } = res;
  if (f != null) {
    const fx = `<f>${escapeXml(f)}</f>`;
    if (typeof v === 'string') return `<c r="${ref}"${sAttr} t="str">${fx}<v>${escapeXml(v)}</v></c>`;
    if (typeof v === 'boolean') return `<c r="${ref}"${sAttr} t="b">${fx}<v>${v ? 1 : 0}</v></c>`;
    if (typeof v === 'number' && Number.isFinite(v)) return `<c r="${ref}"${sAttr}>${fx}<v>${v}</v></c>`;
    return `<c r="${ref}"${sAttr}>${fx}</c>`;
  }
  if (v == null || v === '') return `<c r="${ref}"${sAttr}/>`;
  if (typeof v === 'number') return Number.isFinite(v) ? `<c r="${ref}"${sAttr}><v>${v}</v></c>` : `<c r="${ref}"${sAttr}/>`;
  if (typeof v === 'boolean') return `<c r="${ref}"${sAttr} t="b"><v>${v ? 1 : 0}</v></c>`;
  return `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${escapeXml(v)}</t></is></c>`;
}

function sheetXml(sheet, sheetIdx, ctx, styles) {
  const byRow = new Map();
  eachCell(sheet, sheetIdx, ctx, (r, col, res, styleType, kind, cell) => {
    let s;
    // en las actividades se conserva el estilo exacto de la celda origen
    if (ctx.keepSourceStyle && r.type === 'act' && cell?.link && cell.s != null) s = cell.s;
    else s = styles.index(styleType, kind);
    if (!byRow.has(r.n)) byRow.set(r.n, []);
    byRow.get(r.n).push(cellXml(col.letter + r.n, s, res));
  });
  const lastCol = sheet.cols[sheet.cols.length - 1].letter;
  const lastRow = sheet.rows.length;
  const rowsXml = sheet.rows.map(r => {
    const attrs = [`r="${r.n}"`];
    if (r.type === 'header') attrs.push('ht="32" customHeight="1"');
    else if (r.type === 'title') attrs.push('ht="21" customHeight="1"');
    else if (r.type === 'cap' || r.type === 'grand') attrs.push('ht="18" customHeight="1"');
    if (r.level) attrs.push(`outlineLevel="${r.level}"`);
    return `<row ${attrs.join(' ')}>${(byRow.get(r.n) || []).join('')}</row>`;
  }).join('');
  const cols = sheet.cols.map((c, i) => `<col min="${i + 1}" max="${i + 1}" width="${colWidth(sheetIdx, c)}" customWidth="1"/>`).join('');
  const h = sheet.headerRow;
  const xs = sheet.freezeCol || 0;
  const topLeft = `${xs ? 'B' : 'A'}${h + 1}`;
  const pane = xs
    ? `<pane xSplit="${xs}" ySplit="${h}" topLeftCell="${topLeft}" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="${topLeft}" sqref="${topLeft}"/>`
    : `<pane ySplit="${h}" topLeftCell="${topLeft}" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="${topLeft}" sqref="${topLeft}"/>`;
  const maxLevel = Math.max(0, ...sheet.rows.map(r => r.level || 0));
  const autoFilter = sheetIdx === 0 ? `<autoFilter ref="A${h}:${lastCol}${lastRow}"/>` : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
    + `<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">`
    + `<sheetPr><tabColor rgb="FF1F3864"/><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr>`
    + `<dimension ref="A1:${lastCol}${lastRow}"/>`
    + `<sheetViews><sheetView showGridLines="0" zoomScale="90" zoomScaleNormal="90" workbookViewId="0">${pane}</sheetView></sheetViews>`
    + `<sheetFormatPr defaultRowHeight="15"${maxLevel ? ` outlineLevelRow="${maxLevel}"` : ''}/>`
    + `<cols>${cols}</cols>`
    + `<sheetData>${rowsXml}</sheetData>`
    + autoFilter
    + `<pageMargins left="0.4" right="0.4" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>`
    + `</worksheet>`;
}

// ---------------------------------------------------------------- libro
export async function writeInject(book, layout, { sourceName, keepSourceStyle = true, JSZip } = {}) {
  const zip = book.zip;
  const uniq = uniqueSheetNames(book.sheets.map(s => s.name));
  const names = { [SHEET_MAIN]: uniq(SHEET_MAIN), [SHEET_CHECK]: uniq(SHEET_CHECK), [SHEET_MATRIX]: uniq(SHEET_MATRIX) };
  const stylesPath = 'xl/styles.xml';
  const stylesXml = (await zip.file(stylesPath)?.async('string')) || '';
  const styles = new StyleAppender(stylesXml);
  const ctx = { mode: 'inject', sourceName, names, keepSourceStyle };

  let wb = book.wbXml;
  let rels = book.relsXml;
  let ct = await zip.file('[Content_Types].xml').async('string');
  // prefijo usado para r:id en las hojas existentes
  const idAttr = (/<sheet\b[^>]*?\s([\w]+:id)=/.exec(wb) || [])[1] || 'r:id';
  if (idAttr === 'r:id' && !/xmlns:r=/.test(wb)) wb = wb.replace(/<workbook\b/, `<workbook xmlns:r="${REL_NS}"`);
  let maxSheetId = Math.max(0, ...book.sheets.map(s => s.sheetId || 0), ...[...wb.matchAll(/sheetId="(\d+)"/g)].map(m => +m[1]));
  const existingRids = new Set([...rels.matchAll(/Id="([^"]+)"/g)].map(m => m[1]));
  let n = 1;
  const newSheets = [];
  layout.sheets.forEach((sheet, i) => {
    let file; do { file = `xl/worksheets/sheetReorg${n++}.xml`; } while (zip.file(file));
    let rid; let k = 1; do { rid = `rIdReorg${k++}`; } while (existingRids.has(rid));
    existingRids.add(rid);
    const name = names[sheet.name] || sheet.name;
    zip.file(file, sheetXml(sheet, i, ctx, styles));
    newSheets.push({ name, rid, file, sheetId: ++maxSheetId });
  });
  zip.file(stylesPath, styles.result());

  const sheetsBefore = (wb.match(/<(?:\w+:)?sheet\s/g) || []).length;
  const sheetTags = newSheets.map(s => `<sheet name="${escapeXml(s.name)}" sheetId="${s.sheetId}" ${idAttr}="${s.rid}"/>`).join('');
  wb = wb.replace(/<\/((?:\w+:)?)sheets>/, (m, p) => `${sheetTags.replace(/<sheet /g, `<${p}sheet `)}</${p}sheets>`);
  // nombre oculto del autofiltro de la hoja principal (como lo escribe Excel)
  const main = layout.sheets[0];
  const lastCol = main.cols[main.cols.length - 1].letter;
  const qn = `'${newSheets[0].name.replace(/'/g, "''")}'`;
  const dn = `<definedName name="_xlnm._FilterDatabase" localSheetId="${sheetsBefore}" hidden="1">${escapeXml(`${qn}!$A$${main.headerRow}:$${lastCol}$${main.rows.length}`)}</definedName>`;
  if (/<\/(?:\w+:)?definedNames>/.test(wb)) wb = wb.replace(/<\/((?:\w+:)?)definedNames>/, (m, p) => `${dn.replace(/definedName /, `${p}definedName `).replace('</definedName>', `</${p}definedName>`)}</${p}definedNames>`);
  else {
    const anchor = ['</externalReferences>', '</functionGroups>', '</sheets>'].find(t => wb.includes(t));
    if (anchor) wb = wb.replace(anchor, `${anchor}<definedNames>${dn}</definedNames>`);
  }
  // recálculo completo al abrir, para que los vínculos queden vivos
  if (/<(?:\w+:)?calcPr\b/.test(wb)) {
    wb = wb.replace(/<((?:\w+:)?)calcPr\b([^>]*?)(\/?)>/, (m, p, at, sl) => `<${p}calcPr${at.replace(/\sfullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"${sl}>`);
  } else {
    const anchors = ['</definedNames>', '</externalReferences>', '</functionGroups>', '</sheets>'];
    const a = anchors.find(t => wb.includes(t));
    if (a) wb = wb.replace(a, `${a}<calcPr fullCalcOnLoad="1"/>`);
  }
  zip.file('xl/workbook.xml', wb);

  const relTags = newSheets.map(s => `<Relationship Id="${s.rid}" Type="${WS_TYPE}" Target="${s.file.replace(/^xl\//, '')}"/>`).join('');
  rels = rels.replace(/<\/((?:\w+:)?)Relationships>/, (m, p) => `${relTags.replace(/<Relationship /g, `<${p}Relationship `)}</${p}Relationships>`);
  zip.file('xl/_rels/workbook.xml.rels', rels);

  const ctTags = newSheets.map(s => `<Override PartName="/${s.file}" ContentType="${WS_CT}"/>`).join('');
  ct = ct.replace(/<\/((?:\w+:)?)Types>/, (m, p) => `${ctTags.replace(/<Override /g, `<${p}Override `)}</${p}Types>`);
  zip.file('[Content_Types].xml', ct);

  // los valores cacheados cambian: el calcChain viejo sigue siendo válido (no se borran celdas)
  return { zip, names, newSheets, stylesOk: styles.ok };
}
