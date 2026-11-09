// Lector .xlsx liviano sobre JSZip: lee nombres de hojas, cadenas compartidas,
// valores cacheados, índices de estilo, combinaciones y anchos de columna.
// No recalcula nada: usa el valor que Excel dejó guardado en cada celda.

const NS = '(?:[A-Za-z_][\\w.-]*:)?'; // prefijo de namespace opcional (x:, etc.)

export function decodeXml(s) {
  if (!s || s.indexOf('&') < 0) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (m, e) => {
    if (e === 'amp') return '&';
    if (e === 'lt') return '<';
    if (e === 'gt') return '>';
    if (e === 'quot') return '"';
    if (e === 'apos') return "'";
    if (e[1] === 'x' || e[1] === 'X') return String.fromCodePoint(parseInt(e.slice(2), 16));
    return String.fromCodePoint(parseInt(e.slice(1), 10));
  });
}

export function escapeXml(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    // caracteres de control no válidos en XML 1.0
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
}

function attrs(tag) {
  const out = {};
  const re = /([\w:.-]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  let m;
  while ((m = re.exec(tag))) out[m[1].replace(/^.*:/, m[1].startsWith('r:') ? 'r:' : '')] = decodeXml(m[3] ?? m[4]);
  // conservar r:id tal cual
  const rid = /\br:id\s*=\s*"([^"]*)"/.exec(tag);
  if (rid) out['r:id'] = rid[1];
  return out;
}

export function colToNum(col) {
  let n = 0;
  for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
export function numToCol(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
export function splitRef(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  return m ? { col: colToNum(m[1]), row: parseInt(m[2], 10) } : null;
}

function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/');
  parts.pop();
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

function parseSharedStrings(xml) {
  const out = [];
  const siRe = new RegExp(`<${NS}si\\b[^>]*>([\\s\\S]*?)</${NS}si>|<${NS}si\\b[^>]*/>`, 'g');
  const tRe = new RegExp(`<${NS}t\\b[^>]*>([\\s\\S]*?)</${NS}t>`, 'g');
  const rphRe = new RegExp(`<${NS}rPh\\b[\\s\\S]*?</${NS}rPh>`, 'g');
  let m;
  while ((m = siRe.exec(xml))) {
    const body = (m[1] || '').replace(rphRe, '');
    let s = '';
    let t;
    tRe.lastIndex = 0;
    while ((t = tRe.exec(body))) s += t[1];
    out.push(decodeXml(s));
  }
  return out;
}

/** Abre el libro: devuelve { zip, sheets:[{name,path,state,index}], sst, ... } */
export async function openWorkbook(data, JSZip) {
  const zip = await JSZip.loadAsync(data);
  const wbPath = 'xl/workbook.xml';
  const wbXml = await zip.file(wbPath)?.async('string');
  if (!wbXml) throw new Error('El archivo no parece un libro .xlsx/.xlsm válido (falta xl/workbook.xml).');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const rels = {};
  for (const m of relsXml.matchAll(new RegExp(`<${NS}Relationship\\b[^>]*>`, 'g'))) {
    const a = attrs(m[0]);
    rels[a.Id] = { target: a.Target, type: a.Type };
  }
  const sheets = [];
  for (const m of wbXml.matchAll(new RegExp(`<${NS}sheet\\b[^>]*>`, 'g'))) {
    const a = attrs(m[0]);
    const rel = rels[a['r:id']];
    if (!rel || !/worksheet$/.test(rel.type)) continue; // hojas de gráfico, etc.
    sheets.push({ name: a.name, sheetId: +a.sheetId, rid: a['r:id'], state: a.state || 'visible',
      path: resolvePath(wbPath, rel.target), index: sheets.length });
  }
  const sstFile = Object.values(rels).find(r => /sharedStrings$/.test(r.type));
  const sstXml = sstFile ? await zip.file(resolvePath(wbPath, sstFile.target))?.async('string') : null;
  const sst = sstXml ? parseSharedStrings(sstXml) : [];
  return { zip, wbXml, relsXml, rels, sheets, sst };
}

/** Lee solo el comienzo del XML de una hoja (para detectar sin cargar hojas gigantes). */
export function readHead(zip, path, maxChars = 600000) {
  const file = zip.file(path);
  if (!file) return Promise.resolve('');
  return new Promise((resolve, reject) => {
    let buf = '';
    let done = false;
    const stream = file.internalStream('string');
    stream.on('data', chunk => {
      if (done) return;
      buf += chunk;
      if (buf.length >= maxChars) { done = true; stream.pause(); resolve(buf); }
    });
    stream.on('error', e => { if (!done) { done = true; reject(e); } });
    stream.on('end', () => { if (!done) { done = true; resolve(buf); } });
    stream.resume();
  });
}

/**
 * Convierte el XML de una hoja en filas.
 * rows: Map<rowNum, {attrs, cells: Map<colNum, cell>}>; cell = {v, t, s, f, ref}
 * v: number | string | boolean | {error} | null
 */
export function parseSheetXml(xml, sst, { maxRows = Infinity } = {}) {
  const rows = new Map();
  const rowRe = new RegExp(`<${NS}row\\b([^>]*?)(/>|>([\\s\\S]*?)</${NS}row>)`, 'g');
  const cRe = new RegExp(`<${NS}c\\b([^>]*?)(/>|>([\\s\\S]*?)</${NS}c>)`, 'g');
  const vRe = new RegExp(`<${NS}v\\b[^>]*>([\\s\\S]*?)</${NS}v>`);
  const fRe = new RegExp(`<${NS}f\\b[^>]*?(?:/>|>([\\s\\S]*?)</${NS}f>)`);
  const isRe = new RegExp(`<${NS}t\\b[^>]*>([\\s\\S]*?)</${NS}t>`, 'g');
  let m;
  let count = 0;
  let lastRow = 0;
  while ((m = rowRe.exec(xml))) {
    const ra = attrs(m[1]);
    const rnum = ra.r ? parseInt(ra.r, 10) : lastRow + 1;
    lastRow = rnum;
    if (rnum > maxRows) break;
    const cells = new Map();
    let lastCol = 0;
    if (m[3]) {
      let c;
      cRe.lastIndex = 0;
      while ((c = cRe.exec(m[3]))) {
        const ca = attrs(c[1]);
        const pos = ca.r ? splitRef(ca.r) : { col: lastCol + 1, row: rnum };
        lastCol = pos.col;
        const body = c[3] || '';
        const t = ca.t || 'n';
        let v = null;
        const vm = vRe.exec(body);
        const raw = vm ? decodeXml(vm[1]) : null;
        if (t === 's') v = raw != null ? (sst[parseInt(raw, 10)] ?? '') : null;
        else if (t === 'inlineStr') { let s = ''; let tm; isRe.lastIndex = 0; while ((tm = isRe.exec(body))) s += tm[1]; v = decodeXml(s); }
        else if (t === 'str') v = raw ?? '';
        else if (t === 'b') v = raw === '1';
        else if (t === 'e') v = raw != null ? { error: raw } : null;
        else if (raw != null && raw !== '') v = Number(raw);
        const fm = fRe.exec(body);
        cells.set(pos.col, { v, t, s: ca.s != null ? parseInt(ca.s, 10) : 0, f: fm ? decodeXml(fm[1] || '') : null, col: pos.col });
      }
    }
    rows.set(rnum, { attrs: ra, cells });
    if (++count >= maxRows) break;
  }
  return rows;
}

export function parseMerges(xml) {
  const out = [];
  for (const m of xml.matchAll(new RegExp(`<${NS}mergeCell\\b[^>]*ref="([A-Z]+\\d+):([A-Z]+\\d+)"`, 'g'))) {
    const a = splitRef(m[1]); const b = splitRef(m[2]);
    if (a && b) out.push({ r1: a.row, c1: a.col, r2: b.row, c2: b.col });
  }
  return out;
}

export function parseCols(xml) {
  const out = [];
  const head = xml.slice(0, Math.min(xml.length, 200000));
  for (const m of head.matchAll(new RegExp(`<${NS}col\\b[^>]*>`, 'g'))) {
    const a = attrs(m[0]);
    out.push({ min: +a.min, max: +a.max, width: a.width != null ? +a.width : null, style: a.style != null ? +a.style : null, hidden: a.hidden === '1' || a.hidden === 'true' });
  }
  return out;
}

export async function loadSheet(book, sheet) {
  const xml = await book.zip.file(sheet.path).async('string');
  return { xml, rows: parseSheetXml(xml, book.sst), merges: parseMerges(xml), cols: parseCols(xml) };
}
