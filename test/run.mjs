// Prueba local: node test/run.mjs <archivo.xlsx> [carpeta_salida]
import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { analyze, prepare, run, exportNew, exportInject } from '../src/engine.js';

const file = process.argv[2];
const outDir = process.argv[3] || 'test/out';
fs.mkdirSync(outDir, { recursive: true });
const data = fs.readFileSync(file);
const base = path.basename(file).replace(/\.(xlsx|xlsm)$/i, '').replace(/^[0-9a-f]{8}-/, '');
const ext = path.extname(file);
const t0 = Date.now();
const { book, candidates } = await analyze(data, { JSZip });
const best = candidates[0];
const prep = await prepare(book, best.sheet);
// incluir todas las columnas "VR. TOTAL" adicionales para probar varias columnas sumadas
const extras = prep.available.filter(a => !a.role && a.sumDefault && /TOTAL/i.test(a.label)).map(a => ({ ...a, sum: true }));
const res = run(prep, { extras, fileName: path.basename(file) });
const s = res.layout.summary;
console.log(`${base}: hoja="${best.sheet.name}" enc=${prep.header.row} UE=${s.ues} act=${s.acts} cap=${s.chapters} extras=${extras.map(e => e.letter).join(',')}`);
console.log(`  CD  orig=${s.cd.orig.toFixed(2)} reorg=${s.cd.reorg.toFixed(2)} dif=${(s.cd.reorg - s.cd.orig).toFixed(4)}`);
console.log(`  TOT orig=${s.total.orig.toFixed(2)} reorg=${s.total.reorg.toFixed(2)} dif=${(s.total.reorg - s.total.orig).toFixed(4)}  cierra=${s.closesAll}`);
const bad = s.ueChecks.filter(u => u.ok === false);
if (bad.length) console.log('  UE sin cuadre:', bad.map(b => b.ue.label));
const newBuf = await exportNew(res, { ExcelJS });
fs.writeFileSync(path.join(outDir, `${base}_NUEVO.xlsx`), Buffer.from(newBuf));
const inj = await exportInject(data, res, best.sheet.name, { JSZip });
fs.writeFileSync(path.join(outDir, `${base}_VINCULADO${ext}`), Buffer.from(inj.data));
console.log(`  ok en ${Date.now() - t0} ms · alertas=${s.alerts.length}`);
