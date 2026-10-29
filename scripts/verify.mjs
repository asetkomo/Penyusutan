// Uji: parser + hitung harus identik dengan dpu_2024_NEW.xlsx (nilai cache Excel) untuk semua baris.
import fs from 'fs'; import ExcelJS from 'exceljs';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { extractRecords } from '../src/lib/parsePdf.js'; import { hitung } from '../src/lib/calc.js'; import { audit } from '../src/lib/audit.js';
const U = process.env.DATA_DIR || '/mnt/user-data/uploads/';
const pdf = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(U + 'dpu.pdf')), verbosity: 0 }).promise;
const pages = []; for (let p = 1; p <= pdf.numPages; p++) { const tc = await (await pdf.getPage(p)).getTextContent(); pages.push(tc.items.map((i) => ({ str: i.str, x: i.transform[4], y: i.transform[5] }))); }
const { rows, warnings } = extractRecords(pages);
const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(U + 'dpu_2024_NEW.xlsx');
const ws = wb.getWorksheet('Sheet'); const master = new Map();
wb.getWorksheet('1').eachRow((r, n) => { if (n > 1) master.set(String(r.getCell(2).value), Number(r.getCell(1).value)); });
const v = (c) => (c && typeof c === 'object' && 'result' in c ? c.result : c); const norm = (s) => String(s ?? '').replace(/\s+/g, ' ').trim();
const xl = []; ws.eachRow((r, n) => { if (n >= 4) xl.push(r); });
console.log('PDF rows', rows.length, '| Excel rows', xl.length, '| warnings', warnings.length);
let bad = 0; const cols = { nama: 10, nibar: 11, register: 12, spesNama: 13, spesLain: 14, merek: 15, lokasi: 16, satuan: 21, cara: 24, status: 26, ket: 27 };
const miss = (m) => { if (bad++ < 12) console.log('  BEDA', m); };
rows.forEach((r, i) => {
  const e = xl[i]; if (!e) return miss('baris Excel tidak ada ' + i);
  for (const [k, c] of Object.entries(cols)) if (r.type === 'item' || k === 'nama') if (norm(r[k]) !== norm(v(e.getCell(c).value))) miss(`#${i} ${k}: "${norm(r[k])}" vs "${norm(v(e.getCell(c).value))}"`);
  if (Math.abs((r.jumlah ?? 0) - (v(e.getCell(20).value) ?? 0)) > 1e-9) miss(`#${i} jumlah`);
  if (Math.abs((r.nilai ?? 0) - (v(e.getCell(23).value) ?? 0)) > 0.005) miss(`#${i} nilai ${r.nilai} vs ${v(e.getCell(23).value)}`);
  if (r.type === 'item') {
    const d = v(e.getCell(25).value); if (r.tanggal !== d.toISOString().slice(0, 10)) miss(`#${i} tanggal ${r.tanggal} vs ${d.toISOString()}`);
    if (Math.abs((r.harga ?? 0) - v(e.getCell(22).value)) > 0.005) miss(`#${i} harga`);
    const h = hitung(r, master.get(r.kode), { tahun: 2024 });
    const exp = { masa: 28, perBulan: 29, lalu: 30, akumLalu: 31, sisaLalu: 33, kini: 34, beban: 35, akum: 36, buku: 37, sisa: 38 };
    for (const [k, c] of Object.entries(exp)) if (Math.abs(h[k] - v(e.getCell(c).value)) > 0.01) miss(`#${i} hitung ${k}: ${h[k]} vs ${v(e.getCell(c).value)}`);
  }
});
const items = rows.filter((r) => r.type === 'item');
console.log('item', items.length, '| group', rows.length - items.length, '| selisih ditemukan', bad);
const hs = items.map((r) => hitung(r, master.get(r.kode), { tahun: 2024 }));
const f = audit(rows, hs, { tahun: 2024 }); const by = {}; f.forEach((x) => (by[x.level + ' ' + x.kode] = (by[x.level + ' ' + x.kode] || 0) + 1));
console.log('temuan audit', by); f.filter((x) => x.level !== 'INFO').slice(0, 8).forEach((x) => console.log('  ', x.level, x.kode, x.nibar.slice(-8), x.msg));
console.log('total nilai', items.reduce((a, r) => a + r.nilai, 0).toFixed(2), '| beban 2024', hs.reduce((a, h) => a + h.beban, 0).toFixed(2), '| akum', hs.reduce((a, h) => a + h.akum, 0).toFixed(2));
process.exit(bad ? 1 : 0);
