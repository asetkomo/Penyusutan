import fs from 'fs'; import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { extractRecords } from '../src/lib/parsePdf.js'; import { hitung, cariMasaLama } from '../src/lib/calc.js'; import { audit } from '../src/lib/audit.js'; import { buildWorkbook } from '../src/lib/exportXlsx.js';
const U = process.env.DATA_DIR || '/mnt/user-data/uploads/';
const pdf = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(U + 'dpu.pdf')), verbosity: 0 }).promise;
const pages = []; for (let p = 1; p <= pdf.numPages; p++) { const tc = await (await pdf.getPage(p)).getTextContent(); pages.push(tc.items.map((i) => ({ str: i.str, x: i.transform[4], y: i.transform[5] }))); }
const { rows } = extractRecords(pages); const masa = JSON.parse(fs.readFileSync('public/masa_manfaat.json'));
const rules = [{ kode: '1.3.2.05.02.01', lama: 120 }]; const items = rows.filter((r) => r.type === 'item');
const hasil = items.map((r) => hitung(r, masa[r.kode], { tahun: 2024, masaLama: cariMasaLama(r.kode, rules) }));
const base = items.map((r) => hitung(r, masa[r.kode], { tahun: 2024 }));
let bad = 0;
hasil.forEach((h, i) => { const b = base[i];
  if (['beban', 'akum', 'buku', 'akumLalu'].some((k) => Math.abs(h[k] - b[k]) > 1e-6)) bad++; // beban 2024 tidak boleh berubah oleh koreksi
  if (h.masaLama) { if (Math.abs(h.akumLalu + h.beban - h.akum) > 0.01) bad++; if (h.koreksi < -0.001) bad++; if (!items[i].kode.startsWith('1.3.2.05.02.01')) bad++; }
  else if (h.koreksi !== undefined) bad++; });
const k = hasil.filter((h) => h.masaLama); const S = (f) => k.reduce((a, h) => a + f(h), 0);
console.log('aset terdampak', k.length, '| nilai', items.filter((_, i) => hasil[i].masaLama).reduce((a, r) => a + r.nilai, 0).toFixed(2));
console.log('akum awal lama', S((h) => h.akumLaluLama).toFixed(2), '| baru', S((h) => h.akumLalu).toFixed(2), '| KOREKSI', S((h) => h.koreksi).toFixed(2), '| beban 2024 (60 bln)', S((h) => h.beban).toFixed(2));
console.log('aset dgn koreksi > 0:', k.filter((h) => h.koreksi > 0.005).length, '| koreksi 0 (sudah habis di kedua skema / baru 2024):', k.filter((h) => h.koreksi <= 0.005).length, '| pelanggaran invarian', bad);
const top = items.map((r, i) => ({ r, h: hasil[i] })).filter((x) => x.h.koreksi > 0).sort((a, b) => b.h.koreksi - a.h.koreksi).slice(0, 5);
top.forEach((x) => console.log('  top', x.r.kode, x.r.tanggal, x.r.nilai, '->', x.h.koreksi.toFixed(2)));
const wb = await buildWorkbook({ rows, hasil, temuan: audit(rows, hasil, { tahun: 2024 }), masa, tahun: 2024, rules }); await wb.xlsx.writeFile('/tmp/hasil_koreksi.xlsx');
process.exit(bad ? 1 : 0);
