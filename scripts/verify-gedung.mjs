// Uji: parser PDF + template + hitung gedung harus identik dengan diskop_penyusutan_2026.xlsx (nilai cache Excel).
import fs from 'fs'; import ExcelJS from 'exceljs'; import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { extractRecords, detectFormat, extractMeta } from '../src/lib/parsePdf.js'; import { bacaTemplate } from '../src/lib/templateGedung.js'; import { hitungGedung } from '../src/lib/calcGedung.js';
const U = process.env.DATA_DIR || '/mnt/user-data/uploads/';
const pdf = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(U + 'gedung_diskop.pdf')), verbosity: 0 }).promise;
const pages = []; for (let p = 1; p <= pdf.numPages; p++) { const tc = await (await pdf.getPage(p)).getTextContent(); pages.push(tc.items.map((i) => ({ str: i.str, x: i.transform[4], y: i.transform[5] }))); }
const fmt = detectFormat(pages); const { rows, warnings } = extractRecords(pages, fmt); const meta = extractMeta(pages);
const tpl = await bacaTemplate(fs.readFileSync(U + 'Format_penyusutan_gedung.xlsx')); const masa = JSON.parse(fs.readFileSync('public/masa_manfaat.json'));
const items = rows.filter((r) => r.type === 'item');
console.log('format', fmt, '| meta', JSON.stringify(meta), '| item', items.length, 'group', rows.length - items.length, '| template', tpl.rows.length, 'tahunLalu', tpl.tahunLalu, '| warn', warnings.length + tpl.warnings.length);
const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(U + 'diskop_penyusutan_2026__1_.xlsx'); const ws = wb.getWorksheet('Sheet');
const v = (c) => (c && typeof c === 'object' && !(c instanceof Date) ? ('result' in c ? c.result : '') : c); const norm = (s) => String(s ?? '').replace(/\s+/g, ' ').trim();
const byN = new Map(items.map((r) => [r.nibar, r])); const tN = new Map(tpl.rows.map((r) => [r.nibar, r]));
let bad = 0, cmp = 0; const miss = (m) => { if (bad++ < 40) console.log('  BEDA', m); };
for (let r = 3; r <= 70; r++) { const row = ws.getRow(r); const nib = norm(v(row.getCell(11).value)); const it = byN.get(nib), tp = tN.get(nib);
  if (!it) { miss(`r${r} NIBAR ${nib} tidak ada di PDF`); continue; } if (!tp) { miss(`r${r} tidak ada di template`); continue; }
  const T = { 7: it.seg[6], 10: it.nama, 12: it.register, 13: it.spesNama, 14: it.spesLain, 16: it.lokasi, 17: it.koordinat, 18: it.statusTanah, 20: it.satuan, 23: it.cara, 25: it.status, 26: it.ket };
  for (const [c, x] of Object.entries(T)) { cmp++; if (norm(x) !== norm(v(row.getCell(+c).value))) miss(`r${r} kol${c}: "${norm(x)}" vs "${norm(v(row.getCell(+c).value))}"`); }
  const N = { 15: it.jmlLantai, 19: it.jumlah, 21: it.harga, 22: it.nilai }; for (const [c, x] of Object.entries(N)) { cmp++; if (Math.abs((x ?? 0) - (v(row.getCell(+c).value) ?? 0)) > 0.005) miss(`r${r} kol${c} ${x} vs ${v(row.getCell(+c).value)}`); }
  cmp++; if (it.tanggal !== v(row.getCell(24).value).toISOString().slice(0, 10)) miss(`r${r} tgl ${it.tanggal}`);
  const h = hitungGedung(it, masa[it.kode], tp, 2026); const E = { AA: 27, AB: 28, AC: 29, AD: 30, AE: 31, AG: 33, AH: 34, AI: 35, AJ: 36, AK: 37, AL: 38, AM: 39, AN: 40, AO: 41 };
  for (const [k, c] of Object.entries(E)) { const ev = v(row.getCell(c).value); cmp++; if (h[k] == null ? (ev !== '' && ev != null) : Math.abs(h[k] - (ev ?? 0)) > 0.01) miss(`r${r} ${k}: ${h[k]} vs ${ev}`); }
  cmp++; if (h.AP !== v(row.getCell(42).value)) miss(`r${r} AP ${h.AP} vs ${v(row.getCell(42).value)}`); }
const hs = items.map((r) => hitungGedung(r, masa[r.kode], tN.get(r.nibar) || {}, 2026));
console.log('dibandingkan', cmp, 'sel | selisih', bad, '| PDF tanpa template:', items.filter((r) => !tN.has(r.nibar)).length, '| template tanpa PDF:', tpl.rows.filter((r) => !byN.has(r.nibar)).length);
console.log('total nilai', items.reduce((a, r) => a + r.nilai, 0).toFixed(2), '| beban', hs.reduce((a, h) => a + h.AL, 0).toFixed(2), '| akum', hs.reduce((a, h) => a + h.AM, 0).toFixed(2), '| nilai buku', hs.reduce((a, h) => a + h.AN, 0).toFixed(2));
process.exit(bad ? 1 : 0);
