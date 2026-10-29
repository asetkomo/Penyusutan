// Membaca template input Gedung (Format_penyusutan_gedung.xlsx): kunci NIBAR + saldo akhir tahun lalu + kapitalisasi tahun berjalan.
import ExcelJS from 'exceljs';
import { parseTgl } from './parsePdf.js';

const val = (c) => { const v = c?.value; return v && typeof v === 'object' && !(v instanceof Date) ? (v.result ?? v.text ?? null) : v; };
const toNum = (v) => { if (v == null || v === '') return null; const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/\./g, '').replace(',', '.')); return Number.isFinite(n) ? n : null; };
const toISO = (v) => {
  if (v == null || v === '') return null;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'number') return new Date(Date.UTC(1899, 11, 30) + v * 864e5).toISOString().slice(0, 10);
  const s = String(v).trim(); return parseTgl(s) || (/^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : null);
};
const KEYS = [['nibar', (h) => h.includes('nibar')], ['tgl', (h) => /(tgl|tanggal) perolehan/.test(h)], ['akm', (h) => h.startsWith('akm')],
  ['nb', (h) => h.includes('nilai buku')], ['sisa', (h) => h.includes('sisa masa manfaat')], ['kap', (h) => h.includes('kapitalisasi tahun berjalan')],
  ['tglKap', (h) => /(tgl|tanggal) kapitalisasi/.test(h)], ['tambahan', (h) => h.includes('tambahan masa manfaat')]];

export async function bacaTemplate(buffer) {
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buffer); const ws = wb.worksheets[0];
  let hr = 0; const map = {}; let tahunLalu = null;
  for (let r = 1; r <= Math.min(10, ws.rowCount) && !hr; r++) ws.getRow(r).eachCell((c, n) => { if (/nibar/i.test(String(val(c) ?? ''))) hr = r; });
  if (!hr) throw new Error('Header "NIBAR" tidak ditemukan pada template.');
  ws.getRow(hr).eachCell((c, n) => { const h = String(val(c) ?? '').toLowerCase().replace(/\s+/g, ' ').trim(); for (const [k, f] of KEYS) if (!map[k] && f(h)) { map[k] = n; if (k === 'akm') { const y = /(\d{4})/.exec(h); if (y) tahunLalu = +y[1]; } } });
  const miss = ['nibar', 'akm', 'nb', 'sisa'].filter((k) => !map[k]); if (miss.length) throw new Error('Kolom template tidak ditemukan: ' + miss.join(', '));
  const rows = []; const warnings = [];
  for (let r = hr + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r); const raw = val(row.getCell(map.nibar)); if (raw == null || String(raw).trim() === '') continue;
    const nibar = typeof raw === 'number' ? String(BigInt(Math.round(raw))) : String(raw).replace(/\D/g, '');
    if (nibar.length !== 45) warnings.push(`Template baris ${r}: NIBAR berpanjang ${nibar.length} digit (seharusnya 45); pastikan kolom berformat teks.`);
    const g = (k) => (map[k] ? val(row.getCell(map[k])) : null);
    rows.push({ baris: r, nibar, tgl: toISO(g('tgl')), akm: toNum(g('akm')), nb: toNum(g('nb')), sisa: toNum(g('sisa')), kap: toNum(g('kap')), tglKap: toISO(g('tglKap')), tambahan: toNum(g('tambahan')) });
  }
  return { rows, tahunLalu, warnings };
}
