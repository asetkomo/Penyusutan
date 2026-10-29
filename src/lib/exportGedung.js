// Excel hasil Gedung & Bangunan: format dan rumus sama dengan diskop_penyusutan_2026.xlsx (kolom A-AP), sheet Pivot (SUMIFS), master '1', Temuan.
import ExcelJS from 'exceljs';
const COL = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const UANG = '#,##0.00';

export async function buildGedungWorkbook({ pairs, hasil, temuan, masa, tahun }) {
  const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('Sheet', { views: [{ state: 'frozen', xSplit: 10, ySplit: 2 }] }); const ty = tahun - 1;
  const H = ['Kode Aset 7 Digit', 'Kode Aset 6 Digit', 'Nama Aset', 'NIBAR', ' Nomor Register', 'Spesifikasi Nama Barang', 'Spesifikasi Lainnya', 'Jumlah Lantai', 'Lokasi', 'Titik Koordinat', 'Status Kepemilikan Tanah', 'Jumlah', 'Satuan', 'Harga Satuan Perolehan', 'Nilai Perolehan (Rp)', 'Cara Perolehan', 'Tanggal  Perolehan', 'Status Penggunaan', 'Keterangan', 'Masa Manfaat (bulan)',
    `AKM akhir ${ty} (Rp)`, `Nilai buku akhir ${ty} (Rp)`, `Sisa masa manfaat akhir ${ty} (bln)`, 'Kapitalisasi tahun berjalan (Rp)', 'Tanggal kapitalisasi', 'Tambahan masa manfaat (bln)', '% Kapitalisasi', 'Bulan sebelum kapitalisasi', 'Bulan sesudah kapitalisasi',
    `Sisa masa manfaat akhir ${tahun} (bln)`, `Beban penyusutan ${tahun} (Rp)`, `Akumulasi penyusutan akhir ${tahun} (Rp)`, `Nilai buku akhir ${tahun} (Rp)`, 'Selisih nilai perolehan (harus 0)', 'Status cek'];
  ws.addRow([1, 2, 3, 4, 5, 6, 7, ...H]); ws.addRow(Array.from({ length: 42 }, (_, i) => i + 1));
  const first = 3; const last = first + pairs.length - 1; const yEnd = `DATE(${tahun},12,31)`;
  pairs.forEach(({ it, tp }, i) => {
    const n = first + i; const h = hasil[i]; const row = ws.getRow(n); const set = (c, v) => { if (v !== null && v !== undefined && v !== '') row.getCell(c).value = v; };
    it.seg.forEach((s, k) => set(1 + k, s));
    row.getCell(8).value = { formula: `A${n}&"."&B${n}&"."&C${n}&"."&D${n}&"."&E${n}&"."&F${n}&"."&G${n}`, result: it.kode };
    row.getCell(9).value = { formula: `LEFT(H${n},14)`, result: it.kode6 };
    [[10, it.nama], [11, it.nibar], [12, it.register], [13, it.spesNama], [14, it.spesLain], [15, it.jmlLantai ?? 0], [16, it.lokasi], [17, it.koordinat], [18, it.statusTanah], [19, it.jumlah], [20, it.satuan], [21, it.harga], [22, it.nilai], [23, it.cara], [25, it.status], [26, it.ket]].forEach(([c, v]) => set(c, v));
    if (it.tanggal) { row.getCell(24).value = new Date(it.tanggal + 'T00:00:00Z'); row.getCell(24).numFmt = 'dd/mm/yyyy'; } else set(24, it.tanggalText);
    const F = (c, formula, result) => { row.getCell(c).value = { formula, result: result === null || result === undefined || Number.isNaN(result) ? '' : result }; };
    F(27, `IFERROR(INDEX('1'!A:A,MATCH(H${n},'1'!B:B,0)),0)`, h.AA);
    if (tp) { set(28, tp.akm); set(29, tp.nb); set(30, tp.sisa); set(31, tp.kap); if (tp.tglKap) { row.getCell(32).value = new Date(tp.tglKap + 'T00:00:00Z'); row.getCell(32).numFmt = 'dd/mm/yyyy'; } set(33, tp.tambahan); }
    F(34, `IF(N(AE${n})>0,AE${n}/V${n},"")`, h.AH); F(35, `IF(N(AE${n})>0,MONTH(AF${n})-1,12)`, h.AI); F(36, `IF(N(AE${n})>0,13-MONTH(AF${n}),0)`, h.AJ);
    F(37, `IF(N(AE${n})>0,MAX(0,MIN(AD${n}-12+N(AG${n}),AA${n}-AJ${n})),MAX(0,AD${n}-12))`, h.AK);
    F(38, `MIN(AC${n}+N(AE${n}),IF(N(AE${n})>0,IF(AD${n}>0,AC${n}/AD${n}*AI${n},0)+IF(AK${n}>0,(AC${n}+AE${n})/AK${n}*AJ${n},0),IF(AD${n}>0,AC${n}/AD${n}*MIN(12,AD${n}),0)))`, h.AL);
    F(39, `AB${n}+AL${n}`, h.AM); F(40, `V${n}-AM${n}`, h.AN); F(41, `ROUND(V${n}-(AB${n}+AC${n}+N(AE${n})),2)`, h.AO);
    F(42, `IF(ABS(AO${n})>1,"CEK NILAI PEROLEHAN",IF(AND(N(AE${n})>0,YEAR(AF${n})<>YEAR(${yEnd})),"CEK TGL KAPITALISASI",IF(AND(N(AE${n})>0,N(AG${n})=0),"CEK TAMBAHAN MASA MANFAAT","OK")))`, h.AP);
    [21, 22, 28, 29, 31, 38, 39, 40].forEach((c) => (row.getCell(c).numFmt = UANG)); row.getCell(34).numFmt = '0.00%';
  });
  const tn = last + 2; const tot = ws.getRow(tn); tot.getCell(10).value = 'JUMLAH'; tot.font = { bold: true };
  const sum = (c, f) => { const L = COL(c); tot.getCell(c).value = { formula: `SUM(${L}${first}:${L}${last})`, result: hasil.reduce((a, h, i) => a + f(h, pairs[i]), 0) }; tot.getCell(c).numFmt = UANG; };
  sum(22, (h, p) => p.it.nilai || 0); sum(28, (h) => h.AB); sum(29, (h) => h.AC); sum(31, (h) => h.AE); sum(38, (h) => h.AL); sum(39, (h) => h.AM); sum(40, (h) => h.AN); sum(41, (h) => h.AO);
  tot.getCell(42).value = { formula: `IF(COUNTIF(AP${first}:AP${last},"OK")=${pairs.length},"SEMUA OK","ADA YANG PERLU DICEK")`, result: hasil.every((h) => h.AP === 'OK') ? 'SEMUA OK' : 'ADA YANG PERLU DICEK' };
  [1, 2].forEach((r) => { ws.getRow(r).font = { bold: true }; ws.getRow(r).alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' }; });
  ws.columns.forEach((c, i) => { c.width = i < 7 ? 5 : [9, 10, 11].includes(i) ? 24 : [12, 13, 15, 17, 18, 24].includes(i) ? 30 : i >= 27 ? 19 : 15; });

  // Sheet Pivot: rekap per kode 6 digit dengan SUMIFS (ikut berubah bila data diubah).
  const pv = wb.addWorksheet('Pivot'); pv.getRow(3).values = ['Row Labels', 'Nilai Perolehan', `AKM akhir ${ty}`, `Beban penyusutan ${tahun}`, `Akumulasi penyusutan akhir ${tahun}`, `Nilai buku akhir ${tahun}`]; pv.getRow(3).font = { bold: true };
  const k6 = [...new Set(pairs.map((p) => p.it.kode6))].sort(); const R = (L) => `Sheet!$${L}$${first}:$${L}$${last}`; const src = [['V', (h, p) => p.it.nilai || 0], ['AB', (h) => h.AB], ['AL', (h) => h.AL], ['AM', (h) => h.AM], ['AN', (h) => h.AN]];
  k6.forEach((k, j) => { const n = 4 + j; const row = pv.getRow(n); row.getCell(1).value = k;
    src.forEach(([L, f], c) => { row.getCell(2 + c).value = { formula: `SUMIFS(${R(L)},${R('I')},$A${n})`, result: hasil.reduce((a, h, i) => a + (pairs[i].it.kode6 === k ? f(h, pairs[i]) : 0), 0) }; row.getCell(2 + c).numFmt = UANG; }); });
  const gt = pv.getRow(4 + k6.length); gt.getCell(1).value = 'Grand Total'; gt.font = { bold: true };
  src.forEach(([L, f], c) => { const C = COL(2 + c); gt.getCell(2 + c).value = { formula: `SUM(${C}4:${C}${3 + k6.length})`, result: hasil.reduce((a, h, i) => a + f(h, pairs[i]), 0) }; gt.getCell(2 + c).numFmt = UANG; });
  pv.columns = [{ width: 18 }, ...Array(5).fill({ width: 26 })];

  const m = wb.addWorksheet('1'); m.addRow(['Masa Manfaat', 'KODE 7 DIGIT']); Object.entries(masa).filter(([k]) => k.startsWith('1.3.3.')).forEach(([k, v]) => m.addRow([v, k]));
  const t = wb.addWorksheet('Temuan'); t.addRow(['Tingkat', 'Kode', 'NIBAR / Kode Induk', 'Uraian', 'Keterangan']); t.getRow(1).font = { bold: true };
  temuan.forEach((x) => t.addRow([x.level, x.kode, x.nibar, x.nama, x.msg])); t.columns = [{ width: 9 }, { width: 26 }, { width: 50 }, { width: 36 }, { width: 110 }];
  return wb;
}
export async function gedungBlob(args) { const buf = await (await buildGedungWorkbook(args)).xlsx.writeBuffer(); return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }); }
