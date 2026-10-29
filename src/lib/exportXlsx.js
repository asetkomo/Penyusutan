// Membuat file Excel berformat sama dengan dpu_2024_NEW.xlsx: kolom A-AL, rumus hidup + nilai cache, sheet master '1', sheet Temuan.
import ExcelJS from 'exceljs';
const HEAD = ['Kode Barang', 1, 2, 3, 4, 5, 6, 'Kode barang 7 digit', 'Kode Barang 6 Digit', 'Nama Barang', 'NIBAR', ' Nomor Register', 'Spesifikasi Nama Barang', 'Spesifikasi Lainnya', 'Merek/tipe', 'Lokasi', 'Nomor Polisi', 'Nomor Rangka', 'Nomor BPKB', 'Jumlah', 'Satuan', 'Harga Satuan Perolehan Rp.', 'Nilai Perolehan Rp.', 'Cara Perolehan', 'Tanggal  Perolehan', 'Status Penggunaan', 'Keterangan', 'Masa Manfaat (bulan)', 'Nilai Amortisasi per bulan', 'Masa Manfaat yang telah dilalui (s.d. akhir tahun lalu)', 'Akumulasi Penyusutan Sebelumnya (Rp)', 'Nilai Buku Sebelumnya (Rp)', 'Sisa Masa Manfaat (awal tahun)', 'Masa Manfaat yang telah dilalui (s.d. akhir tahun ini)', 'Beban Penyusutan Tahun Berjalan', 'Akumulasi Penyusutan', 'Nilai Buku (Rp)', 'Sisa Masa Manfaat', 'KOREKSI: Masa Manfaat lama (bulan)', 'KOREKSI: Bulan terpakai s.d. akhir tahun lalu (masa lama)', 'KOREKSI: Akumulasi Penyusutan awal DILAPORKAN (masa lama)', 'KOREKSI: Penyesuaian akumulasi awal (AE - AO)'];

export async function buildWorkbook({ rows, hasil, temuan, masa, tahun, offset = 0, meta = {}, rules = [] }) {
  const wb = new ExcelJS.Workbook(); const ws = wb.addWorksheet('Sheet', { views: [{ state: 'frozen', xSplit: 22, ySplit: 3 }] });
  ws.addRow(HEAD); ws.getRow(1).height = 60; ws.addRow([]); ws.addRow(HEAD.map((_, i) => (i >= 7 ? i - 6 : '')));
  ws.getCell('AD2').value = tahun - 1; ws.getCell('AI2').value = tahun;
  const yEnd = (y) => `DATE(${y},12,31)`; const off = offset ? `-${offset}` : '';
  let hi = 0;
  rows.forEach((r) => {
    const n = ws.rowCount + 1; const isItem = r.type === 'item'; const h = isItem ? hasil[hi++] : null;
    const row = ws.getRow(n);
    r.seg.forEach((s, k) => { if (s) row.getCell(1 + k).value = s; });
    row.getCell(8).value = { formula: `A${n}&"."&B${n}&"."&C${n}&"."&D${n}&"."&E${n}&"."&F${n}&"."&G${n}`, result: r.kode };
    row.getCell(10).value = r.nama; row.getCell(20).value = r.jumlah; row.getCell(23).value = r.nilai;
    if (!isItem) { row.font = { bold: true }; return; }
    row.getCell(9).value = { formula: `LEFT(H${n},14)`, result: r.kode6 };
    [[11, r.nibar], [12, r.register], [13, r.spesNama], [14, r.spesLain], [15, r.merek], [16, r.lokasi], [17, r.nopol], [18, r.rangka], [19, r.bpkb], [21, r.satuan], [22, r.harga], [24, r.cara], [26, r.status], [27, r.ket]].forEach(([c, v]) => { if (v !== '' && v != null) row.getCell(c).value = v; });
    if (r.tanggal) { row.getCell(25).value = new Date(r.tanggal + 'T00:00:00Z'); row.getCell(25).numFmt = 'dd/mm/yyyy'; } else row.getCell(25).value = r.tanggalText;
    const F = (c, formula, result) => { row.getCell(c).value = { formula, result: Number.isFinite(result) ? result : 0 }; };
    F(28, `IFERROR(INDEX('1'!A:A,MATCH(H${n},'1'!B:B,0)),0)`, h.masa || 0);
    F(29, `IF(AB${n}=0,0,W${n}/AB${n})`, h.perBulan);
    F(30, `IFERROR(MIN(AB${n},MAX(0,DATEDIF(Y${n},${yEnd(tahun - 1)},"M")+1${off})),0)`, h.lalu);
    F(31, `AC${n}*AD${n}`, h.akumLalu); F(32, `MAX(0,W${n}-AE${n})`, h.bukuLalu ?? r.nilai); F(33, `AB${n}-AD${n}`, h.sisaLalu);
    F(34, `IFERROR(MIN(AB${n},MAX(0,DATEDIF(Y${n},${yEnd(tahun)},"M")+1${off})),0)`, h.kini);
    if (h.masaLama) {
      row.getCell(39).value = h.masaLama;
      F(40, `IFERROR(MIN(AM${n},MAX(0,DATEDIF(Y${n},${yEnd(tahun - 1)},"M")+1${off})),0)`, h.lamaLama);
      F(41, `W${n}/AM${n}*AN${n}`, h.akumLaluLama); F(42, `AE${n}-AO${n}`, h.koreksi);
    }
    F(35, `MAX(0,AH${n}-AD${n})*AC${n}`, h.beban); F(36, `AH${n}*AC${n}`, h.akum); F(37, `W${n}-AJ${n}`, h.buku ?? r.nilai); F(38, `AB${n}-AH${n}`, h.sisa);
  });
  ws.eachRow((row, n) => { if (n > 3) { [22, 23, 29, 31, 32, 35, 36, 37, 41, 42].forEach((c) => (row.getCell(c).numFmt = '#,##0.00')); } });
  [1, 2, 3].forEach((n) => { ws.getRow(n).font = { bold: true }; ws.getRow(n).alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' }; });
  ws.columns.forEach((c, i) => { c.width = [1, 2, 3, 4, 5, 6].includes(i) ? 4 : i === 9 ? 34 : i >= 21 && i <= 22 ? 18 : i >= 27 ? 17 : 14; });
  const lastRow = ws.rowCount;
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: lastRow, column: 38 } };
  addPivot6(wb, rows, hasil, lastRow);
  const kor = hasil.map((h, i) => ({ h, r: rows.filter((x) => x.type === 'item')[i] })).filter((x) => x.h.masaLama);
  if (kor.length) {
    const k = wb.addWorksheet('Koreksi');
    k.addRow([`Penyesuaian saldo awal akibat koreksi masa manfaat - tahun ${tahun}`]).font = { bold: true };
    k.addRow(['Akumulasi penyusutan akhir tahun lalu yang dilaporkan memakai masa manfaat lama; seharusnya memakai masa manfaat baru. Selisihnya adalah koreksi saldo awal. Beban tahun berjalan sudah memakai masa manfaat baru.']);
    k.addRow([]); k.addRow(['Kode barang (awalan)', 'Masa lama (bulan)', 'Jumlah aset', 'Nilai perolehan', 'Akumulasi awal dilaporkan (masa lama)', 'Akumulasi awal seharusnya (masa baru)', 'Koreksi (tambah akumulasi penyusutan)']).font = { bold: true };
    const first = k.rowCount + 1;
    rules.filter((r) => kor.some((x) => x.r.kode === r.kode || x.r.kode.startsWith(r.kode + '.'))).forEach((r) => {
      const n = k.rowCount + 1; const mine = kor.filter((x) => x.r.kode === r.kode || x.r.kode.startsWith(r.kode + '.'));
      const sm = (f) => mine.reduce((a, x) => a + (x.h[f] || 0), 0); const crit = `Sheet!$H:$H,A${n}&"*",Sheet!$AM:$AM,">0"`;
      const row = k.addRow([r.kode, +r.lama]);
      row.getCell(3).value = { formula: `COUNTIFS(${crit})`, result: mine.length };
      row.getCell(4).value = { formula: `SUMIFS(Sheet!$W:$W,${crit})`, result: mine.reduce((a, x) => a + (x.r.nilai || 0), 0) };
      row.getCell(5).value = { formula: `SUMIFS(Sheet!$AO:$AO,${crit})`, result: sm('akumLaluLama') };
      row.getCell(6).value = { formula: `SUMIFS(Sheet!$AE:$AE,${crit})`, result: sm('akumLalu') };
      row.getCell(7).value = { formula: `F${n}-E${n}`, result: sm('koreksi') };
    });
    const last = k.rowCount; const tot = k.addRow(['TOTAL']); tot.font = { bold: true };
    [3, 4, 5, 6, 7].forEach((c) => { const L = 'ABCDEFG'[c - 1]; tot.getCell(c).value = { formula: `SUM(${L}${first}:${L}${last})`, result: kor.reduce((a, x) => a + (c === 3 ? 1 : c === 4 ? x.r.nilai || 0 : c === 5 ? x.h.akumLaluLama : c === 6 ? x.h.akumLalu : x.h.koreksi), 0) }; });
    k.eachRow((row, n) => { if (n >= first) [4, 5, 6, 7].forEach((c) => (row.getCell(c).numFmt = '#,##0.00')); });
    k.addRow([]); k.addRow(['Catatan: koreksi dicatat sebagai penyesuaian saldo awal (koreksi kesalahan periode lalu). Perlakuan akuntansi/jurnalnya perlu dikonfirmasi dengan kebijakan akuntansi dan peraturan yang berlaku di daerah Anda.']);
    k.columns = [{ width: 24 }, { width: 14 }, { width: 12 }, { width: 22 }, { width: 28 }, { width: 28 }, { width: 28 }];
  }
  const m = wb.addWorksheet('1'); m.addRow(['Masa Manfaat', 'KODE 7 DIGIT']);
  Object.entries(masa).forEach(([k, v]) => m.addRow([v, k]));
  const t = wb.addWorksheet('Temuan'); t.addRow(['Tingkat', 'Kode', 'NIBAR / Kode Induk', 'Uraian', 'Keterangan']); t.getRow(1).font = { bold: true };
  temuan.filter((x) => x.level !== 'INFO' || temuan.length < 2000).forEach((x) => t.addRow([x.level, x.kode, x.nibar, x.nama, x.msg]));
  t.columns = [{ width: 9 }, { width: 28 }, { width: 50 }, { width: 36 }, { width: 100 }];
  return wb;
}
export async function workbookBlob(args) { const buf = await (await buildWorkbook(args)).xlsx.writeBuffer(); return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }); }

// Sheet "Pivot 6 digit": rekap per kode barang 6 digit memakai SUMIFS ke sheet 'Sheet' (hidup, ikut berubah bila data diubah).
// Hanya baris barang (kolom I terisi) yang dijumlahkan, sehingga baris induk/subtotal PDF tidak terhitung ganda.
const ACC = '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-';
function addPivot6(wb, rows, hasil, lastRow) {
  const pv = wb.addWorksheet('Pivot 6 digit'); const uraian = new Map(); rows.forEach((r) => { if (r.type === 'group') uraian.set(r.kode, r.nama); });
  const agg = new Map(); let hi = 0;
  rows.forEach((r) => {
    if (r.type !== 'item') return; const h = hasil[hi++]; const a = agg.get(r.kode6) || [0, 0, 0, 0, 0];
    a[0] += r.nilai || 0; a[1] += h.akumLalu || 0; a[2] += h.beban || 0; a[3] += h.akum || 0; a[4] += h.buku ?? r.nilai ?? 0; agg.set(r.kode6, a);
  });
  const SRC = ['W', 'AE', 'AI', 'AJ', 'AK']; const hdr = ['Row Labels', 'Nilai Perolehan', 'Akumulasi awal', 'Beban', 'Akumulasi Penyusutan Akhir', 'Nilai Buku', 'Uraian Kode 6 Digit'];
  pv.addRow(hdr); pv.getRow(1).height = 30;
  const keys = [...agg.keys()].sort(); const last = keys.length + 1; const tot = [0, 0, 0, 0, 0];
  keys.forEach((k, i) => {
    const n = i + 2; const row = pv.getRow(n); row.getCell(1).value = k; row.getCell(1).alignment = { horizontal: 'left' };
    SRC.forEach((c, j) => { row.getCell(2 + j).value = { formula: `SUMIFS(Sheet!$${c}$4:$${c}$${lastRow},Sheet!$I$4:$I$${lastRow},$A${n})`, result: agg.get(k)[j] }; tot[j] += agg.get(k)[j]; });
    row.getCell(7).value = uraian.get(k) || '';
  });
  const g = last + 1; const gr = pv.getRow(g); gr.getCell(1).value = 'Grand Total';
  SRC.forEach((_, j) => { const L = 'BCDEF'[j]; gr.getCell(2 + j).value = { formula: `SUM(${L}2:${L}${last})`, result: tot[j] }; });
  const root = rows[0] && rows[0].kode === '1.3.2' ? rows[0] : null;
  if (root) { const c = pv.getRow(g + 2); c.getCell(1).value = 'Selisih terhadap total PDF (Sheet!W4), harus 0'; c.getCell(2).value = { formula: `ROUND(B${g}-Sheet!W4,2)`, result: Math.round((tot[0] - root.nilai) * 100) / 100 }; c.getCell(2).numFmt = '#,##0.00'; c.font = { italic: true }; }
  for (let n = 2; n <= g; n++) for (let c = 2; c <= 6; c++) pv.getRow(n).getCell(c).numFmt = ACC;
  const fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
  pv.getRow(1).eachCell((c) => { c.font = { bold: true }; c.fill = fill; c.alignment = { wrapText: true, vertical: 'middle', horizontal: c.col === 1 ? 'left' : 'center' }; c.border = { bottom: { style: 'thin' } }; });
  gr.eachCell({ includeEmpty: true }, (c) => { c.font = { bold: true }; c.fill = fill; c.border = { top: { style: 'thin' } }; });
  [13.1, 15.3, 15.1, 14.3, 19.4, 14.3, 42].forEach((w, i) => (pv.getColumn(i + 1).width = w));
  pv.views = [{ state: 'frozen', ySplit: 1 }];
}
