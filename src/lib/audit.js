// Pemeriksaan kewajaran data ala pemeriksa: setiap temuan diberi tingkat ERROR / WARN / INFO.
const rp = (v) => Math.round(v * 100) / 100;
export function audit(rows, hasil, { tahun, batasKapitalisasi = 0 }) {
  const f = []; const add = (level, kode, row, msg) => f.push({ level, kode, nibar: row?.nibar || '', nama: row?.nama || '', msg });
  const items = rows.filter((r) => r.type === 'item'); const seen = new Map();
  items.forEach((r, i) => {
    const h = hasil[i];
    if (h.error) add('ERROR', 'KODE_TIDAK_ADA', r, `Kode ${r.kode} tidak ada di master masa manfaat; penyusutan tidak dihitung.`);
    if (!r.tanggal) add('ERROR', 'TANGGAL_TIDAK_VALID', r, `Tanggal perolehan "${r.tanggalText}" tidak terbaca.`);
    else if (r.tanggal > `${tahun}-12-31`) add('ERROR', 'TANGGAL_SETELAH_PERIODE', r, `Diperoleh ${r.tanggalText}, setelah akhir tahun ${tahun}.`);
    if (r.nilai == null || r.nilai <= 0) add('ERROR', 'NILAI_KOSONG', r, 'Nilai perolehan kosong atau nol.');
    else if (r.harga != null && r.jumlah != null && Math.abs(r.harga * r.jumlah - r.nilai) > 1) add('WARN', 'NILAI_TIDAK_SAMA', r, `Harga satuan x jumlah (${rp(r.harga * r.jumlah)}) berbeda dari nilai perolehan (${r.nilai}).`);
    if (batasKapitalisasi > 0 && r.harga != null && r.harga < batasKapitalisasi) add('WARN', 'DI_BAWAH_BATAS_KAPITALISASI', r, `Harga satuan ${r.harga} di bawah batas kapitalisasi ${batasKapitalisasi}.`);
    if (r.tanggal && r.tanggal.endsWith('-12-31') && r.tanggal < `${tahun}-01-01`) add('INFO', 'TANGGAL_31_DES', r, 'Tanggal 31 Desember sering dipakai sebagai tanggal estimasi; bulan perolehan dihitung penuh sehingga memengaruhi 1 bulan penyusutan.');
    if (!h.error && h.kini >= h.masa) add('INFO', 'HABIS_DISUSUTKAN', r, `Masa manfaat ${h.masa} bulan sudah habis; nilai buku akhir ${rp(h.buku)}.`);
    const key = r.nibar + '|' + r.register;
    if (r.nibar && seen.has(key)) add('ERROR', 'NIBAR_GANDA', r, `NIBAR/register sama dengan baris ke-${seen.get(key)}; kemungkinan pencatatan ganda.`);
    else seen.set(key, i + 1);
  });
  f.push(...rekonGroups(rows));
  return f;
}

// Rekonsiliasi: total tiap baris induk di PDF vs jumlah baris rinci di bawahnya.
export function rekonGroups(rows) {
  const f = []; const add = (level, kode, row, msg) => f.push({ level, kode, nibar: row?.nibar || '', nama: row?.nama || '', msg });
  const items = rows.filter((r) => r.type === 'item');
  rows.filter((r) => r.type === 'group').forEach((g) => {
    const kids = items.filter((r) => r.kode.startsWith(g.kode + '.') || r.kode === g.kode);
    const sn = kids.reduce((a, r) => a + (r.nilai || 0), 0); const sj = kids.reduce((a, r) => a + (r.jumlah || 0), 0);
    if (g.nilai != null && Math.abs(sn - g.nilai) > 1) add('ERROR', 'REKON_NILAI', g, `Total ${g.kode} di PDF ${g.nilai}, jumlah rincian ${rp(sn)} (selisih ${rp(g.nilai - sn)}).`);
    if (g.jumlah != null && Math.abs(sj - g.jumlah) > 0.001) add('ERROR', 'REKON_JUMLAH', g, `Jumlah unit ${g.kode} di PDF ${g.jumlah}, rincian ${sj}.`);
  });
  return f;
}
