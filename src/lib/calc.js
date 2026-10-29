// Penyusutan metode garis lurus per bulan, mengikuti logika kolom AB..AL pada dpu_2024_NEW.xlsx.
// Perbedaan yang disengaja: tidak memakai teks tanggal "31/12/2023" (bergantung locale), tidak membagi nol.
const ym = (iso) => { const [y, m] = iso.split('-').map(Number); return y * 12 + (m - 1); };

// Jumlah bulan terpakai sampai 31 Desember tahun `thn`, dibatasi masa manfaat.
// offset=0: bulan perolehan dihitung penuh (sama dengan DATEDIF(...)+1 di Excel); offset=1: mulai bulan berikutnya.
export function bulanTerpakai(tglISO, thn, masa, offset = 0) {
  if (!tglISO || !masa) return 0;
  const n = (thn * 12 + 11) - ym(tglISO) + 1 - offset;
  return Math.max(0, Math.min(masa, n));
}

// Aturan koreksi: [{kode:'1.3.2.05.02.01', lama:120}] -> masa manfaat yang dipakai di saldo awal (tahun lalu) yang keliru.
export function cariMasaLama(kode, rules = []) {
  let best = null;
  for (const r of rules) { const k = (r.kode || '').trim(); if (k && +r.lama > 0 && (kode === k || kode.startsWith(k + '.')) && (!best || k.length > best.k.length)) best = { k, lama: +r.lama }; }
  return best ? best.lama : null;
}

export function hitung(it, masa, { tahun, offset = 0, masaLama = null }) {
  if (!masa) return { masa: null, error: 'KODE_TIDAK_ADA' };
  const nilai = it.nilai ?? 0;
  const perBulan = nilai / masa;
  const lalu = bulanTerpakai(it.tanggal, tahun - 1, masa, offset);
  const kini = bulanTerpakai(it.tanggal, tahun, masa, offset);
  const base = {
    masa, perBulan, lalu, akumLalu: perBulan * lalu, bukuLalu: Math.max(0, nilai - perBulan * lalu), sisaLalu: masa - lalu,
    kini, beban: Math.max(0, kini - lalu) * perBulan, akum: perBulan * kini, buku: nilai - perBulan * kini, sisa: masa - kini,
  };
  // Penyesuaian saldo awal: akumulasi akhir tahun lalu yang dilaporkan (masa lama) vs seharusnya (masa baru).
  if (masaLama && masaLama !== masa) {
    const lamaLama = bulanTerpakai(it.tanggal, tahun - 1, masaLama, offset);
    const akumLaluLama = (nilai / masaLama) * lamaLama;
    Object.assign(base, { masaLama, lamaLama, akumLaluLama, koreksi: base.akumLalu - akumLaluLama });
  }
  return base;
}
