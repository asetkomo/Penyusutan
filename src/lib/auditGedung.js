// Pemeriksaan kewajaran data Gedung & Bangunan (ERROR / WARN / INFO).
import { rekonGroups } from './audit.js';
import { bulanTerpakai } from './calc.js';

export function auditGedung({ rows, pairs, hasil, orphan, dupTemplate, tahun }) {
  const f = []; const add = (level, kode, it, msg) => f.push({ level, kode, nibar: it?.nibar || '', nama: it?.nama || it?.spesNama || '', msg });
  const seen = new Map();
  pairs.forEach(({ it, tp }, i) => {
    const h = hasil[i];
    if (!tp) { add('ERROR', 'TEMPLATE_TIDAK_ADA', it, 'Barang ada di PDF tetapi tidak ada di template; saldo awal dianggap 0 sehingga hasil tidak valid.'); return; }
    if (!h.AA) add('ERROR', 'MASA_TIDAK_ADA', it, `Kode ${it.kode} tidak ada di master masa manfaat.`);
    if (h.AP === 'CEK NILAI PEROLEHAN') add('ERROR', 'NILAI_TIDAK_BALANS', it, `Nilai perolehan PDF ${it.nilai} <> AKM + nilai buku + kapitalisasi (selisih ${h.AO}).`);
    if (h.AP === 'CEK TGL KAPITALISASI') add('ERROR', 'TGL_KAPITALISASI', it, `Tanggal kapitalisasi (${tp.tglKap || 'kosong'}) bukan di tahun ${tahun}.`);
    if (h.AP === 'CEK TAMBAHAN MASA MANFAAT') add('WARN', 'TAMBAHAN_MASA_KOSONG', it, 'Ada kapitalisasi tetapi tambahan masa manfaat 0/kosong.');
    if (tp.tgl && it.tanggal && tp.tgl !== it.tanggal) add('WARN', 'TANGGAL_BEDA', it, `Tanggal perolehan PDF ${it.tanggal} berbeda dari template ${tp.tgl}.`);
    if ((h.AB < 0) || (h.AC < 0) || (h.AB > it.nilai + 1)) add('ERROR', 'SALDO_TIDAK_WAJAR', it, `AKM ${h.AB} / nilai buku ${h.AC} tidak wajar terhadap nilai perolehan ${it.nilai}.`);
    if (h.AE > it.nilai) add('ERROR', 'KAPITALISASI_LEBIH_BESAR', it, `Kapitalisasi ${h.AE} melebihi nilai perolehan ${it.nilai}.`);
    // Konsistensi sisa masa manfaat dengan garis lurus murni dari tanggal perolehan (tanpa kapitalisasi tahun-tahun lalu).
    if (h.AA && it.tanggal) { const exp = h.AA - bulanTerpakai(it.tanggal, tahun - 1, h.AA, 0);
      if (Math.abs(exp - h.AD) > 1) add('WARN', 'SISA_MASA_BEDA', it, `Sisa masa manfaat akhir ${tahun - 1} di template ${h.AD} bln; garis lurus dari tanggal perolehan ${it.tanggal} memberi ${exp} bln (selisih ${h.AD - exp}).`); }
    if (h.AK === 0 && h.AL > 0) add('INFO', 'HABIS_TAHUN_INI', it, `Masa manfaat habis tahun ini; beban ${Math.round(h.AL * 100) / 100}.`);
    else if (h.AD === 0 && h.AC === 0) add('INFO', 'SUDAH_HABIS', it, 'Sudah habis disusutkan sebelum tahun berjalan.');
    if (it.nibar) { if (seen.has(it.nibar)) add('ERROR', 'NIBAR_GANDA_PDF', it, `NIBAR ganda di PDF (baris ke-${seen.get(it.nibar)}).`); else seen.set(it.nibar, i + 1); }
  });
  orphan.forEach((t) => add('ERROR', 'PDF_TIDAK_ADA', { nibar: t.nibar }, `NIBAR di template (baris ${t.baris}) tidak ditemukan di PDF; tidak ikut dihitung.`));
  dupTemplate.forEach((t) => add('ERROR', 'NIBAR_GANDA_TEMPLATE', { nibar: t.nibar }, `NIBAR muncul lebih dari sekali di template (baris ${t.baris}); hanya baris pertama dipakai.`));
  f.push(...rekonGroups(rows));
  return f;
}

// Gabungkan PDF + template berdasarkan NIBAR. Urutan hasil mengikuti template; barang PDF tanpa template ditaruh di akhir.
export function gabung(items, tplRows) {
  const byN = new Map(); const dupTemplate = [];
  tplRows.forEach((t) => { if (byN.has(t.nibar)) dupTemplate.push(t); else byN.set(t.nibar, t); });
  const pdfN = new Map(); items.forEach((it) => { if (!pdfN.has(it.nibar)) pdfN.set(it.nibar, it); });
  const pairs = []; const used = new Set();
  tplRows.forEach((t) => { const it = pdfN.get(t.nibar); if (it && !used.has(t.nibar)) { pairs.push({ it, tp: byN.get(t.nibar) }); used.add(t.nibar); } });
  items.forEach((it) => { if (!used.has(it.nibar) || !byN.has(it.nibar)) { pairs.push({ it, tp: null }); used.add(it.nibar); } });
  const orphan = tplRows.filter((t) => !pdfN.has(t.nibar) && !dupTemplate.includes(t));
  return { pairs, orphan, dupTemplate };
}
