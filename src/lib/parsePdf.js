// Parser laporan e-BMD: FORMAT II.O.2.2 (Peralatan & Mesin) dan II.O.2.3 (Gedung & Bangunan).
// Fungsi murni: menerima halaman berisi item teks {str,x,y} dari pdf.js, tidak bergantung browser/Node.
export const num = (s) => { if (!s) return null; const v = parseFloat(String(s).replace(/\./g, '').replace(',', '.')); return Number.isFinite(v) ? v : null; };
export const parseTgl = (s) => { const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec((s || '').trim()); return m ? `${m[3]}-${m[2]}-${m[1]}` : null; };
const clean = (s) => s.replace(/\s+/g, ' ').trim();

const CFG = {
  mesin: {
    segX: [35, 47, 58, 69, 83, 97, 111], footX: 1400, hy: /\w-$/,
    cols: [['nama', 125], ['nibar', 280], ['register', 385], ['spesNama', 490], ['spesLain', 592], ['merek', 685], ['lokasi', 780], ['nopol', 900], ['rangka', 960], ['bpkb', 1040], ['jumlah', 1145], ['satuan', 1190], ['harga', 1250], ['nilai', 1345], ['cara', 1405], ['tanggal', 1520], ['status', 1555], ['ket', 1650]],
    item: (t, L) => ({ nibar: L('nibar').join('').replace(/\D/g, ''), register: L('register').join('').replace(/\D/g, ''), spesNama: t('spesNama'), spesLain: t('spesLain'), merek: t('merek'), lokasi: t('lokasi'), nopol: t('nopol'), rangka: t('rangka'), bpkb: t('bpkb') }),
  },
  gedung: {
    segX: [33, 44, 55, 66, 80, 95, 108], footX: 1300, hy: /-$/, // gedung: koordinat negatif ('Kota Mojokerto -' / '7.48...') disambung tanpa spasi
    cols: [['nama', 120], ['nibar', 290], ['register', 395], ['spesNama', 500], ['spesLain', 612], ['jmlLantai', 735], ['lokasi', 770], ['koordinat', 880], ['statusTanah', 995], ['jumlah', 1100], ['satuan', 1150], ['harga', 1200], ['nilai', 1280], ['cara', 1360], ['tanggal', 1480], ['status', 1520], ['ket', 1630]],
    item: (t, L) => ({ nibar: L('nibar').join('').replace(/\D/g, ''), register: L('register').join('').replace(/\D/g, ''), spesNama: t('spesNama'), spesLain: t('spesLain'), jmlLantai: num(t('jmlLantai')) || 0, lokasi: t('lokasi'), koordinat: t('koordinat'), statusTanah: t('statusTanah') }),
  },
};

export function detectFormat(pages) {
  const txt = (pages[0] || []).map((i) => i.str).join(' ');
  if (/II\.O\.2\.3/.test(txt) || /GEDUNG DAN BANGUNAN/i.test(txt)) return 'gedung';
  return 'mesin';
}

export function extractRecords(pages, fmt = 'mesin') {
  const cfg = CFG[fmt]; const colOf = (x) => { let c = null; for (const [k, x0] of cfg.cols) if (x >= x0) c = k; return c; };
  const rows = []; const warnings = []; let cur = null; let stop = false;
  pages.forEach((items, pi) => {
    if (stop) return;
    const hdr = items.find((i) => i.str.trim() === '18' && i.x > 1650 && i.x < 1750);
    const top = hdr ? hdr.y - 8 : Infinity;
    const foot = items.find((i) => /^Mojokerto,/.test(i.str.trim()) && i.x > cfg.footX);
    const bot = foot ? foot.y + 6 : 40;
    const body = items.filter((i) => i.str.trim() && i.y < top && i.y > bot).sort((a, b) => b.y - a.y || a.x - b.x);
    // Teks satu baris logis bisa bergeser 1-3 pt di sumbu Y; item ditempelkan ke record terdekat di atasnya (toleransi 4 pt).
    const starts = body.filter((i) => i.x >= 28 && i.x <= 40 && i.str.trim() === '1').map((i) => i.y);
    const recs = starts.map((y) => { const r = { page: pi + 1, y, seg: [], cells: {} }; rows.push(r); return r; });
    for (const it of body) {
      let k = -1; starts.forEach((y, idx) => { if (y >= it.y - 4) k = idx; });
      const r = k >= 0 ? recs[k] : cur; if (!r) continue;
      if (it.x < 120) {
        let bi = 0, bd = 1e9; cfg.segX.forEach((sx, j) => { const d = Math.abs(it.x - sx); if (d < bd) { bd = d; bi = j; } });
        if (bd <= 6) { r.seg[bi] = it.str.trim(); continue; }
      }
      const c = colOf(it.x); if (!c) continue;
      (r.cells[c] ||= []).push({ s: it.str.trim(), y: it.y, x: it.x });
    }
    if (recs.length) cur = recs[recs.length - 1];
    if (foot) stop = true;
  });
  const out = [];
  for (const r of rows) {
    const toks = (k) => (r.cells[k] || []).slice().sort((a, b) => (Math.abs(b.y - a.y) > 4 ? b.y - a.y : a.x - b.x));
    const L = (k) => toks(k).map((o) => o.s);
    // Baris yang terpotong di tanda hubung ("CU-" / "YN18WKJ") disambung tanpa spasi.
    const t = (k) => clean(toks(k).reduce((acc, o, i, a) => (i === 0 ? o.s : acc + (Math.abs(a[i - 1].y - o.y) > 4 && cfg.hy.test(a[i - 1].s) ? '' : ' ') + o.s), ''));
    const kode7 = r.seg.filter(Boolean).join('.');
    const base = { page: r.page, seg: r.seg.map((s) => s || ''), kode: kode7, nama: t('nama'), jumlah: num(t('jumlah')), nilai: num(t('nilai')) };
    if (!r.cells.nibar) { out.push({ type: 'group', ...base }); continue; }
    if (r.seg.filter(Boolean).length !== 7) warnings.push(`Hal ${r.page}: kode barang tidak lengkap "${kode7}"`);
    out.push({
      type: 'item', ...base, kode6: r.seg.slice(0, 6).join('.'), ...cfg.item(t, L),
      satuan: t('satuan'), harga: num(t('harga')), cara: t('cara'), tanggalText: t('tanggal'), tanggal: parseTgl(t('tanggal')), status: t('status'), ket: t('ket'),
    });
  }
  return { rows: out, warnings };
}

// Metadata kop laporan (halaman 1).
export function extractMeta(pages) {
  const it = pages[0] || [];
  const line = (label) => { const l = it.find((i) => i.str.trim() === label); return l ? it.filter((i) => Math.abs(i.y - l.y) < 3 && i.x > 120).sort((a, b) => a.x - b.x).map((i) => i.str.trim()).join(' ') : ''; };
  const th = it.map((i) => /^TAHUN\s+(\d{4})/.exec(i.str.trim())).find(Boolean);
  const pm = it.find((i) => /^PEMERINTAH/.test(i.str.trim()));
  return { pengguna: line('Pengguna Barang'), kodeLokasi: line('Kode Lokasi'), tahun: th ? +th[1] : null, pemda: pm ? pm.str.trim() : '' };
}
