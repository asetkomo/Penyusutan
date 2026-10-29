import { useMemo, useState, useEffect } from 'react';
import { loadPages } from './lib/loadPdf.js';
import { extractRecords, extractMeta, detectFormat } from './lib/parsePdf.js';
import { bacaTemplate } from './lib/templateGedung.js';
import { hitungGedung } from './lib/calcGedung.js';
import { auditGedung, gabung } from './lib/auditGedung.js';
import { gedungBlob } from './lib/exportGedung.js';
import { sb, saveRun, listRuns, loadRun, deleteRun } from './lib/supabase.js';

const rp = (v) => new Intl.NumberFormat('id-ID', { maximumFractionDigits: 2, minimumFractionDigits: 2 }).format(v || 0);
const tgl = (iso) => (iso ? iso.split('-').reverse().join('/') : '');

export default function GedungApp({ masa, src }) {
  const [rows, setRows] = useState([]); const [meta, setMeta] = useState({}); const [namaPdf, setNamaPdf] = useState('');
  const [tpl, setTpl] = useState(null); const [namaTpl, setNamaTpl] = useState(''); const [tahun, setTahun] = useState(new Date().getFullYear());
  const [busy, setBusy] = useState(''); const [msg, setMsg] = useState(''); const [warn, setWarn] = useState([]);
  const [user, setUser] = useState(null); const [runs, setRuns] = useState([]);
  useEffect(() => { if (sb) { sb.auth.getUser().then(({ data }) => setUser(data.user)); sb.auth.onAuthStateChange((_, s) => setUser(s?.user ?? null)); } }, []);
  const refresh = async () => setRuns((await listRuns()).filter((r) => r.params?.jenis === 'gedung'));
  useEffect(() => { if (user) refresh(); else setRuns([]); }, [user]);

  const items = useMemo(() => rows.filter((r) => r.type === 'item'), [rows]);
  const ready = items.length > 0 && tpl && masa;
  const g = useMemo(() => (ready ? gabung(items, tpl.rows) : null), [ready, items, tpl]);
  const hasil = useMemo(() => (g ? g.pairs.map(({ it, tp }) => hitungGedung(it, masa[it.kode], tp || {}, tahun)) : []), [g, masa, tahun]);
  const temuan = useMemo(() => (g ? auditGedung({ rows, pairs: g.pairs, hasil, orphan: g.orphan, dupTemplate: g.dupTemplate, tahun }) : []), [g, hasil, rows, tahun]);
  const S = (k) => hasil.reduce((a, h) => a + (h[k] || 0), 0); const cnt = (l) => temuan.filter((t) => t.level === l).length;
  const ringkasan = { item: items.length, nilai: g ? g.pairs.reduce((a, p) => a + (p.it.nilai || 0), 0) : 0, akmLalu: S('AB'), beban: S('AL'), akum: S('AM'), buku: S('AN'), error: cnt('ERROR'), warn: cnt('WARN') };

  async function onPdf(e) {
    const f = e.target.files[0]; if (!f) return; setBusy('Membaca PDF…'); setMsg('');
    try { const pages = await loadPages(f, (p, n) => setBusy(`Membaca halaman ${p}/${n}…`));
      if (detectFormat(pages) !== 'gedung') throw new Error('PDF ini bukan FORMAT II.O.2.3 (Gedung & Bangunan). Untuk Peralatan & Mesin gunakan tab sebelahnya.');
      const out = extractRecords(pages, 'gedung'); const m = extractMeta(pages);
      if (!out.rows.some((r) => r.type === 'item')) throw new Error('Tidak ada baris barang terbaca dari PDF.');
      setRows(out.rows); setWarn(out.warnings); setMeta(m); setNamaPdf(f.name); if (m.tahun) setTahun(m.tahun);
    } catch (err) { setMsg('Gagal: ' + err.message); } setBusy('');
  }
  async function onTpl(e) {
    const f = e.target.files[0]; if (!f) return; setBusy('Membaca template…'); setMsg('');
    try { const t = await bacaTemplate(await f.arrayBuffer()); setTpl(t); setNamaTpl(f.name); if (t.tahunLalu) setTahun((y) => (meta.tahun ? y : t.tahunLalu + 1)); }
    catch (err) { setMsg('Gagal: ' + err.message); } setBusy('');
  }
  async function unduh() { setBusy('Menyusun Excel…'); const blob = await gedungBlob({ pairs: g.pairs, hasil, temuan, masa, tahun }); const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `penyusutan_gedung_${tahun}_${(meta.kodeLokasi || '').replace(/\W/g, '')}.xlsx`; a.click(); setBusy(''); }
  async function simpan() { setBusy('Menyimpan…'); try { await saveRun({ namaFile: namaPdf, meta, tahun, params: { jenis: 'gedung', namaTpl }, ringkasan, rows: [...rows, ...tpl.rows.map((t) => ({ type: 'tpl', ...t }))] }); await refresh(); setMsg('Tersimpan.'); } catch (e) { setMsg('Gagal simpan: ' + e.message); } setBusy(''); }
  async function muat(r) { setBusy('Memuat…'); const all = await loadRun(r.id); setRows(all.filter((x) => x.type !== 'tpl')); setTpl({ rows: all.filter((x) => x.type === 'tpl'), tahunLalu: r.tahun - 1, warnings: [] }); setMeta({ pengguna: r.pengguna_barang }); setTahun(r.tahun); setNamaPdf(r.nama_file); setNamaTpl(r.params?.namaTpl || '(tersimpan)'); setBusy(''); }

  return (<div>
    <div className="mut">Gedung & Bangunan (FORMAT II.O.2.3): PDF e-BMD + template saldo awal (Format_penyusutan_gedung.xlsx) → Excel penyusutan. Master masa manfaat: {src}.</div>
    <div className="card"><h2>1. Unggah PDF dan template</h2><div className="grid">
      <div><label>PDF Daftar BMD Gedung & Bangunan (e-BMD)</label><input type="file" accept="application/pdf" onChange={onPdf} disabled={!masa || !!busy} /></div>
      <div><label>Template saldo awal (.xlsx): NIBAR, AKM, nilai buku, sisa masa manfaat, kapitalisasi</label><input type="file" accept=".xlsx" onChange={onTpl} disabled={!masa || !!busy} /></div></div>
      {busy && <p>{busy}</p>}{msg && <p className="WARN">{msg}</p>}
      {items.length > 0 && <p className="mut">PDF: {namaPdf} · {meta.pengguna} · {items.length} barang</p>}
      {tpl && <p className="mut">Template: {namaTpl} · {tpl.rows.length} baris{tpl.tahunLalu ? ` · saldo akhir ${tpl.tahunLalu}` : ''}</p>}
      {[...warn, ...(tpl?.warnings || [])].slice(0, 10).map((w, i) => <div key={i} className="WARN">{w}</div>)}
      {tpl?.tahunLalu && meta.tahun && tpl.tahunLalu + 1 !== meta.tahun && <p className="WARN">Perhatian: template berisi saldo akhir {tpl.tahunLalu}, sedangkan PDF tahun {meta.tahun}.</p>}</div>

    {ready && <>
      <div className="card"><h2>2. Parameter</h2><div className="row"><div><label>Tahun penyusutan (akhir periode 31 Des)</label><input type="number" value={tahun} onChange={(e) => setTahun(+e.target.value)} /></div></div>
        <p className="mut">Rumus mengikuti Excel acuan: beban = nilai buku awal ÷ sisa masa manfaat × bulan berjalan (maks. 12); jika ada kapitalisasi, dipecah sebelum/sesudah bulan kapitalisasi dengan sisa masa manfaat baru.</p></div>

      <div className="card"><h2>3. Ringkasan {tahun}</h2><div className="grid">
        <div className="stat"><span className="mut">Nilai perolehan</span><b>{rp(ringkasan.nilai)}</b></div>
        <div className="stat"><span className="mut">AKM akhir {tahun - 1}</span><b>{rp(ringkasan.akmLalu)}</b></div>
        <div className="stat"><span className="mut">Beban penyusutan {tahun}</span><b>{rp(ringkasan.beban)}</b></div>
        <div className="stat"><span className="mut">Akumulasi akhir {tahun}</span><b>{rp(ringkasan.akum)}</b></div>
        <div className="stat"><span className="mut">Nilai buku akhir {tahun}</span><b>{rp(ringkasan.buku)}</b></div>
        <div className="stat"><span className="mut">Temuan</span><b><span className="ERROR">{ringkasan.error} error</span> · <span className="WARN">{ringkasan.warn} peringatan</span></b></div></div>
        <div className="row" style={{ marginTop: 14 }}><button onClick={unduh} disabled={!!busy}>Unduh Excel</button>
          {sb && user && <button className="sec" onClick={simpan} disabled={!!busy}>Simpan ke Supabase</button>}</div>
        {sb && !user && <p className="mut">Untuk menyimpan riwayat, masuk lewat kartu Akun di tab Peralatan & Mesin.</p>}</div>

      <div className="card"><h2>Temuan pemeriksaan</h2>
        {!temuan.filter((t) => t.level !== 'INFO').length && <p className="ok">Tidak ada error/peringatan. Nilai perolehan, saldo awal, dan total kelompok PDF seimbang.</p>}
        <div className="tw"><table><thead><tr><th>Tingkat</th><th>Kode</th><th>Uraian</th><th>Keterangan</th></tr></thead><tbody>
          {[...temuan].sort((a, b) => 'EWI'.indexOf(a.level[0]) - 'EWI'.indexOf(b.level[0])).slice(0, 300).map((t, i) => <tr key={i}><td className={t.level}>{t.level}</td><td>{t.kode}</td><td>{t.nama}</td><td>{t.msg}</td></tr>)}</tbody></table></div></div>

      <div className="card"><h2>Pratinjau perhitungan</h2><div className="tw"><table><thead><tr><th>Kode</th><th>Nama</th><th>Tgl</th><th className="n">Nilai</th><th className="n">Masa</th><th className="n">AKM {tahun - 1}</th><th className="n">Sisa {tahun - 1}</th><th className="n">Kapitalisasi</th><th className="n">Beban {tahun}</th><th className="n">Akumulasi</th><th className="n">Nilai buku</th><th>Cek</th></tr></thead><tbody>
        {g.pairs.map(({ it }, i) => { const h = hasil[i]; return <tr key={i}><td>{it.kode}</td><td>{it.spesNama || it.nama}</td><td>{tgl(it.tanggal)}</td><td className="n">{rp(it.nilai)}</td><td className="n">{h.AA}</td><td className="n">{rp(h.AB)}</td><td className="n">{h.AD}</td><td className="n">{h.AE ? rp(h.AE) : ''}</td><td className="n">{rp(h.AL)}</td><td className="n">{rp(h.AM)}</td><td className="n">{rp(h.AN)}</td><td className={h.AP === 'OK' ? 'ok' : 'ERROR'}>{h.AP}</td></tr>; })}</tbody></table></div></div></>}

    {sb && user && runs.length > 0 && <div className="card"><h2>Riwayat Gedung & Bangunan</h2><div className="tw"><table><tbody>{runs.map((r) => <tr key={r.id}><td>{new Date(r.created_at).toLocaleString('id-ID')}</td><td>{r.pengguna_barang}</td><td>{r.tahun}</td><td className="n">{rp(r.ringkasan?.beban)}</td><td><button className="sec" onClick={() => muat(r)}>Muat</button> <button className="sec" onClick={async () => { await deleteRun(r.id); refresh(); }}>Hapus</button></td></tr>)}</tbody></table></div></div>}
  </div>);
}
