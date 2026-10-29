import { useMemo, useState, useEffect } from 'react';
import { loadPages } from './lib/loadPdf.js';
import { extractRecords, extractMeta, detectFormat } from './lib/parsePdf.js';
import { getMaster } from './lib/useMaster.js';
import GedungApp from './GedungApp.jsx';
import { hitung, cariMasaLama } from './lib/calc.js';
import { audit } from './lib/audit.js';
import { workbookBlob } from './lib/exportXlsx.js';
import { sb, loadMasterDb, saveRun, listRuns, loadRun, deleteRun } from './lib/supabase.js';

const rp = (v) => new Intl.NumberFormat('id-ID', { maximumFractionDigits: 2, minimumFractionDigits: 2 }).format(v || 0);
const tgl = (iso) => (iso ? iso.split('-').reverse().join('/') : '');

function MesinApp() {
  const [masa, setMasa] = useState(null); const [src, setSrc] = useState('memuat…');
  const [rows, setRows] = useState([]); const [meta, setMeta] = useState({}); const [namaFile, setNamaFile] = useState('');
  const [tahun, setTahun] = useState(new Date().getFullYear() - 1); const [offset, setOffset] = useState(0); const [batas, setBatas] = useState(0);
  const [rules, setRules] = useState([{ kode: '1.3.2.05.02.01', lama: 120 }]); const [warn, setWarn] = useState([]); const [busy, setBusy] = useState(''); const [msg, setMsg] = useState('');
  const [user, setUser] = useState(null); const [runs, setRuns] = useState([]); const [auth, setAuth] = useState({ email: '', password: '' });

  useEffect(() => { getMaster().then(({ masa, src }) => { setMasa(masa); setSrc(src); }); if (sb) { sb.auth.getUser().then(({ data }) => setUser(data.user)); sb.auth.onAuthStateChange((_, s) => setUser(s?.user ?? null)); } }, []);
  useEffect(() => { if (user) listRuns().then(setRuns); else setRuns([]); }, [user]);

  const items = useMemo(() => rows.filter((r) => r.type === 'item'), [rows]);
  const hasil = useMemo(() => (masa ? items.map((r) => hitung(r, masa[r.kode], { tahun, offset, masaLama: cariMasaLama(r.kode, rules) })) : []), [items, masa, tahun, offset, rules]);
  const temuan = useMemo(() => (masa && rows.length ? audit(rows, hasil, { tahun, batasKapitalisasi: +batas || 0 }) : []), [rows, hasil, tahun, batas, masa]);
  const sum = (k) => hasil.reduce((a, h) => a + (h[k] || 0), 0);
  const cnt = (l) => temuan.filter((t) => t.level === l).length;
  const kor = hasil.filter((h) => h.masaLama);
  const korSum = { n: kor.length, lama: kor.reduce((a, h) => a + h.akumLaluLama, 0), baru: kor.reduce((a, h) => a + h.akumLalu, 0), selisih: kor.reduce((a, h) => a + h.koreksi, 0) };
  const ringkasan = { koreksi: korSum.selisih, item: items.length, nilai: items.reduce((a, r) => a + (r.nilai || 0), 0), beban: sum('beban'), akum: sum('akum'), buku: sum('buku'), error: cnt('ERROR'), warn: cnt('WARN') };

  async function onFile(e) {
    const f = e.target.files[0]; if (!f) return; setBusy('Membaca PDF…'); setMsg('');
    try { const pages = await loadPages(f, (p, n) => setBusy(`Membaca halaman ${p}/${n}…`)); if (detectFormat(pages) === 'gedung') throw new Error('PDF ini FORMAT II.O.2.3 (Gedung & Bangunan). Gunakan tab "Gedung & Bangunan".');
      const out = extractRecords(pages); const m = extractMeta(pages);
      if (!out.rows.some((r) => r.type === 'item')) throw new Error('Tidak ada baris barang terbaca. Pastikan PDF berasal dari e-BMD FORMAT II.O.2.2.');
      setRows(out.rows); setWarn(out.warnings); setMeta(m); setNamaFile(f.name); if (m.tahun) setTahun(m.tahun);
    } catch (err) { setMsg('Gagal: ' + err.message); } setBusy('');
  }
  async function unduh() {
    setBusy('Menyusun Excel…'); const blob = await workbookBlob({ rows, hasil, temuan, masa, tahun, offset, meta, rules });
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `penyusutan_${tahun}_${(meta.kodeLokasi || 'dpu').replace(/\W/g, '')}.xlsx`; a.click(); setBusy('');
  }
  const doAuth = async (up) => { const { error } = await (up ? sb.auth.signUp(auth) : sb.auth.signInWithPassword(auth)); setMsg(error ? error.message : up ? 'Cek email untuk konfirmasi pendaftaran.' : ''); };
  async function simpan() { setBusy('Menyimpan…'); try { await saveRun({ namaFile, meta, tahun, params: { offset, batas, rules }, ringkasan, rows }); setRuns(await listRuns()); setMsg('Tersimpan.'); } catch (e) { setMsg('Gagal simpan: ' + e.message); } setBusy(''); }
  async function muat(r) { setBusy('Memuat…'); setRows(await loadRun(r.id)); setMeta({ pengguna: r.pengguna_barang }); setTahun(r.tahun); setOffset(r.params?.offset || 0); setBatas(r.params?.batas || 0); setRules(r.params?.rules || []); setNamaFile(r.nama_file); setBusy(''); }

  return (<div>
        <div className="mut">Ubah PDF e-BMD (Daftar BMD Pengguna Barang, FORMAT II.O.2.2) menjadi Excel penyusutan garis lurus. Master masa manfaat: {src}.</div>

    <div className="card"><h2>1. Unggah PDF dari SIPD e-BMD</h2>
      <input type="file" accept="application/pdf" onChange={onFile} disabled={!masa || !!busy} />
      {busy && <p>{busy}</p>}{msg && <p className="WARN">{msg}</p>}
      {rows.length > 0 && <p className="mut">{namaFile} · {meta.pemda} · {meta.pengguna} · kode lokasi {meta.kodeLokasi} · {items.length} barang</p>}
      {warn.map((w, i) => <div key={i} className="WARN">{w}</div>)}</div>

    {rows.length > 0 && <>
      <div className="card"><h2>2. Parameter</h2><div className="row">
        <div><label>Tahun penyusutan (akhir periode 31 Des)</label><input type="number" value={tahun} onChange={(e) => setTahun(+e.target.value)} /></div>
        <div><label>Awal penyusutan</label><select value={offset} onChange={(e) => setOffset(+e.target.value)}><option value={0}>Bulan perolehan dihitung penuh (sama dgn Excel Anda)</option><option value={1}>Mulai bulan berikutnya</option></select></div>
        <div><label>Batas kapitalisasi per unit (Rp, 0 = tidak diperiksa)</label><input type="number" value={batas} onChange={(e) => setBatas(e.target.value)} /></div></div>
        <p className="mut">Metode garis lurus tanpa nilai residu: beban per bulan = nilai perolehan ÷ masa manfaat (bulan), dibatasi masa manfaat.</p></div>

      <div className="card"><h2>Koreksi masa manfaat tahun lalu (penyesuaian saldo awal)</h2>
        <p className="mut">Untuk kode yang masa manfaatnya keliru di tahun lalu. Masa manfaat <b>baru</b> diambil dari master; isi masa manfaat <b>lama</b> yang dipakai di saldo awal. Selisih akumulasi penyusutan awal dihitung sebagai koreksi.</p>
        {rules.map((r, i) => <div className="row" key={i} style={{ marginBottom: 8 }}>
          <div><label>Kode barang (awalan, mis. 1.3.2.05.02.01)</label><input value={r.kode} onChange={(e) => setRules(rules.map((x, j) => (j === i ? { ...x, kode: e.target.value } : x)))} /></div>
          <div><label>Masa manfaat lama (bulan)</label><input type="number" value={r.lama} onChange={(e) => setRules(rules.map((x, j) => (j === i ? { ...x, lama: +e.target.value } : x)))} /></div>
          <div><label>Masa manfaat baru (master)</label><input disabled value={[...new Set(items.filter((it) => cariMasaLama(it.kode, [r])).map((it) => masa[it.kode]))].join(', ') || '—'} /></div>
          <button className="sec" onClick={() => setRules(rules.filter((_, j) => j !== i))}>Hapus</button></div>)}
        <button className="sec" onClick={() => setRules([...rules, { kode: '', lama: 120 }])}>+ Tambah kode</button>
        {korSum.n > 0 ? <div className="grid" style={{ marginTop: 12 }}>
          <div className="stat"><span className="mut">Aset terdampak</span><b>{korSum.n}</b></div>
          <div className="stat"><span className="mut">Akumulasi awal dilaporkan (masa lama)</span><b>{rp(korSum.lama)}</b></div>
          <div className="stat"><span className="mut">Akumulasi awal seharusnya (masa baru)</span><b>{rp(korSum.baru)}</b></div>
          <div className="stat"><span className="mut">Koreksi saldo awal</span><b>{rp(korSum.selisih)}</b></div></div>
          : <p className="mut" style={{ marginTop: 12 }}>Tidak ada aset yang cocok dengan aturan koreksi.</p>}</div>

      <div className="card"><h2>3. Ringkasan {tahun}</h2><div className="grid">
        <div className="stat"><span className="mut">Nilai perolehan</span><b>{rp(ringkasan.nilai)}</b></div>
        <div className="stat"><span className="mut">Beban penyusutan {tahun}</span><b>{rp(ringkasan.beban)}</b></div>
        <div className="stat"><span className="mut">Akumulasi penyusutan</span><b>{rp(ringkasan.akum)}</b></div>
        <div className="stat"><span className="mut">Nilai buku</span><b>{rp(ringkasan.buku)}</b></div>
        <div className="stat"><span className="mut">Temuan</span><b><span className="ERROR">{ringkasan.error} error</span> · <span className="WARN">{ringkasan.warn} peringatan</span></b></div></div>
        <div className="row" style={{ marginTop: 14 }}><button onClick={unduh} disabled={!!busy}>Unduh Excel</button>
          {sb && user && <button className="sec" onClick={simpan} disabled={!!busy}>Simpan ke Supabase</button>}</div></div>

      <div className="card"><h2>Temuan pemeriksaan</h2>
        {!temuan.filter((t) => t.level !== 'INFO').length && <p className="ok">Tidak ada error/peringatan. Total tiap kelompok di PDF cocok dengan jumlah rincian.</p>}
        <div className="tw"><table><thead><tr><th>Tingkat</th><th>Kode</th><th>NIBAR / induk</th><th>Keterangan</th></tr></thead><tbody>
          {[...temuan].sort((a, b) => 'EWI'.indexOf(a.level[0]) - 'EWI'.indexOf(b.level[0])).slice(0, 300).map((t, i) => <tr key={i}><td className={t.level}>{t.level}</td><td>{t.kode}</td><td>{t.nibar.slice(-18) || t.nama}</td><td>{t.msg}</td></tr>)}</tbody></table></div>
        <p className="mut">Menampilkan maks. 300 baris; seluruh temuan ada di sheet "Temuan" pada Excel.</p></div>

      <div className="card"><h2>Pratinjau perhitungan</h2><div className="tw"><table><thead><tr><th>Kode</th><th>Nama</th><th>Tgl</th><th className="n">Nilai</th><th className="n">Masa (bln)</th><th className="n">Beban {tahun}</th><th className="n">Akumulasi</th><th className="n">Nilai buku</th><th className="n">Sisa (bln)</th></tr></thead><tbody>
        {items.slice(0, 300).map((r, i) => { const h = hasil[i]; return <tr key={i}><td>{r.kode}</td><td>{r.nama}</td><td>{tgl(r.tanggal)}</td><td className="n">{rp(r.nilai)}</td><td className="n">{h.masa ?? '—'}</td><td className="n">{rp(h.beban)}</td><td className="n">{rp(h.akum)}</td><td className="n">{rp(h.buku)}</td><td className="n">{h.sisa ?? '—'}</td></tr>; })}</tbody></table></div>
        {items.length > 300 && <p className="mut">Pratinjau 300 dari {items.length} baris; Excel memuat semuanya.</p>}</div></>}

    {sb && <div className="card"><h2>Akun & riwayat (Supabase)</h2>
      {!user ? <div className="row"><div><label>Email</label><input value={auth.email} onChange={(e) => setAuth({ ...auth, email: e.target.value })} /></div><div><label>Kata sandi</label><input type="password" value={auth.pw} onChange={(e) => setAuth({ ...auth, password: e.target.value })} /></div><button onClick={() => doAuth(false)}>Masuk</button><button className="sec" onClick={() => doAuth(true)}>Daftar</button></div>
        : <><div className="row"><span>{user.email}</span><button className="sec" onClick={() => sb.auth.signOut()}>Keluar</button></div>
          <div className="tw"><table><tbody>{runs.filter((r) => !r.params?.jenis).map((r) => <tr key={r.id}><td>{new Date(r.created_at).toLocaleString('id-ID')}</td><td>{r.pengguna_barang}</td><td>{r.tahun}</td><td className="n">{rp(r.ringkasan?.beban)}</td><td><button className="sec" onClick={() => muat(r)}>Muat</button> <button className="sec" onClick={async () => { await deleteRun(r.id); setRuns(await listRuns()); }}>Hapus</button></td></tr>)}</tbody></table></div></>}</div>}
  </div>);
}

export default function App() {
  const [jenis, setJenis] = useState('mesin'); const [m, setM] = useState({ masa: null, src: 'memuat…' });
  useEffect(() => { getMaster().then(setM); }, []);
  const tab = (k, t) => <button className={jenis === k ? '' : 'sec'} onClick={() => setJenis(k)}>{t}</button>;
  return (<main>
    <h1>Penghitung Penyusutan Aset Tetap</h1>
    <div className="row" style={{ margin: '10px 0' }}>{tab('mesin', 'Peralatan & Mesin')}{tab('gedung', 'Gedung & Bangunan')}</div>
    <div style={{ display: jenis === 'mesin' ? 'block' : 'none' }}><MesinApp /></div>
    <div style={{ display: jenis === 'gedung' ? 'block' : 'none' }}><GedungApp masa={m.masa} src={m.src} /></div>
  </main>);
}
