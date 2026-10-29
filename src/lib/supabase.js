import { createClient } from '@supabase/supabase-js';
const url = import.meta.env.VITE_SUPABASE_URL, key = import.meta.env.VITE_SUPABASE_ANON_KEY;
export const sb = url && key ? createClient(url, key) : null;

export async function loadMasterDb() {
  if (!sb) throw new Error('Supabase belum dikonfigurasi');
  const out = {}; for (let from = 0; ; from += 1000) {
    const { data, error } = await sb.from('masa_manfaat').select('kode7,masa_manfaat_bulan').like('kode7', '1.3.%').order('kode7').range(from, from + 999);
    if (error) throw error; data.forEach((r) => (out[r.kode7] = r.masa_manfaat_bulan)); if (data.length < 1000) break;
  }
  if (!Object.keys(out).length) throw new Error('Tabel masa_manfaat kosong');
  return out;
}
export async function saveRun({ namaFile, meta, tahun, params, ringkasan, rows }) {
  const { data: run, error } = await sb.from('penyusutan_runs').insert({ nama_file: namaFile, pengguna_barang: meta.pengguna, tahun, params, ringkasan }).select('id').single();
  if (error) throw error;
  for (let i = 0; i < rows.length; i += 500) {
    const chunk = rows.slice(i, i + 500).map((data, j) => ({ run_id: run.id, urut: i + j, data }));
    const { error: e2 } = await sb.from('penyusutan_items').insert(chunk);
    if (e2) { await sb.from('penyusutan_runs').delete().eq('id', run.id); throw e2; }
  }
  return run.id;
}
export const listRuns = async () => (await sb.from('penyusutan_runs').select('id,nama_file,pengguna_barang,tahun,ringkasan,params,created_at').order('created_at', { ascending: false }).limit(20)).data || [];
export async function loadRun(id) {
  const rows = []; for (let from = 0; ; from += 1000) {
    const { data, error } = await sb.from('penyusutan_items').select('data').eq('run_id', id).order('urut').range(from, from + 999);
    if (error) throw error; rows.push(...data.map((d) => d.data)); if (data.length < 1000) break;
  }
  return rows;
}
export const deleteRun = (id) => sb.from('penyusutan_runs').delete().eq('id', id);
