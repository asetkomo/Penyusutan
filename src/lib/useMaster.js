import { loadMasterDb } from './supabase.js';
let p; // master dimuat sekali dan dipakai bersama oleh kedua tab
export const getMaster = () => (p ||= (async () => {
  try { return { masa: await loadMasterDb(), src: 'Supabase (tabel masa_manfaat)' }; }
  catch { const r = await fetch('/masa_manfaat.json'); return { masa: await r.json(), src: 'berkas bawaan aplikasi (Supabase belum aktif/kosong)' }; }
})());
