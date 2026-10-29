// Penyusutan Gedung & Bangunan mengikuti kolom AA..AP pada diskop_penyusutan_2026.xlsx (rumus sama persis).
const n = (v) => (Number.isFinite(+v) && v !== null && v !== '' ? +v : 0);
const r2 = (v) => Math.sign(v) * Math.round(Math.abs(v) * 100) / 100;

export function hitungGedung(it, masa, inp, tahun) {
  const V = n(it.nilai), AA = n(masa), AB = n(inp.akm), AC = n(inp.nb), AD = n(inp.sisa), AE = n(inp.kap), AG = n(inp.tambahan);
  const kap = AE > 0; const mKap = inp.tglKap ? +inp.tglKap.slice(5, 7) : 1; const yKap = inp.tglKap ? +inp.tglKap.slice(0, 4) : 1900;
  const AH = kap ? AE / V : null, AI = kap ? mKap - 1 : 12, AJ = kap ? 13 - mKap : 0;
  const AK = kap ? Math.max(0, Math.min(AD - 12 + AG, AA - AJ)) : Math.max(0, AD - 12);
  const dep = kap ? (AD > 0 ? (AC / AD) * AI : 0) + (AK > 0 ? ((AC + AE) / AK) * AJ : 0) : AD > 0 ? (AC / AD) * Math.min(12, AD) : 0;
  const AL = Math.min(AC + AE, dep), AM = AB + AL, AN = V - AM, AO = r2(V - (AB + AC + AE));
  const AP = Math.abs(AO) > 1 ? 'CEK NILAI PEROLEHAN' : kap && yKap !== tahun ? 'CEK TGL KAPITALISASI' : kap && AG === 0 ? 'CEK TAMBAHAN MASA MANFAAT' : 'OK';
  return { AA, AB, AC, AD, AE, AF: inp.tglKap || null, AG: inp.tambahan ?? null, AH, AI, AJ, AK, AL, AM, AN, AO, AP };
}
