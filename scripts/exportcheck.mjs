// Menulis hasil ekspor ke berkas (untuk pemeriksaan dengan LibreOffice).  node scripts/exportcheck.mjs /tmp/hasil.xlsx
import fs from 'fs'; import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { extractRecords } from '../src/lib/parsePdf.js'; import { hitung } from '../src/lib/calc.js'; import { audit } from '../src/lib/audit.js'; import { buildWorkbook } from '../src/lib/exportXlsx.js';
const D = process.env.DATA_DIR || '/mnt/user-data/uploads/';
const pdf = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(D + 'dpu.pdf')), verbosity: 0 }).promise;
const pages = []; for (let p = 1; p <= pdf.numPages; p++) { const tc = await (await pdf.getPage(p)).getTextContent(); pages.push(tc.items.map((i) => ({ str: i.str, x: i.transform[4], y: i.transform[5] }))); }
const { rows } = extractRecords(pages); const masa = JSON.parse(fs.readFileSync(new URL('../public/masa_manfaat.json', import.meta.url)));
const hasil = rows.filter((r) => r.type === 'item').map((r) => hitung(r, masa[r.kode], { tahun: 2024 }));
const wb = await buildWorkbook({ rows, hasil, temuan: audit(rows, hasil, { tahun: 2024 }), masa, tahun: 2024, meta: {} });
await wb.xlsx.writeFile(process.argv[2] || '/tmp/hasil.xlsx'); console.log('ditulis', process.argv[2] || '/tmp/hasil.xlsx');
