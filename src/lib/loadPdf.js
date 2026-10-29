import * as pdfjs from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;
// PDF dibaca sepenuhnya di browser; berkas tidak dikirim ke server mana pun.
export async function loadPages(file, onProgress) {
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise; const pages = [];
  for (let p = 1; p <= pdf.numPages; p++) { const tc = await (await pdf.getPage(p)).getTextContent(); pages.push(tc.items.map((i) => ({ str: i.str, x: i.transform[4], y: i.transform[5] }))); onProgress?.(p, pdf.numPages); }
  return pages;
}
