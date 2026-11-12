/**
 * Test-only: tiny PDF, workbook and deck built in memory, so the conversion
 * tests need no binary fixtures checked in.
 */
import ExcelJS from 'exceljs';
import { zipSync } from 'fflate';

function escapePdfString(s: string): string {
  return s.replace(/[\\()]/g, (c) => `\\${c}`);
}

/** A minimal valid PDF, one page per entry (Helvetica 12 pt); `''` is a page with no text layer. */
export function makePdf(pagesText: string[]): Buffer {
  const objects: string[] = [];
  const pageObj = (i: number) => 4 + i * 2;
  objects[1] = '<< /Type /Catalog /Pages 2 0 R >>';
  objects[2] = `<< /Type /Pages /Kids [${pagesText.map((_, i) => `${pageObj(i)} 0 R`).join(' ')}] /Count ${pagesText.length} >>`;
  objects[3] = '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>';
  pagesText.forEach((t, i) => {
    const stream = t === '' ? '' : `BT\n/F1 12 Tf\n72 720 Td\n(${escapePdfString(t)}) Tj\nET`;
    objects[pageObj(i)] =
      `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents ${pageObj(i) + 1} 0 R >>`;
    objects[pageObj(i) + 1] = `<< /Length ${Buffer.byteLength(stream, 'latin1')} >>\nstream\n${stream}\nendstream`;
  });
  let out = '%PDF-1.4\n';
  const offsets: number[] = [];
  for (let id = 1; id < objects.length; id++) {
    offsets[id] = Buffer.byteLength(out, 'latin1');
    out += `${id} 0 obj\n${objects[id]}\nendobj\n`;
  }
  const xref = Buffer.byteLength(out, 'latin1');
  out += `xref\n0 ${objects.length}\n0000000000 65535 f \n`;
  for (let id = 1; id < objects.length; id++) out += `${String(offsets[id]).padStart(10, '0')} 00000 n \n`;
  out += `trailer\n<< /Size ${objects.length} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(out, 'latin1');
}

/** A two-sheet workbook. */
export async function makeWorkbook(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('Budget').addRow(['item', 'cost']);
  wb.getWorksheet('Budget')!.addRow(['coffee', 3]);
  wb.addWorksheet('Notes').addRow(['remember']);
  return Buffer.from(await wb.xlsx.writeBuffer());
}

const P_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A one-slide deck titled "Quarterly plan". */
export function makeDeck(): Buffer {
  const enc = new TextEncoder();
  return Buffer.from(
    zipSync({
      'ppt/presentation.xml': enc.encode(
        `<p:presentation ${P_NS}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`,
      ),
      'ppt/_rels/presentation.xml.rels': enc.encode(
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slide" Target="slides/slide1.xml"/></Relationships>`,
      ),
      'ppt/slides/slide1.xml': enc.encode(
        `<p:sld ${P_NS}><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="t"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Quarterly plan</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
      ),
    }),
  );
}
