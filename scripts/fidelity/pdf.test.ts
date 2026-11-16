import { expect, test } from 'bun:test';
import { openPdf } from './pdf.ts';

/** A PDF of the given objects (numbered from 1, the first the catalog), with its cross-reference. */
function pdfOf(objects: readonly string[]): Uint8Array {
  let out = '%PDF-1.7\n';
  const offsets: number[] = [];
  objects.forEach((body, i) => {
    offsets.push(out.length);
    out += `${i + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xref = out.length;
  out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const o of offsets) out += `${String(o).padStart(10, '0')} 00000 n \n`;
  out += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return new TextEncoder().encode(out);
}

/** A page with three lines: plain text, text used as a clip, and text in a shading pattern. */
function page(): Uint8Array {
  const content = [
    'BT /F1 24 Tf 1 0 0 1 20 250 Tm (Plain words) Tj ET',
    'q 7 Tr BT /F1 24 Tf 1 0 0 1 20 200 Tm (Clipped words) Tj ET Q',
    'q /Pattern cs /P0 scn BT /F1 24 Tf 20 150 Td (Shaded words) Tj ET Q',
    'BT /F1 24 Tf 20 100 Td (Black words) Tj ET',
  ].join('\n');
  return pdfOf([
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 300] /Contents 4 0 R ' +
      '/Resources << /Font << /F1 5 0 R >> /Pattern << /P0 6 0 R >> >> >>',
    `<< /Length ${content.length} >>\nstream\n${content}\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    '<< /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 400 0] ' +
      '/Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> >> >>',
  ]);
}

test('text filled with a pattern or used as a clip is checked by pixels only', async () => {
  const pdf = await openPdf(page());
  try {
    expect((await pdf.lines(0)).map((l) => l.text)).toEqual(['Plain words', 'Black words']);
  } finally {
    await pdf.close();
  }
});
