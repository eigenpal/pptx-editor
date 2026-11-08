// The highlight of a text range, as rectangles per laid-out line in the text frame. Shared by
// the local selection and the selections of collaboration peers. Editing chrome only: never
// part of the slide paint. On a line in visual order (`Line.bidi`), the range covers one band
// for each stretch that stands together on screen (DOCX Editor `selection-rects.ts`).

import type { TextPoint } from '../contracts/index.ts';
import type { LayoutItem } from '../layout/index.ts';
import { SVG_NS } from '../output/index.ts';
import { textFrameMatrix } from './surface-geometry.ts';
import { bidiSelectionBands } from './visual-caret.ts';

/** One `<g>` of line rectangles for the range `a`–`b`, or nothing for an empty range. */
export function selectionRects(
  doc: Document,
  item: LayoutItem,
  a: TextPoint,
  b: TextPoint,
  className: string
): Element[] {
  if (item.kind !== 'shape' || !item.text) return [];
  const m = textFrameMatrix(item);
  if (!m) return [];
  const [from, to] =
    a.paragraph < b.paragraph || (a.paragraph === b.paragraph && a.offset <= b.offset)
      ? [a, b]
      : [b, a];
  if (from.paragraph === to.paragraph && from.offset === to.offset) return [];
  const g = doc.createElementNS(SVG_NS, 'g');
  g.setAttribute('transform', `matrix(${m.join(' ')})`);
  for (const line of item.text.lines) {
    if (line.paragraph < from.paragraph || line.paragraph > to.paragraph) continue;
    const s = line.paragraph === from.paragraph ? Math.max(from.offset, line.start) : line.start;
    const e = line.paragraph === to.paragraph ? Math.min(to.offset, line.end) : line.end;
    if (e <= s && !(line.paragraph < to.paragraph && line.end === line.start)) continue;
    const bands = line.bidi ? bidiSelectionBands(line, s, e) : [];
    if (bands.length === 0) {
      const x1 = line.caretX[s - line.start] ?? line.caretX[0];
      const x2 = line.caretX[Math.min(e - line.start, line.caretX.length - 1)] ?? x1;
      bands.push([Math.min(x1, x2), Math.max(x1, x2)]);
    }
    for (const [x1, x2] of bands) {
      const r = doc.createElementNS(SVG_NS, 'rect');
      r.setAttribute('class', className);
      r.setAttribute('x', String(x1));
      r.setAttribute('y', String(line.top));
      r.setAttribute('width', String(Math.max(x2 - x1, 6000)));
      r.setAttribute('height', String(line.height));
      g.append(r);
    }
  }
  return [g];
}
