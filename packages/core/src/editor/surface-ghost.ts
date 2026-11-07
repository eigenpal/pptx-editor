// The ghost bullet: while the caret is in an empty list item, the surface shows the bullet or
// number the item will get, dimmed, so the user sees a new item before typing.
//
// An empty paragraph paints no bullet, so the slide paint (and the fidelity
// harness) never draws it. Layout records it as `Line.pendingBullet`; this module draws it in
// the editing overlay only, at the same glyph positions paint would use.

import type { EditorSelection } from '../contracts/index.ts';
import type { LayoutItem, Matrix } from '../layout/index.ts';
import { SVG_NS } from '../output/index.ts';
import { paintClusterRanges } from '../layout/text/paint-clusters.ts';

const EMU_PER_PT = 12_700;

/** The dimmed bullet for the caret's empty list item, or null when there is none. */
export function ghostBullet(
  doc: Document,
  item: LayoutItem,
  sel: Extract<EditorSelection, { kind: 'text' }>,
  /** Text frame space to slide space (`textFrameMatrix` in `surface.ts`). */
  m: Matrix | undefined
): SVGElement | null {
  if (item.kind !== 'shape' || !item.text) return null;
  if (sel.anchor.paragraph !== sel.focus.paragraph || sel.anchor.offset !== sel.focus.offset)
    return null;
  const line = item.text.lines.find((l) => l.paragraph === sel.focus.paragraph);
  const run = line?.pendingBullet;
  if (!line || !run || !m) return null;
  const g = doc.createElementNS(SVG_NS, 'g');
  g.setAttribute('class', 'pptx-surface__ghost-bullet');
  g.setAttribute('transform', `matrix(${m.join(' ')}) scale(${EMU_PER_PT})`);
  for (const piece of paintClusterRanges(run.text)) {
    const text = doc.createElementNS(SVG_NS, 'text');
    text.setAttribute(
      'x',
      run.x
        .slice(piece.start, piece.end)
        .map((v) => String(v / EMU_PER_PT))
        .join(' ')
    );
    text.setAttribute('y', String((line.baseline - run.rise) / EMU_PER_PT));
    text.setAttribute('font-family', `${run.face.id}, sans-serif`);
    text.setAttribute('font-size', String(run.sizePt));
    const fill = run.props.fill;
    if (fill.type === 'solid')
      text.setAttribute('fill', `rgb(${fill.color.r} ${fill.color.g} ${fill.color.b})`);
    text.setAttribute('xml:space', 'preserve');
    text.textContent = run.text.slice(piece.start, piece.end);
    g.appendChild(text);
  }
  return g;
}
