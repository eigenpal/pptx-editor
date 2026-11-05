// Collaboration peers on the canvas, after DOCX Editor's remote-selection and remote-caret
// painters (`surface-remote-selection.ts`, `surface-remote-caret.ts`): a colored outline on the
// shapes a peer selected, and a colored range and caret with the participant's name in text.
//
// Editing chrome only: drawn in the surface overlay, never into the slide paint, thumbnails,
// or the fidelity harness. The color is the participant's checked hex color, set as
// `--pptx-remote-color`; without one, CSS falls back to the theme accent.

import type { LayoutItem, SlideLayout } from '../layout/index.ts';
import { SVG_NS } from '../output/index.ts';
import type { PaintedPresence } from './collaboration-presence.ts';
import { findItem } from '../automation/commands/placement.ts';
import { caretAt } from './surface-geometry.ts';
import { selectionRects } from './surface-text-selection.ts';
import { textItemOf } from './table-cells.ts';

/** Longest name shown in a caret label, in characters. */
const MAX_LABEL = 24;

function outline(doc: Document, item: LayoutItem, px: number): SVGElement {
  const g = doc.createElementNS(SVG_NS, 'g');
  g.setAttribute('transform', `matrix(${item.matrix.join(' ')})`);
  const rect = doc.createElementNS(SVG_NS, 'rect');
  rect.setAttribute('class', 'pptx-surface__remote-outline');
  rect.setAttribute('width', String(item.element.xfrm.cx));
  rect.setAttribute('height', String(Math.max(item.element.xfrm.cy, 1)));
  rect.setAttribute('stroke-width', String(px * 2));
  g.append(rect);
  return g;
}

function label(doc: Document, name: string, x: number, y: number, px: number): SVGElement {
  const text = name.length > MAX_LABEL ? `${name.slice(0, MAX_LABEL - 1)}…` : name;
  const g = doc.createElementNS(SVG_NS, 'g');
  g.setAttribute('class', 'pptx-surface__remote-label');
  const height = px * 16;
  const box = doc.createElementNS(SVG_NS, 'rect');
  box.setAttribute('x', String(x));
  box.setAttribute('y', String(y - height));
  // No layout pass is available here; 7 px per character fits the 11 px UI font.
  box.setAttribute('width', String(px * (text.length * 7 + 10)));
  box.setAttribute('height', String(height));
  box.setAttribute('rx', String(px * 3));
  const t = doc.createElementNS(SVG_NS, 'text');
  t.setAttribute('x', String(x + px * 5));
  t.setAttribute('y', String(y - px * 4));
  t.setAttribute('font-size', String(px * 11));
  t.textContent = text;
  g.append(box, t);
  return g;
}

/** The overlay elements for every peer on this slide. `px` is slide EMU per screen pixel. */
export function presenceElements(
  doc: Document,
  layout: SlideLayout,
  presences: readonly PaintedPresence[],
  px: number
): SVGElement[] {
  const out: SVGElement[] = [];
  for (const peer of presences) {
    const g = doc.createElementNS(SVG_NS, 'g');
    g.setAttribute('class', 'pptx-surface__presence');
    g.setAttribute('data-actor', peer.actorId);
    if (peer.color) g.setAttribute('style', `--pptx-remote-color: ${peer.color}`);
    const sel = peer.selection;
    if (sel.kind === 'shapes') {
      for (const id of sel.nodeIds) {
        const item = findItem(layout.items, id);
        if (item) g.append(outline(doc, item, px));
      }
      const first = sel.nodeIds.map((id) => findItem(layout.items, id)).find((i) => i);
      if (first) g.append(label(doc, peer.name, first.matrix[4], first.matrix[5], px));
    } else if (sel.kind === 'text') {
      const item = textItemOf(layout, sel);
      const caret = item ? caretAt(item, sel.focus) : undefined;
      if (!item || !caret) continue;
      for (const r of selectionRects(
        doc,
        item,
        sel.anchor,
        sel.focus,
        'pptx-surface__remote-selection'
      ))
        g.append(r);
      const line = doc.createElementNS(SVG_NS, 'line');
      line.setAttribute('class', 'pptx-surface__remote-caret');
      line.setAttribute('x1', String(caret.x1));
      line.setAttribute('y1', String(caret.y1));
      line.setAttribute('x2', String(caret.x2));
      line.setAttribute('y2', String(caret.y2));
      line.setAttribute('stroke-width', String(px * 2));
      g.append(line, label(doc, peer.name, caret.x1, caret.y1, px));
    }
    if (g.childNodes.length > 0) out.push(g);
  }
  return out;
}
