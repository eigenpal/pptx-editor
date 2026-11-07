// The animation order marks on the canvas: while `showAnimationOrder` is on (the Animations
// panel shows it), each animated shape of the main sequence carries a small tag with the click
// steps that play its effects (0 for effects that play when the slide starts).
//
// Editing chrome only: drawn in the surface overlay, never into the slide paint, thumbnails, or
// the fidelity harness. The Animations panel lists the same effects as text, for screen readers.

import type { AnimationEffect } from '../contracts/index.ts';
import type { SlideLayout } from '../layout/index.ts';
import { SVG_NS } from '../output/index.ts';
import { findItem } from '../automation/commands/placement.ts';

/** One mark: the shape node and its label. */
export interface AnimationMark {
  readonly nodeId: string;
  readonly label: string;
}

/** The marks of a slide's effects: one for each animated shape, with its steps in order. */
export function animationMarks(effects: readonly AnimationEffect[]): AnimationMark[] {
  const steps = new Map<string, number[]>();
  for (const e of effects) {
    if (e.nodeId === null || e.step === null) continue;
    const list = steps.get(e.nodeId) ?? [];
    if (!list.includes(e.step)) list.push(e.step);
    steps.set(e.nodeId, list);
  }
  return [...steps].map(([nodeId, list]) => ({ nodeId, label: list.join(', ') }));
}

/** Longest label drawn, in characters. */
const MAX_LABEL = 16;

/** The overlay elements of the marks. `px` is slide EMU per screen pixel. */
export function animationMarkElements(
  doc: Document,
  layout: SlideLayout,
  marks: readonly AnimationMark[],
  px: number
): SVGElement[] {
  const out: SVGElement[] = [];
  for (const mark of marks) {
    const item = findItem(layout.items, mark.nodeId);
    if (!item) continue;
    const text =
      mark.label.length > MAX_LABEL ? `${mark.label.slice(0, MAX_LABEL - 1)}…` : mark.label;
    // Drawn in screen pixels: the group scales them to slide units. A font size in slide units
    // (EMU) is past what a browser draws.
    const width = text.length * 7 + 10;
    // At the shape's top-left corner, outside the shape, so the tag covers no content.
    const x = item.matrix[4] - (width + 2) * px;
    const y = item.matrix[5];
    const g = doc.createElementNS(SVG_NS, 'g');
    g.setAttribute('class', 'pptx-surface__animation-mark');
    g.setAttribute('data-node', mark.nodeId);
    g.setAttribute('transform', `translate(${x} ${y}) scale(${px})`);
    const box = doc.createElementNS(SVG_NS, 'rect');
    box.setAttribute('width', String(width));
    box.setAttribute('height', '16');
    box.setAttribute('rx', '3');
    box.setAttribute('stroke-width', '1.5');
    const t = doc.createElementNS(SVG_NS, 'text');
    t.setAttribute('x', '5');
    t.setAttribute('y', '12');
    t.setAttribute('font-size', '11');
    t.textContent = text;
    g.append(box, t);
    out.push(g);
  }
  return out;
}
