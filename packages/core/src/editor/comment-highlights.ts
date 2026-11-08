// What a comment highlights on the slide, after DOCX Editor's comment bands (the
// `.docx-comment-band` rules of packages/core/src/styles/editor.css and their author hooks, at
// f189d6503). A thread on text draws a band over each line of its range; a thread on a shape,
// a picture, or another element draws a ring around the element. Both take the color of the
// author who started the thread, from the editor's review roster (`review-authors.ts`). A reply
// has no range of its own in either comment format: it shares its thread's, so a thread is one
// band or one ring, however many replies it has.
//
// A band is light at rest and stronger while its thread is open: the picked thread
// (`activeComment`), or the one the caret is in. It follows the caret, so it never animates. A
// ring is a stroke, not a wash: a wash disappears over a solid fill or a photo.
//
// The bands sit over the slide paint and blend with it, as a highlighter does: they multiply
// over light paper, so dark text keeps its color, and screen over dark paper, so light text
// keeps its color. The slide keeps its own colors in a dark chrome theme, so the paper is the
// slide's, not the chrome's: a band over light text is on dark paper. Editing chrome only: the
// slide paint, thumbnails, the slide show, exports, and saved files never show it.

import type { CommentThread, EditorSelection } from '../contracts/index.ts';
import type { LayoutItem, SlideLayout } from '../layout/index.ts';
import { SVG_NS } from '../output/index.ts';
import { findItem } from '../automation/commands/placement.ts';
import { comparePoints } from '../automation/commands/text-model.ts';
import { textFillAt, textFrameMatrix } from './surface-geometry.ts';
import { selectionRects } from './surface-text-selection.ts';
import { reviewAuthorHooks, type ReviewAuthorInfo } from './review-authors.ts';

/** What the highlights read from the editor. */
export interface CommentHighlightHost {
  threads(): readonly CommentThread[];
  active(): string | null;
  author(name: string): ReviewAuthorInfo | undefined;
}

/** The highlights of one slide: bands over light paper, bands over dark paper, and rings. */
export interface CommentHighlightLayers {
  readonly light: Element[];
  readonly dark: Element[];
  readonly rings: Element[];
}

/** A ring sits this many screen pixels outside its element, with corners this round. */
const RING_OUTSET_PX = 4;
const RING_RADIUS_PX = 3;

type TextAnchor = Extract<CommentThread['anchor'], { kind: 'text' }>;

/** Whether the text caret of `sel` is in the range of `anchor`. */
function caretIn(sel: EditorSelection, anchor: TextAnchor): boolean {
  if (sel.kind !== 'text' || sel.nodeId !== anchor.nodeId || sel.cell || sel.diagramNode)
    return false;
  return comparePoints(anchor.start, sel.focus) <= 0 && comparePoints(sel.focus, anchor.end) <= 0;
}

/** Relative luminance (WCAG 2), 0 to 1, of an sRGB color with 0 to 255 channels. */
function luminance({ r, g, b }: { r: number; g: number; b: number }): number {
  const lin = (c: number) => {
    const s = c / 255;
    return s <= 0.04045 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
}

/** Whether the range of `anchor` in `item` is light text, so it stands on dark paper. */
function onDarkPaper(item: LayoutItem, anchor: TextAnchor): boolean {
  if (item.kind !== 'shape' || !item.text) return false;
  const line = item.text.lines.find(
    (l) => l.paragraph === anchor.start.paragraph && anchor.start.offset < Math.max(l.end, 1)
  );
  const fill = line && textFillAt(line, anchor.start.offset + 1);
  return !!fill && luminance(fill) > 0.5;
}

/** Put the author hooks of `thread` and its state on `el`. */
function hook(el: Element, thread: CommentThread, host: CommentHighlightHost, open: boolean): void {
  el.setAttribute('data-comment', thread.id);
  const hooks = reviewAuthorHooks(thread.author, host.author(thread.author));
  if (hooks) {
    el.setAttribute('data-review-author', hooks.author);
    el.setAttribute('data-review-author-slot', String(hooks.slot));
    el.setAttribute('style', `--doc-review-author-current: ${hooks.color}`);
  }
  if (open) el.classList.add('is-active');
  if (thread.resolved) el.classList.add('is-resolved');
}

/** The band of a text thread: one rectangle per line, with a rule along its foot. */
function band(doc: Document, item: LayoutItem, anchor: TextAnchor, px: number): Element | null {
  const [g] = selectionRects(doc, item, anchor.start, anchor.end, 'pptx-surface__comment-band');
  if (!g) return null;
  g.setAttribute('class', 'pptx-surface__comment-band-group');
  // The rule is one screen pixel high in slide units; the text frame's matrix keeps its scale
  // for a shape that is not scaled, as every shape of a slide's top level is.
  for (const r of [...g.children]) {
    const rule = r.cloneNode(false) as Element;
    rule.setAttribute('class', 'pptx-surface__comment-band-rule');
    const y = Number(r.getAttribute('y')) + Number(r.getAttribute('height')) - px;
    rule.setAttribute('y', String(y));
    rule.setAttribute('height', String(px));
    r.after(rule);
  }
  return g;
}

/**
 * The ring of an element thread: a rounded rectangle just outside the element's box, 1.5 screen
 * pixels wide at rest and 2.5 while `open`.
 */
function ring(doc: Document, item: LayoutItem, px: number, open: boolean): Element {
  const { cx, cy } = item.element.xfrm;
  const [a, b, c, d] = item.matrix;
  // The item's own units: a shape in a scaled group has a scaled frame.
  const sx = Math.hypot(a, b) || 1;
  const sy = Math.hypot(c, d) || 1;
  const ox = (px * RING_OUTSET_PX) / sx;
  const oy = (px * RING_OUTSET_PX) / sy;
  const g = doc.createElementNS(SVG_NS, 'g');
  g.setAttribute('class', 'pptx-surface__comment-ring-group');
  g.setAttribute('transform', `matrix(${item.matrix.join(' ')})`);
  const rect = doc.createElementNS(SVG_NS, 'rect');
  rect.setAttribute('class', 'pptx-surface__comment-ring');
  rect.setAttribute('x', String(-ox));
  rect.setAttribute('y', String(-oy));
  rect.setAttribute('width', String(cx + 2 * ox));
  rect.setAttribute('height', String(Math.max(cy, 1) + 2 * oy));
  rect.setAttribute('rx', String((px * RING_RADIUS_PX) / sx));
  rect.setAttribute('ry', String((px * RING_RADIUS_PX) / sy));
  rect.setAttribute('stroke-width', String(((open ? 2.5 : 1.5) * px) / Math.sqrt(sx * sy)));
  g.append(rect);
  return g;
}

/**
 * The highlights of the threads of `layout`. `sel` is the editor's selection: a band is open
 * while the caret is in its range, and a ring while its element is the one selected element.
 * `px` is one screen pixel in slide units.
 */
export function commentHighlights(
  doc: Document,
  layout: SlideLayout,
  host: CommentHighlightHost | undefined,
  sel: EditorSelection,
  px: number
): CommentHighlightLayers {
  const out: CommentHighlightLayers = { light: [], dark: [], rings: [] };
  const threads = host?.threads() ?? [];
  if (threads.length === 0) return out;
  const active = host!.active();
  for (const thread of threads) {
    const anchor = thread.anchor;
    if (anchor.kind !== 'text' && anchor.kind !== 'shape') continue;
    const item = findItem(layout.items, anchor.nodeId);
    if (!item) continue;
    if (anchor.kind === 'text') {
      const g = band(doc, item, anchor, px);
      if (!g) continue;
      hook(g, thread, host!, thread.id === active || caretIn(sel, anchor));
      (onDarkPaper(item, anchor) ? out.dark : out.light).push(g);
    } else {
      const picked =
        sel.kind === 'shapes' && sel.nodeIds.length === 1 && sel.nodeIds[0] === anchor.nodeId;
      const open = thread.id === active || picked;
      const g = ring(doc, item, px, open);
      hook(g, thread, host!, open);
      out.rings.push(g);
    }
  }
  return out;
}

/**
 * Where the marker of a text thread goes, in slide units: at the top of the range's first
 * line, where the range ends on that line. Undefined when the range is not laid out.
 */
export function textAnchorPlace(
  layout: SlideLayout,
  anchor: TextAnchor
): { x: number; y: number } | undefined {
  const item = findItem(layout.items, anchor.nodeId);
  if (item?.kind !== 'shape' || !item.text) return undefined;
  const m = textFrameMatrix(item);
  if (!m) return undefined;
  const { start, end } = anchor;
  const line = item.text.lines.find(
    (l) =>
      (l.paragraph === start.paragraph && start.offset < Math.max(l.end, l.start + 1)) ||
      l.paragraph > start.paragraph
  );
  if (!line) return undefined;
  const to = line.paragraph === end.paragraph ? Math.min(end.offset, line.end) : line.end;
  let x = line.caretX[0] ?? 0;
  for (let i = 1; i <= to - line.start && i < line.caretX.length; i += 1)
    x = Math.max(x, line.caretX[i]!);
  return { x: m[0] * x + m[2] * line.top + m[4], y: m[1] * x + m[3] * line.top + m[5] };
}
