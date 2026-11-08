// Comment markers on the slide: one marker for each thread of the current slide, at the top right
// corner of the element it points at, at its place, or in the slide's top left corner. The review
// module's `pane.commentMarkers` picks the look, after DOCX Editor's review markers
// (docx-editor/packages/pro/src/react/review-icons.tsx and the `.docx-review__marker` and
// `.docx-review__badge` rules of packages/core/src/styles/editor.css, at adf0e178c): `icon`, the
// comment icon on a small card, or `initials`, a speech bubble in the author's color with the
// initials and the reply count. A press picks the thread and asks for the Comments tab of the
// side panel (`requestPane`). Pointer chrome: keyboard and screen reader users reach the same
// threads in the comments pane. Editing chrome only: the slide paint, thumbnails, the slide show,
// exports, and saved files never show it.

import type { CommentMarkerStyle, CommentThread } from '../contracts/index.ts';
import type { SlideLayout } from '../layout/index.ts';
import { SVG_NS } from '../output/index.ts';
import { findItem } from '../automation/commands/placement.ts';
import { commentAuthorSlot, commentInitials } from './comment-authors.ts';
import { ICON_PATHS } from './icon-paths.ts';

/** What the markers read from the editor. */
export interface CommentMarkerHost {
  /** The current slide's threads; empty without a review module. */
  threads(): readonly CommentThread[];
  /** The picked thread, or null. */
  active(): string | null;
  /** How a marker looks. */
  style(): CommentMarkerStyle;
  /** Pick a thread and show it in the comments pane. */
  pick(id: string): void;
}

/** A marker is 28 px high; markers at one place sit 4 px apart. */
const SIZE_PX = 28;
const GAP_PX = 4;
/** The initials badge: 26 px high, 26 to 36 px wide, inside the 28 px marker box. */
const BADGE_PX = 26;
const ICON_PX = 16;

/** Where a thread's marker goes, in slide units, before markers at one place spread out. */
function placeOf(thread: CommentThread, layout: SlideLayout): { x: number; y: number } {
  const a = thread.anchor;
  if (a.kind === 'point') return { x: a.x, y: a.y };
  if (a.kind === 'shape') {
    const item = findItem(layout.items, a.nodeId);
    if (item) {
      const [m0, m1, , , e, f] = item.matrix;
      return { x: m0 * item.element.xfrm.cx + e, y: m1 * item.element.xfrm.cx + f };
    }
  }
  return { x: 0, y: 0 };
}

/** A rectangle with its own radius at each corner: top left, top right, bottom right, bottom left. */
function roundedRect(
  x: number,
  y: number,
  w: number,
  h: number,
  [tl, tr, br, bl]: readonly [number, number, number, number]
): string {
  return (
    `M${x + tl} ${y}H${x + w - tr}Q${x + w} ${y} ${x + w} ${y + tr}` +
    `V${y + h - br}Q${x + w} ${y + h} ${x + w - br} ${y + h}` +
    `H${x + bl}Q${x} ${y + h} ${x} ${y + h - bl}V${y + tl}Q${x} ${y} ${x + tl} ${y}Z`
  );
}

export class CommentMarkers {
  constructor(private readonly host: CommentMarkerHost | undefined) {}

  /** The markers for `layout`; `px` is one screen pixel in slide units. */
  elements(doc: Document, layout: SlideLayout, px: number): Element[] {
    const threads = this.host?.threads() ?? [];
    if (threads.length === 0) return [];
    const active = this.host!.active();
    const style = this.host!.style();
    const size = SIZE_PX * px;
    const taken = new Map<string, number>();
    const out: Element[] = [];
    for (const thread of threads) {
      const place = placeOf(thread, layout);
      // Inside the slide, then to the right of the markers already at this place.
      const x0 = Math.min(Math.max(place.x, 0), Math.max(0, layout.slide.width - size));
      const y = Math.min(Math.max(place.y - size, 0), Math.max(0, layout.slide.height - size));
      const key = `${Math.round(x0 / size)}:${Math.round(y / size)}`;
      const n = taken.get(key) ?? 0;
      taken.set(key, n + 1);
      const x = x0 + n * (size + GAP_PX * px);
      out.push(this.#draw(doc, thread, style, x, y, px, thread.id === active));
    }
    return out;
  }

  #draw(
    doc: Document,
    thread: CommentThread,
    style: CommentMarkerStyle,
    x: number,
    y: number,
    px: number,
    active: boolean
  ): Element {
    const g = doc.createElementNS(SVG_NS, 'g');
    const state = [active ? 'is-active' : '', thread.resolved ? 'is-resolved' : '']
      .filter(Boolean)
      .join(' ');
    g.setAttribute('class', `pptx-surface__comment-marker ${state}`.trim());
    g.setAttribute('data-comment-marker', thread.id);
    g.setAttribute('data-marker', style);
    g.setAttribute('transform', `translate(${x} ${y})`);
    // The author's color slot, as on the thread's card in the comments pane.
    g.setAttribute(
      'style',
      `--doc-review-author-current: var(--doc-review-author-${commentAuthorSlot(thread.author)})`
    );
    const title = doc.createElementNS(SVG_NS, 'title');
    // File text: `textContent` only, never markup.
    title.textContent = thread.author ? `${thread.author}: ${thread.text}` : thread.text;
    g.append(title);
    if (style === 'initials') this.#badge(doc, g, thread, px);
    else this.#icon(doc, g, thread, px);
    return g;
  }

  /** `icon`: the comment icon on a small rounded card; a resolved thread shows a check mark. */
  #icon(doc: Document, g: Element, thread: CommentThread, px: number): void {
    const size = SIZE_PX * px;
    const box = doc.createElementNS(SVG_NS, 'path');
    box.setAttribute('d', roundedRect(0, 0, size, size, [6 * px, 6 * px, 6 * px, 6 * px]));
    box.setAttribute('class', 'pptx-surface__comment-marker-box');
    box.setAttribute('stroke-width', String(px));
    const offset = (SIZE_PX - ICON_PX) / 2;
    g.append(
      box,
      this.#glyph(doc, thread.resolved ? 'check' : 'comment', offset, offset, ICON_PX, px)
    );
  }

  /** `initials`: a speech bubble in the author's color, its tail at the anchor. */
  #badge(doc: Document, g: Element, thread: CommentThread, px: number): void {
    const initials = commentInitials(thread);
    const w = Math.min(36, Math.max(BADGE_PX, 8 + [...initials].length * 7)) * px;
    const h = BADGE_PX * px;
    const top = SIZE_PX * px - h;
    const radii = [13 * px, 13 * px, 13 * px, 3 * px] as const;
    // A ring in the surface color lifts the badge off the slide and off a neighboring marker.
    const ring = doc.createElementNS(SVG_NS, 'path');
    ring.setAttribute('d', roundedRect(0, top, w, h, radii));
    ring.setAttribute('class', 'pptx-surface__comment-marker-ring');
    ring.setAttribute('stroke-width', String(4 * px));
    const box = doc.createElementNS(SVG_NS, 'path');
    box.setAttribute('d', roundedRect(0, top, w, h, radii));
    box.setAttribute('class', 'pptx-surface__comment-marker-badge');
    box.setAttribute('stroke-width', String(px));
    g.append(ring, box);
    if (thread.resolved) {
      g.append(this.#glyph(doc, 'check', (w / px - 14) / 2, SIZE_PX - BADGE_PX + 6, 14, px));
    } else {
      const text = doc.createElementNS(SVG_NS, 'text');
      text.setAttribute('x', String(w / 2));
      text.setAttribute('y', String(top + h / 2 + 3.8 * px));
      text.setAttribute('font-size', String(10.5 * px));
      text.setAttribute('text-anchor', 'middle');
      text.setAttribute('class', 'pptx-surface__comment-marker-text');
      text.textContent = initials;
      g.append(text);
    }
    const replies = thread.replies.length;
    if (replies > 0) {
      const r = 7.5 * px;
      const cx = w + 6 * px - r;
      const cy = top - 5 * px + r;
      const dot = doc.createElementNS(SVG_NS, 'circle');
      dot.setAttribute('cx', String(cx));
      dot.setAttribute('cy', String(cy));
      dot.setAttribute('r', String(r));
      dot.setAttribute('class', 'pptx-surface__comment-marker-count');
      dot.setAttribute('stroke-width', String(px));
      const count = doc.createElementNS(SVG_NS, 'text');
      count.setAttribute('x', String(cx));
      count.setAttribute('y', String(cy + 3.2 * px));
      count.setAttribute('font-size', String(9 * px));
      count.setAttribute('text-anchor', 'middle');
      count.setAttribute('class', 'pptx-surface__comment-marker-count-text');
      count.textContent = replies > 9 ? '9+' : String(replies);
      g.append(dot, count);
    }
  }

  /** A Material Symbol at (`x`, `y`) screen pixels in the marker, `size` pixels wide. */
  #glyph(doc: Document, name: string, x: number, y: number, size: number, px: number): Element {
    const s = (size * px) / 960;
    const icon = doc.createElementNS(SVG_NS, 'path');
    icon.setAttribute('d', (ICON_PATHS[name] ?? []).join(''));
    // The path's view box is `0 -960 960 960`: its bottom edge is y = 0.
    icon.setAttribute('transform', `translate(${x * px} ${(y + size) * px}) scale(${s})`);
    icon.setAttribute('class', 'pptx-surface__comment-marker-icon');
    return icon;
  }

  /** A press on `target`: a marker picks its thread. Returns whether the press was on one. */
  press(target: Element): boolean {
    const id = target.closest?.('[data-comment-marker]')?.getAttribute('data-comment-marker');
    if (!id || !this.host) return false;
    this.host.pick(id);
    return true;
  }
}
