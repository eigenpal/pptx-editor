// Slide show paint: the slide on show, painted by the same `paintSlide` as the editor and the
// fidelity harness (no second renderer), fitted to the screen with black letterbox bars.
// Hyperlinks open on click only, through `sanitizeHref`.

import { apply, invert, type LayoutItem } from '../layout/index.ts';
import { paintSlide } from '../output/index.ts';
import { relationshipsOf, sanitizeHref } from '../store/index.ts';
import type { DocumentSession } from './document-session.ts';
import { altTextAt } from './picture-state.ts';
import { textFrameMatrix } from './surface.ts';

/** A slide box on the screen, in CSS pixels on whole device pixels. @public */
export interface SlideshowFit {
  /** Left edge from the screen's left edge. */
  readonly left: number;
  /** Top edge from the screen's top edge. */
  readonly top: number;
  /** Slide width on the screen. */
  readonly width: number;
  /** Slide height on the screen. */
  readonly height: number;
}

/**
 * Fit a slide into a screen, keeping its aspect ratio, centered. The box edges fall on whole
 * device pixels (`dpr`), so the slide is sharp and the bars do not bleed. @public
 */
export function fitSlideToScreen(
  screenWidth: number,
  screenHeight: number,
  slideWidth: number,
  slideHeight: number,
  dpr = 1
): SlideshowFit {
  const ratio = dpr > 0 && Number.isFinite(dpr) ? dpr : 1;
  const sw = Math.max(0, screenWidth) * ratio;
  const sh = Math.max(0, screenHeight) * ratio;
  if (slideWidth <= 0 || slideHeight <= 0 || sw === 0 || sh === 0)
    return { left: 0, top: 0, width: 0, height: 0 };
  const scale = Math.min(sw / slideWidth, sh / slideHeight);
  const width = Math.min(sw, Math.round(slideWidth * scale));
  const height = Math.min(sh, Math.round(slideHeight * scale));
  const left = Math.floor((sw - width) / 2);
  const top = Math.floor((sh - height) / 2);
  return { left: left / ratio, top: top / ratio, width: width / ratio, height: height / ratio };
}

/** Options for `paintSlideshowSlide`. @public */
export interface SlideshowPaintOptions {
  /** Label for an element the editor does not render (chart, diagram, OLE). */
  readonly frameLabel?: (kind: string) => string;
}

let painted = 0;

/**
 * Paint slide `index` of a session for the show: an `<svg>` without editing hooks, with its text
 * and the alt text of its pictures for assistive technology. The caller sizes it (see
 * `fitSlideToScreen`). Returns null for a slide that does not exist. @public
 */
export function paintSlideshowSlide(
  doc: Document,
  session: DocumentSession,
  index: number,
  options: SlideshowPaintOptions = {}
): SVGSVGElement | null {
  const layout = session.layout(index);
  if (!layout) return null;
  painted += 1;
  const svg = paintSlide(doc, layout, {
    images: session.imageUrl,
    imageSize: session.imageSize,
    idPrefix: `pptx-show${painted}-`,
    hitTargets: false,
    altText: (part, id) => altTextAt(session.store.package, part, id),
    ...(options.frameLabel ? { frameLabel: options.frameLabel } : {}),
  });
  svg.setAttribute('width', '100%');
  svg.setAttribute('height', '100%');
  // A group, not an image: a screen reader reads the slide's text in paint order (the order of
  // `p:spTree`) and each picture's alt text, as on the editable slide.
  svg.setAttribute('role', 'group');
  return svg;
}

/**
 * The hyperlink under a slide point (EMU), as a URL safe to open, or null. Only text links
 * (`a:hlinkClick` on a run) to external targets count; `sanitizeHref` withholds any scheme
 * other than http(s), mailto, tel, and ftp. @public
 */
export function slideshowLinkAt(
  session: DocumentSession,
  index: number,
  x: number,
  y: number
): string | null {
  const layout = session.layout(index);
  if (!layout) return null;
  const rId = linkIn(layout.items, x, y);
  if (!rId) return null;
  const rel = relationshipsOf(session.store.package, rId.part).find((r) => r.id === rId.id);
  if (!rel?.external || !rel.externalTarget) return null;
  const safe = sanitizeHref(rel.externalTarget);
  return safe.ok && safe.href ? safe.href : null;
}

function linkIn(
  items: readonly LayoutItem[],
  x: number,
  y: number
): { part: string; id: string } | null {
  // Top-most first: later items paint over earlier ones.
  for (let i = items.length - 1; i >= 0; i -= 1) {
    const item = items[i];
    if (item.kind === 'group') {
      const inner = linkIn(item.children, x, y);
      if (inner) return inner;
      continue;
    }
    if (item.kind !== 'shape' || !item.text) continue;
    const m = textFrameMatrix(item);
    if (!m) continue;
    const p = apply(invert(m), x, y);
    for (const line of item.text.lines) {
      if (p.y < line.top || p.y > line.top + line.height) continue;
      for (const run of line.runs) {
        const id = run.props.hyperlink;
        if (!id || run.kind !== 'text' || run.x.length === 0) continue;
        if (p.x >= run.x[0] && p.x <= run.x[run.x.length - 1])
          return { part: item.element.partName, id };
      }
    }
  }
  return null;
}
