// Two-tone selection marks: under each selection outline, the marquee, and the caret, the
// overlay draws a wider light halo (`.pptx-surface__halo`, stroked with `--doc-selection-halo`,
// the same in light and dark chrome: it contrasts with the slide, not with the chrome). A blue
// mark with a light halo stays visible on a dark slide and on a light one, as the handles are
// (a light fill with a blue edge). A caret in text takes the text's color instead, and no halo.

/**
 * The overlay marks that get a halo, how much wider than the mark the halo is, and the halo's
 * modifier. The first match wins. The outline of one shape of several gets a narrow halo: it sits
 * just outside the shape, and a shape that touches it keeps its text clear.
 */
const HALOED: readonly (readonly [string, number, string?])[] = [
  ['pptx-surface__outline--member', 2, 'pptx-surface__halo--member'],
  ['pptx-surface__outline', 3],
  ['pptx-surface__group-outline', 3],
  ['pptx-surface__marquee', 3],
  ['pptx-surface__caret', 2.5],
];

function haloOf(mark: Element): Element | null {
  const names = (mark.getAttribute('class') ?? '').split(' ');
  // A caret in the text's own color contrasts with the slide already.
  if (names.includes('pptx-surface__caret--text')) return null;
  const kind = HALOED.find(([name]) => names.includes(name));
  if (!kind) return null;
  const halo = mark.cloneNode(false) as Element;
  const width = Number(mark.getAttribute('stroke-width') ?? '0');
  halo.setAttribute(
    'class',
    `pptx-surface__halo ${kind[2] ?? kind[0].replace('pptx-surface__', 'pptx-surface__halo--')}`
  );
  halo.setAttribute('stroke-width', String(width * kind[1]));
  // A solid halo under a dashed mark makes dark and light dashes: both show on any slide.
  halo.removeAttribute('stroke-dasharray');
  halo.removeAttribute('data-handle');
  return halo;
}

/**
 * Put a halo under every selection mark in `parts` (top level, or the first child of an
 * outline group). Returns the parts to draw, in order. @internal
 */
export function withHalos(parts: readonly Element[]): Element[] {
  const out: Element[] = [];
  for (const part of parts) {
    const halo = haloOf(part);
    if (halo) out.push(halo);
    const inner = part.firstElementChild;
    const innerHalo = part.localName === 'g' && inner ? haloOf(inner) : null;
    if (innerHalo && inner) part.insertBefore(innerHalo, inner);
    out.push(part);
  }
  return out;
}
