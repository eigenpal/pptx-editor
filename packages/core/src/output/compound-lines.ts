// Compound lines (design D5 of `fill-and-line-fidelity`). SVG strokes are centered, so a band at
// distances `a`…`b` from the center line is a stroke of width `2b` minus a stroke of width `2a`,
// built in a mask. On a closed path the outer bands are masked to the outside of the path and
// the inner bands are clipped to its inside. The specification does not give the band fractions,
// so they are an APPROXIMATION until reference PDFs exist (listed on the public fidelity page).

import { svg, num, type Defs } from './svg.ts';

/** Line and gap widths, outer side first, as fractions of the line width. */
export const COMPOUND_BANDS: Readonly<Record<string, readonly number[]>> = {
  dbl: [1 / 3, 1 / 3, 1 / 3],
  thickThin: [0.6, 0.2, 0.2],
  thinThick: [0.2, 0.2, 0.6],
  tri: [1 / 6, 1 / 6, 1 / 3, 1 / 6, 1 / 6],
};

/** A drawn band as distances from the center line, on one side (`a` < `b`). */
export interface Band {
  readonly a: number;
  readonly b: number;
}

/**
 * The drawn bands of a compound line of `width`, split by side of the center line. An open
 * path has no inside, so asymmetric types draw as `dbl` there.
 */
export function compoundBands(
  cmpd: string,
  width: number,
  closed: boolean
): { outer: Band[]; inner: Band[] } | null {
  let bands = Object.hasOwn(COMPOUND_BANDS, cmpd) ? COMPOUND_BANDS[cmpd] : undefined;
  if (!bands) return null;
  if (!closed && (cmpd === 'thickThin' || cmpd === 'thinThick')) bands = COMPOUND_BANDS.dbl;
  const outer: Band[] = [];
  const inner: Band[] = [];
  let p = 0;
  bands.forEach((f, i) => {
    const p0 = p;
    p += f;
    if (i % 2 === 1) return; // a gap
    // Offset from the center line: positive on the outer side.
    const o0 = (0.5 - p0) * width;
    const o1 = (0.5 - p) * width;
    if (o0 > 0) outer.push({ a: Math.max(0, o1), b: o0 });
    if (o1 < 0) inner.push({ a: Math.max(0, -o0), b: -o1 });
  });
  return { outer, inner };
}

/** Mask content for bands: widest first, so each black stroke cuts only its own band. */
function bandStrokes(
  doc: Document,
  d: string,
  bands: readonly Band[],
  stroke: Record<string, string | undefined>
): SVGElement[] {
  const out: SVGElement[] = [];
  const common = {
    d,
    fill: 'none',
    'stroke-linejoin': stroke['stroke-linejoin'],
    'stroke-linecap': stroke['stroke-linecap'],
    'stroke-miterlimit': stroke['stroke-miterlimit'],
    'stroke-dasharray': stroke['stroke-dasharray'],
  };
  for (const band of [...bands].sort((x, y) => y.b - x.b)) {
    out.push(svg(doc, 'path', { ...common, stroke: '#ffffff', 'stroke-width': num(2 * band.b) }));
    if (band.a > 0)
      out.push(svg(doc, 'path', { ...common, stroke: '#000000', 'stroke-width': num(2 * band.a) }));
  }
  return out;
}

/**
 * The SVG for a compound outline of path `d` in a `w`×`h` local box, or null when `cmpd` is a
 * single line. `stroke` is the single-line stroke (its paint, width, joins, caps, dashes).
 */
export function compoundStroke(
  defs: Defs,
  d: string,
  cmpd: string,
  stroke: Record<string, string | undefined>,
  w: number,
  h: number
): SVGElement | null {
  const width = Number(stroke['stroke-width']);
  const closed = /z\s*$/i.test(d);
  const bands = Number.isFinite(width) && width > 0 ? compoundBands(cmpd, width, closed) : null;
  if (!bands || stroke.stroke === 'none') return null;
  const doc = defs.doc;
  const grow = width + Math.max(w, h);
  const area = {
    x: num(-grow),
    y: num(-grow),
    width: num(w + 2 * grow),
    height: num(h + 2 * grow),
  };
  const painted = (maskUrl: string): SVGElement =>
    svg(doc, 'path', {
      d,
      fill: 'none',
      ...stroke,
      'stroke-width': num(width),
      mask: maskUrl,
    });
  const out = svg(doc, 'g', { 'data-compound': cmpd });
  const mask = (content: SVGElement[]): string =>
    defs.add(
      svg(doc, 'mask', { id: defs.id('cmpd'), maskUnits: 'userSpaceOnUse', ...area }, content)
    );
  if (!closed) {
    // Symmetric bands: both sides of the center line at once.
    const all = [...bands.outer, ...bands.inner];
    const merged = all.filter((b, i) => all.findIndex((c) => c.a === b.a && c.b === b.b) === i);
    out.appendChild(painted(mask(bandStrokes(doc, d, merged, stroke))));
    return out;
  }
  if (bands.outer.length)
    out.appendChild(
      painted(
        mask([
          ...bandStrokes(doc, d, bands.outer, stroke),
          // Remove the inside of the path: these bands lie outside it.
          svg(doc, 'path', { d, fill: '#000000', stroke: 'none' }),
        ])
      )
    );
  if (bands.inner.length) {
    // Keep the inside of the path with a mask, not a clipPath: Chromium can misplace a clipPath
    // edge in EMU coordinates under a translation (design D7 of `fill-fidelity`).
    const inside = mask([svg(doc, 'path', { d, fill: '#ffffff', stroke: 'none' })]);
    out.appendChild(
      svg(doc, 'g', { mask: inside }, [painted(mask(bandStrokes(doc, d, bands.inner, stroke)))])
    );
  }
  return out;
}
