// Gradient paint servers: linear gradients,
// where `scaled` stretches the unit-square gradient with the box; path gradients (`circle`,
// `rect`, `shape`) from the focus to the boundary, linear in the distance; and the color rule
// for two-stop gradients. Paint servers use `userSpaceOnUse` over the fill box and set no
// `gradientTransform`: gradient text sets it to place the gradient (`paint-text.ts`).

import {
  toCss,
  type GradientStop,
  type RelRect,
  type ResolvedFill,
  type Rgba,
} from '../cascade/index.ts';
import { Defs, num, svg } from './svg.ts';

type LinearFill = Extract<ResolvedFill, { type: 'linear' }>;
type PathFill = Extract<ResolvedFill, { type: 'path' }>;

/** A rectangle in the fill's local space. */
export interface Box {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

/** Steps of a resampled two-stop gradient, so 33 stops (design D8: fixed, never from the file). */
export const TWO_STOP_STEPS = 32;
/** How far the circle of a circle gradient reaches past a focus at or outside it (measured). */
const FOCUS_MARGIN = 1.005;
/** The fewest and the most bands of a `path="shape"` gradient (design D8: fixed bounds). */
export const MIN_SHAPE_BANDS = 48;
export const MAX_SHAPE_BANDS = 128;

/**
 * The band count of a shape gradient: enough that the steepest color change moves at most about
 * two levels per band (a two-stop curve is twice as steep at its ends), within the fixed bounds.
 */
export function shapeBands(stops: readonly GradientStop[]): number {
  let step = 0;
  for (let i = 1; i < stops.length; i += 1) {
    const a = stops[i - 1].color;
    const b = stops[i].color;
    const span = stops[i].pos - stops[i - 1].pos;
    if (span <= 0) continue;
    const delta = Math.max(Math.abs(a.r - b.r), Math.abs(a.g - b.g), Math.abs(a.b - b.b));
    step = Math.max(step, ((bends(stops) ? 2 : 1) * delta) / span);
  }
  return Math.min(MAX_SHAPE_BANDS, Math.max(MIN_SHAPE_BANDS, Math.ceil(step / 2)));
}

/** The fill box of a `w`×`h` shape. */
const fillBox = (w: number, h: number): Box => ({ x: 0, y: 0, w, h });

/**
 * The tile rectangle of a gradient (`CT_GradientFillProperties/tileRect`): the fill box inset by
 * the tile insets; negative insets extend it. An empty rectangle falls back to the fill box.
 */
export function tileBox(tile: RelRect | undefined, w: number, h: number): Box {
  if (!tile) return fillBox(w, h);
  const bw = w * (1 - tile.l - tile.r);
  const bh = h * (1 - tile.t - tile.b);
  return bw > 0 && bh > 0 ? { x: tile.l * w, y: tile.t * h, w: bw, h: bh } : fillBox(w, h);
}

/** Mix two colors in sRGB, alpha included. */
function mix(a: Rgba, b: Rgba, f: number): Rgba {
  const k = Math.min(1, Math.max(0, f));
  return {
    r: a.r + (b.r - a.r) * k,
    g: a.g + (b.g - a.g) * k,
    b: a.b + (b.b - a.b) * k,
    a: a.a + (b.a - a.a) * k,
  };
}

/**
 * The color between the two stops of a two-stop gradient (design D2). Each sRGB channel follows a
 * quadratic curve that bends toward the larger of its two end values:
 * (1 − u)² a + 2u(1 − u) max(a, b) + u² b. Alpha stays linear. Measured against the reference PDFs
 * on `fixtures/gradients.pdf`, `fixtures/fills-lines.pdf`, and `fixtures/fill-fidelity.pdf`.
 */
export function twoStopColor(a: Rgba, b: Rgba, u: number): Rgba {
  const k = Math.min(1, Math.max(0, u));
  const bend = (x: number, y: number): number =>
    (1 - k) ** 2 * x + 2 * k * (1 - k) * Math.max(x, y) + k * k * y;
  return { r: bend(a.r, b.r), g: bend(a.g, b.g), b: bend(a.b, b.b), a: a.a + (b.a - a.a) * k };
}

/**
 * Whether a gradient's colors bend (design D2): exactly two stops, at 0 and 100 %. A stop list
 * that does not reach an end extends with its end color, which makes three stops,
 * and those interpolate linearly (a smoke-set PDF: stops at 0 and 88 % are linear).
 */
export const bends = (stops: readonly GradientStop[]): boolean =>
  stops.length === 2 && stops[0].pos <= 0 && stops[1].pos >= 1;

/**
 * The color at stop position `t` (0–1) between sorted stops: the two-stop curve, or sRGB
 * interpolation between neighboring stops when there are three or more.
 */
export function colorAt(stops: readonly GradientStop[], t: number): Rgba {
  if (t <= stops[0].pos) return stops[0].color;
  for (let i = 1; i < stops.length; i += 1) {
    const b = stops[i];
    if (t > b.pos) continue;
    const a = stops[i - 1];
    if (b.pos <= a.pos) return b.color;
    const u = (t - a.pos) / (b.pos - a.pos);
    return bends(stops) ? twoStopColor(a.color, b.color, u) : mix(a.color, b.color, u);
  }
  return stops[stops.length - 1].color;
}

/**
 * The stops a paint server draws. SVG interpolates linearly, so a two-stop gradient is
 * resampled at `TWO_STOP_STEPS` positions on its curve; three or more stops stay as they are.
 */
export function paintStops(stops: readonly GradientStop[]): GradientStop[] {
  if (!bends(stops)) return [...stops];
  const [a, b] = stops;
  const out: GradientStop[] = [];
  for (let n = 0; n <= TWO_STOP_STEPS; n += 1) {
    const u = n / TWO_STOP_STEPS;
    out.push({ pos: a.pos + (b.pos - a.pos) * u, color: twoStopColor(a.color, b.color, u) });
  }
  return out;
}

/** A stop offset with four decimals: `num` keeps two, too few for a resampled curve. */
const offset = (pos: number): string =>
  Number.isFinite(pos) ? String(Math.round(pos * 10_000) / 10_000) : '0';

/** A band scale with six decimals, so up to `MAX_SHAPE_BANDS` bands keep distinct scales. */
const fine = (v: number): string =>
  Number.isFinite(v) ? String(Math.round(v * 1_000_000) / 1_000_000) : '0';

/** `<stop>` elements for sorted stops. */
export function stopElements(doc: Document, stops: readonly GradientStop[]): SVGElement[] {
  return stops.map((s) =>
    svg(doc, 'stop', {
      offset: offset(s.pos),
      'stop-color': toCss({ ...s.color, a: 1 }),
      'stop-opacity': s.color.a < 1 ? num(s.color.a) : undefined,
    })
  );
}

/**
 * The end points of a linear gradient over a box (design D1). `a:lin/@ang` is the direction of
 * color change, clockwise from the x axis (`CT_LinearShadeProperties`). With `scaled`, the
 * unit-square gradient stretches with the box: its isolines (−sin a, cos a) become
 * (−w sin a, h cos a), so the color changes along (h cos a, w sin a). The ends are the
 * projections of the farthest corners.
 */
export function linearEnds(
  angle: number,
  scaled: boolean,
  box: Box
): { x1: number; y1: number; x2: number; y2: number } {
  const r = (angle * Math.PI) / 180;
  let dx = Math.cos(r);
  let dy = Math.sin(r);
  if (scaled && box.w > 0 && box.h > 0) {
    const nx = box.h * dx;
    const ny = box.w * dy;
    const len = Math.hypot(nx, ny) || 1;
    dx = nx / len;
    dy = ny / len;
  }
  const half = (Math.abs(box.w * dx) + Math.abs(box.h * dy)) / 2;
  const cx = box.x + box.w / 2;
  const cy = box.y + box.h / 2;
  return { x1: cx - dx * half, y1: cy - dy * half, x2: cx + dx * half, y2: cy + dy * half };
}

/** A linear gradient paint server over a `w`×`h` fill box. */
export function linearGradient(
  defs: Defs,
  fill: LinearFill,
  stops: readonly GradientStop[],
  w: number,
  h: number
): string {
  const ends = linearEnds(fill.angle, fill.scaled, fillBox(w, h));
  const el = svg(
    defs.doc,
    'linearGradient',
    {
      id: defs.id('lin'),
      gradientUnits: 'userSpaceOnUse',
      x1: num(ends.x1),
      y1: num(ends.y1),
      x2: num(ends.x2),
      y2: num(ends.y2),
    },
    stopElements(defs.doc, paintStops(stops))
  );
  return defs.add(el);
}

/** A gray as dark as `1 − a`: stop alpha as mask luminance. */
const gray = (a: number): string => toCss({ r: a * 255, g: a * 255, b: a * 255, a: 1 });

/** The color of an opaque `#rrggbb` string, for stops built through a paint mapping. */
function parseCss(css: string): Rgba {
  const v = Number.parseInt(css.slice(1), 16);
  return { r: (v >> 16) & 255, g: (v >> 8) & 255, b: v & 255, a: 1 };
}

/** A luminance mask over the fill box (with room around it) that holds `content`. */
function alphaMask(defs: Defs, w: number, h: number, content: readonly SVGElement[]): string {
  const id = defs.id('alpha');
  defs.add(
    svg(
      defs.doc,
      'mask',
      {
        id,
        maskUnits: 'userSpaceOnUse',
        x: num(-w),
        y: num(-h),
        width: num(3 * w),
        height: num(3 * h),
      },
      content
    )
  );
  return `url(#${id})`;
}

/** The focus rectangle of a path gradient, in local space. */
function focusRect(fill: PathFill, box: Box): { l: number; t: number; r: number; b: number } {
  const f = fill.fillToRect;
  const l = box.x + f.l * box.w;
  const t = box.y + f.t * box.h;
  return {
    l,
    t,
    r: Math.max(l, box.x + box.w - f.r * box.w),
    b: Math.max(t, box.y + box.h - f.b * box.h),
  };
}

/**
 * The circle and focal point of a circular path gradient over a `w`×`h` fill box (see
 * `circleGradient`): the last stop is on the circle (`cx`, `cy`, `r`), the first at (`fx`, `fy`).
 */
export function circleGeometry(
  fill: PathFill,
  w: number,
  h: number
): { cx: number; cy: number; r: number; fx: number; fy: number } {
  const tile = tileBox(fill.tileRect, w, h);
  const cx = tile.x + tile.w / 2;
  const cy = tile.y + tile.h / 2;
  const f = focusRect(fill, fillBox(w, h));
  const fx = (f.l + f.r) / 2;
  const fy = (f.t + f.b) / 2;
  const r =
    Math.max(Math.hypot(tile.w, tile.h) / 2, FOCUS_MARGIN * Math.hypot(fx - cx, fy - cy)) || 1;
  return { cx, cy, r, fx, fy };
}

/**
 * A circular path gradient (`CT_PathShadeProperties`, `ST_PathShadeType` `circle`; design D2):
 * an SVG focal radial gradient. The last stop is on the circle through the corners of the tile
 * rectangle (the fill box by default), centered on it; the first stop is at the focal point, the
 * center of `a:fillToRect` on the fill box. A focus at or outside that circle stays where it is,
 * and the circle grows to 1.005 times the focus distance, as the reference PDF
 * shadings show. Measured on `fixtures/fill-fidelity.pdf` for a centered, an off-center, a corner, and
 * an outside focus.
 */
export function circleGradient(
  defs: Defs,
  fill: PathFill,
  stops: readonly GradientStop[],
  w: number,
  h: number
): string {
  const { cx, cy, r, fx, fy } = circleGeometry(fill, w, h);
  const el = svg(
    defs.doc,
    'radialGradient',
    {
      id: defs.id('rad'),
      gradientUnits: 'userSpaceOnUse',
      cx: num(cx),
      cy: num(cy),
      r: num(r),
      fx: num(fx),
      fy: num(fy),
    },
    stopElements(defs.doc, paintStops(stops))
  );
  return defs.add(el);
}

/**
 * A rectangular path gradient (`ST_PathShadeType` `rect`; design D1 of `fill-and-line-fidelity`):
 * concentric rectangles from the focus rectangle (the first stop) to the edge of the tile rectangle
 * (the last stop; the fill box by default), from four trapezoids with linear gradients, in a
 * pattern over the fill box.
 */
export function rectGradient(
  defs: Defs,
  fill: PathFill,
  stops: readonly GradientStop[],
  w: number,
  h: number
): string {
  const doc = defs.doc;
  const { l, t, r, b } = focusRect(fill, fillBox(w, h));
  const tile = tileBox(fill.tileRect, w, h);
  const x0 = Math.min(tile.x, l);
  const y0 = Math.min(tile.y, t);
  const x1 = Math.max(tile.x + tile.w, r);
  const y1 = Math.max(tile.y + tile.h, b);
  const painted = paintStops(stops);
  const first = stops[0].color;
  const last = stops[stops.length - 1].color;
  // The construction paints opaque; stop alpha goes into a luminance mask of the same shapes,
  // because the base under the trapezoids would otherwise count the alpha twice.
  const construction = (paint: (c: Rgba) => string): SVGElement[] => {
    const gradient = (gx1: number, gy1: number, gx2: number, gy2: number): string =>
      defs.add(
        svg(
          doc,
          'linearGradient',
          {
            id: defs.id('rect'),
            gradientUnits: 'userSpaceOnUse',
            x1: num(gx1),
            y1: num(gy1),
            x2: num(gx2),
            y2: num(gy2),
          },
          stopElements(
            doc,
            painted.map((s) => ({ pos: s.pos, color: parseCss(paint(s.color)) }))
          )
        )
      );
    const poly = (pts: readonly [number, number][], fillRef: string): SVGElement =>
      svg(doc, 'polygon', {
        points: pts.map(([x, y]) => `${num(x)},${num(y)}`).join(' '),
        fill: fillRef,
      });
    return [
      // The base hides antialiasing seams between the trapezoids.
      svg(doc, 'rect', { width: num(w), height: num(h), fill: paint(last) }),
      poly(
        [
          [x0, y0],
          [x1, y0],
          [r, t],
          [l, t],
        ],
        gradient(0, t, 0, y0)
      ),
      poly(
        [
          [l, b],
          [r, b],
          [x1, y1],
          [x0, y1],
        ],
        gradient(0, b, 0, y1)
      ),
      poly(
        [
          [x0, y0],
          [l, t],
          [l, b],
          [x0, y1],
        ],
        gradient(l, 0, x0, 0)
      ),
      poly(
        [
          [r, t],
          [x1, y0],
          [x1, y1],
          [r, b],
        ],
        gradient(r, 0, x1, 0)
      ),
      ...(r > l && b > t
        ? [
            svg(doc, 'rect', {
              x: num(l),
              y: num(t),
              width: num(r - l),
              height: num(b - t),
              fill: paint(first),
            }),
          ]
        : []),
    ];
  };
  const opaque = construction((c) => toCss({ ...c, a: 1 }));
  const content = stops.some((s) => s.color.a < 1)
    ? [
        svg(
          doc,
          'g',
          {
            mask: alphaMask(
              defs,
              w,
              h,
              construction((c) => gray(c.a))
            ),
          },
          opaque
        ),
      ]
    : opaque;
  const pattern = svg(
    doc,
    'pattern',
    {
      id: defs.id('pathrect'),
      patternUnits: 'userSpaceOnUse',
      x: 0,
      y: 0,
      width: num(w),
      height: num(h),
    },
    content
  );
  return defs.add(pattern);
}

/**
 * A shape path gradient (`ST_PathShadeType` `shape`; design D3): `shapeBands` copies of the
 * outline, scaled about the focus center from 1 down, each filled with the color at the middle of
 * its band (linear in the scale), outside in. The outline is defined once; each band is a `<use>`.
 */
export function shapeGradient(
  defs: Defs,
  fill: PathFill,
  stops: readonly GradientStop[],
  w: number,
  h: number,
  fillPath: string
): string {
  const doc = defs.doc;
  const f = focusRect(fill, fillBox(w, h));
  const cx = (f.l + f.r) / 2;
  const cy = (f.t + f.b) / 2;
  const count = shapeBands(stops);
  const pathId = defs.id('fillpath');
  // The bands follow the shape's own fill rule (paint fills shapes with `evenodd`).
  defs.add(svg(doc, 'path', { id: pathId, d: fillPath, 'fill-rule': 'evenodd' }));
  // Bands nest, so each paints opaque; stop alpha goes into a luminance mask of the same bands.
  const bands = (paint: (c: Rgba) => string): SVGElement[] => {
    const out: SVGElement[] = [];
    for (let i = 0; i < count; i += 1) {
      const outer = 1 - i / count;
      const c = colorAt(stops, outer - 0.5 / count);
      out.push(
        svg(doc, 'use', {
          href: `#${pathId}`,
          transform: `translate(${num(cx)} ${num(cy)}) scale(${fine(outer)}) translate(${num(-cx)} ${num(-cy)})`,
          fill: paint(c),
        })
      );
    }
    return out;
  };
  const opaque = bands((c) => toCss({ ...c, a: 1 }));
  const content = stops.some((s) => s.color.a < 1)
    ? [
        svg(
          doc,
          'g',
          {
            mask: alphaMask(
              defs,
              w,
              h,
              bands((c) => gray(c.a))
            ),
          },
          opaque
        ),
      ]
    : opaque;
  const pattern = svg(
    doc,
    'pattern',
    {
      id: defs.id('pathshape'),
      patternUnits: 'userSpaceOnUse',
      x: 0,
      y: 0,
      width: num(w),
      height: num(h),
    },
    content
  );
  return defs.add(pattern);
}
