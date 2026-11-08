// Fill and outline paint: solid, linear and path gradients (`paint-gradient.ts`), pictures
// (stretch with srcRect and fillRect, tile), patterns, and outline strokes. Paint servers are
// defined in the shape's local coordinate space (userSpaceOnUse over the shape box), so rotation
// and flips carry them. Pattern tiles are the exception: they stay square to the slide.

import {
  toCss,
  type EffectAlign,
  type Rgba,
  type ResolvedFill,
  type ResolvedLine,
  type BlipEffect,
} from '../cascade/index.ts';
import {
  dashArray,
  invert,
  multiply,
  strokeCap,
  strokeWidth,
  type Matrix,
} from '../layout/index.ts';
import { ALIGN_POINTS, lumTransfer } from './effect-transforms.ts';
import { circleGradient, linearGradient, rectGradient, shapeGradient } from './paint-gradient.ts';
import { PATTERN_PIXEL, patternBitmap } from './preset-patterns.ts';
import { Defs, num, svg } from './svg.ts';

/** Resolves a picture fill's relationship id to a displayable URL (never a remote fetch). */
export type ImageUrlPort = (sourcePart: string, rId: string) => string | undefined;

/**
 * Resolves a picture fill's relationship id to the image's natural size in EMU, for tiles. Use
 * the header's resolution, or 72 DPI when it states none (measured on reference PDFs)
 * (`DocumentSession.imageSize`). The store's `imageSizeEmu` defaults to 96 DPI, so a port built
 * on it passes `{ ...size, dpiX: size.dpiX ?? 72, dpiY: size.dpiY ?? 72 }`.
 */
export type ImageSizePort = (
  sourcePart: string,
  rId: string
) => { readonly cx: number; readonly cy: number } | undefined;

/** What paint needs from its host. */
export interface PaintContext {
  readonly defs: Defs;
  readonly images: ImageUrlPort;
  /** Natural image sizes, for tiled pictures. Without it, a tile is the shape's size. */
  readonly imageSize?: ImageSizePort;
  /** The part whose relationships resolve this element's `r:embed` ids. */
  readonly sourcePart: string;
  /**
   * The path being filled (SVG path data), for `path="shape"` gradients. Paint sets it for each
   * filled path; a host leaves it unset.
   */
  readonly fillPath?: string;
  /** The enclosing group's fill, for `a:grpFill`. */
  readonly groupFill?: ResolvedFill;
  readonly groupSourcePart?: string;
}

/** The paint context of one fill, inside this lane. */
export interface FillContext extends PaintContext {
  /**
   * The matrix from the fill box to the slide (EMU), for pattern tiles on the slide's grid. Paint
   * sets it for each shape and picture box.
   */
  readonly fillMatrix?: Matrix;
}

/** An SVG paint (`fill` or `stroke` value) plus its opacity. */
export interface Paint {
  readonly paint: string;
  readonly opacity: number;
}

const NONE: Paint = { paint: 'none', opacity: 1 };

function colorPaint(c: Rgba): Paint {
  return { paint: toCss({ ...c, a: 1 }), opacity: c.a };
}

/** Lighten/darken a paint for preset path fill modes (`a:path/@fill`). */
export function adjustForPathMode(c: Rgba, mode: string): Rgba {
  const f =
    mode === 'darken'
      ? 0.6
      : mode === 'darkenLess'
        ? 0.8
        : mode === 'lighten'
          ? 0.6
          : mode === 'lightenLess'
            ? 0.8
            : 1;
  if (mode.startsWith('darken')) return { r: c.r * f, g: c.g * f, b: c.b * f, a: c.a };
  if (mode.startsWith('lighten'))
    return { r: 255 - (255 - c.r) * f, g: 255 - (255 - c.g) * f, b: 255 - (255 - c.b) * f, a: c.a };
  return c;
}

/** Where a picture fill's image lies in a `w`×`h` box: `fillRect` places it, `srcRect` crops it. */
export function pictureRect(
  fill: Extract<ResolvedFill, { type: 'picture' }>,
  w: number,
  h: number
): { x: number; y: number; width: number; height: number } {
  const fr = fill.fillRect;
  const bx = fr.l * w;
  const by = fr.t * h;
  const bw = w * (1 - fr.l - fr.r);
  const bh = h * (1 - fr.t - fr.b);
  // A crop that leaves nothing visible on an axis (two concurrent crops merged edge by edge)
  // draws uncropped on that axis, as the crop commands read it.
  const raw = fill.srcRect;
  const okW = 1 - raw.l - raw.r >= 0.001;
  const okH = 1 - raw.t - raw.b >= 0.001;
  const sr = {
    l: okW ? raw.l : 0,
    r: okW ? raw.r : 0,
    t: okH ? raw.t : 0,
    b: okH ? raw.b : 0,
  };
  const visW = 1 - sr.l - sr.r;
  const visH = 1 - sr.t - sr.b;
  const iw = bw / visW;
  const ih = bh / visH;
  return { x: bx - sr.l * iw, y: by - sr.t * ih, width: iw, height: ih };
}

/**
 * The stored pixels, never turned by a JPEG's EXIF orientation,
 * while a browser applies it to an SVG `<image>` by default. @internal
 */
export const STORED_PIXELS_STYLE = 'image-orientation: none';

/** Build the SVG image element for a picture fill over a `w`×`h` box, cropped by srcRect. */
export function pictureImage(
  doc: Document,
  defs: Defs,
  fill: Extract<ResolvedFill, { type: 'picture' }>,
  url: string,
  w: number,
  h: number
): SVGElement {
  const r = pictureRect(fill, w, h);
  const filter = blipFilter(doc, defs, fill.effects);
  const alpha = fill.effects.reduce((a, e) => (e.kind === 'alphaModFix' ? a * e.amount : a), 1);
  return svg(doc, 'image', {
    href: url,
    x: num(r.x),
    y: num(r.y),
    width: num(r.width),
    height: num(r.height),
    preserveAspectRatio: 'none',
    // The browser decodes the picture off the frame: a slide of large pictures never holds
    // a frame for seconds, and each picture shows when it is decoded.
    decoding: 'async',
    style: STORED_PIXELS_STYLE,
    opacity: alpha < 1 ? num(alpha) : undefined,
    filter,
  });
}

/** Grayscale, duotone, luminance, and bi-level blip effects as one SVG filter. */
function blipFilter(doc: Document, defs: Defs, effects: readonly BlipEffect[]): string | undefined {
  const primitives: SVGElement[] = [];
  for (const e of effects) {
    if (e.kind === 'grayscale')
      primitives.push(svg(doc, 'feColorMatrix', { type: 'saturate', values: '0' }));
    else if (e.kind === 'duotone') {
      primitives.push(svg(doc, 'feColorMatrix', { type: 'saturate', values: '0' }));
      const ch = (a: number, b: number): string => `${num(a / 255)} ${num(b / 255)}`;
      primitives.push(
        svg(doc, 'feComponentTransfer', {}, [
          svg(doc, 'feFuncR', { type: 'table', tableValues: ch(e.dark.r, e.light.r) }),
          svg(doc, 'feFuncG', { type: 'table', tableValues: ch(e.dark.g, e.light.g) }),
          svg(doc, 'feFuncB', { type: 'table', tableValues: ch(e.dark.b, e.light.b) }),
        ])
      );
    } else if (e.kind === 'lum') {
      // Brightness shifts, contrast scales around mid-gray (ECMA-376 §20.1.8.42).
      const { slope, intercept } = lumTransfer(e.bright, e.contrast);
      const fn = (tag: string): SVGElement =>
        svg(doc, tag, { type: 'linear', slope: num(slope), intercept: num(intercept) });
      primitives.push(
        svg(doc, 'feComponentTransfer', {}, [fn('feFuncR'), fn('feFuncG'), fn('feFuncB')])
      );
    } else if (e.kind === 'biLevel') {
      primitives.push(svg(doc, 'feColorMatrix', { type: 'saturate', values: '0' }));
      // Black below the threshold, white from it (ECMA-376 Part 1, 20.1.8.11): a ramp so steep
      // that one 8-bit step is the whole range.
      const slope = 10_000;
      const fn = (tag: string): SVGElement =>
        svg(doc, tag, { type: 'linear', slope, intercept: num(1 - slope * e.threshold) });
      primitives.push(
        svg(doc, 'feComponentTransfer', {}, [fn('feFuncR'), fn('feFuncG'), fn('feFuncB')])
      );
    }
  }
  if (primitives.length === 0) return undefined;
  const id = defs.id('blip');
  return defs.add(svg(doc, 'filter', { id, 'color-interpolation-filters': 'sRGB' }, primitives));
}

/** The image URL of a picture fill: its SVG image, else its bitmap. Undefined when neither resolves. */
export function pictureUrl(
  ctx: Pick<PaintContext, 'images' | 'sourcePart'>,
  fill: Extract<ResolvedFill, { type: 'picture' }>
): string | undefined {
  return (
    ctx.images(ctx.sourcePart, fill.svgRId ?? fill.rId) ?? ctx.images(ctx.sourcePart, fill.rId)
  );
}

/** Resolve a fill to an SVG paint over a `w`×`h` local box. */
export function fillPaint(
  fill: ResolvedFill,
  w: number,
  h: number,
  ctx: FillContext,
  mode = 'norm'
): Paint {
  const { defs } = ctx;
  const doc = defs.doc;
  switch (fill.type) {
    case 'none':
      return NONE;
    case 'solid':
      return colorPaint(adjustForPathMode(fill.color, mode));
    case 'group':
      return ctx.groupFill && ctx.groupFill.type !== 'group'
        ? fillPaint(
            ctx.groupFill,
            w,
            h,
            { ...ctx, sourcePart: ctx.groupSourcePart ?? ctx.sourcePart },
            mode
          )
        : NONE;
    case 'linear':
    case 'path': {
      const stops = fill.stops.map((s) => ({ ...s, color: adjustForPathMode(s.color, mode) }));
      if (fill.type === 'linear')
        return { paint: linearGradient(defs, fill, stops, w, h), opacity: 1 };
      // Stop alpha varies per stop; the paint server carries it, so the paint itself is opaque.
      const paint =
        fill.path === 'circle'
          ? circleGradient(defs, fill, stops, w, h)
          : fill.path === 'shape' && ctx.fillPath
            ? shapeGradient(defs, fill, stops, w, h, ctx.fillPath)
            : rectGradient(defs, fill, stops, w, h);
      return { paint, opacity: 1 };
    }
    case 'picture': {
      const url = pictureUrl(ctx, fill);
      if (!url) return NONE;
      const id = defs.id('pic');
      if (fill.mode === 'tile' && fill.tile) return tilePaint(fill, url, w, h, ctx);
      const pattern = svg(
        doc,
        'pattern',
        { id, patternUnits: 'userSpaceOnUse', x: 0, y: 0, width: num(w), height: num(h) },
        [pictureImage(doc, defs, fill, url, w, h)]
      );
      return { paint: defs.add(pattern), opacity: 1 };
    }
    case 'pattern':
      // The preset path fill modes lighten and darken both colors.
      return patternPaint(
        { ...fill, fg: adjustForPathMode(fill.fg, mode), bg: adjustForPathMode(fill.bg, mode) },
        ctx
      );
  }
}

/**
 * A preset pattern as its 8×8 bitmap in the two colors (design D2 of `fill-and-line-fidelity`).
 * The tiles sit on a grid from the slide's top-left corner, square to the slide: the reference PDF
 * places every tile at a multiple of 8 pt, and a rotated or flipped shape shows the same unturned
 * tiles (`fixtures/follow-ups.pdf`, slide 3). The pattern space is the slide's, so the tiles keep
 * their size in a scaled group too.
 */
function patternPaint(fill: Extract<ResolvedFill, { type: 'pattern' }>, ctx: FillContext): Paint {
  const { defs } = ctx;
  const doc = defs.doc;
  const bits = patternBitmap(fill.preset) ?? patternBitmap('pct50')!;
  const size = 8 * PATTERN_PIXEL;
  const fg = toCss({ ...fill.fg, a: 1 });
  const kids: SVGElement[] = [
    svg(doc, 'rect', {
      width: num(size),
      height: num(size),
      fill: toCss({ ...fill.bg, a: 1 }),
      'fill-opacity': fill.bg.a < 1 ? num(fill.bg.a) : undefined,
    }),
  ];
  // One rect per run of equal coverage in a row keeps the pattern small. A half-coverage pixel
  // draws the foreground at half opacity: the midpoint of the two colors over an opaque background.
  bits.forEach((row, y) => {
    let x = 0;
    while (x < 8) {
      const coverage = row[x];
      if (!coverage) {
        x += 1;
        continue;
      }
      let end = x;
      while (end < 8 && row[end] === coverage) end += 1;
      const opacity = coverage * fill.fg.a;
      kids.push(
        svg(doc, 'rect', {
          x: num(x * PATTERN_PIXEL),
          y: num(y * PATTERN_PIXEL),
          width: num((end - x) * PATTERN_PIXEL),
          height: num(PATTERN_PIXEL),
          fill: fg,
          'fill-opacity': opacity < 1 ? num(opacity) : undefined,
          'shape-rendering': 'crispEdges',
        })
      );
      x = end;
    }
  });
  const id = defs.id('patt');
  const toSlide = ctx.fillMatrix && slideSpace(ctx.fillMatrix);
  return {
    paint: defs.add(
      svg(
        doc,
        'pattern',
        {
          id,
          patternUnits: 'userSpaceOnUse',
          patternTransform: toSlide,
          width: num(size),
          height: num(size),
        },
        kids
      )
    ),
    opacity: 1,
  };
}

/**
 * The `patternTransform` that maps slide space into a box's local space, or undefined when the box
 * maps to the slide by the identity or cannot be inverted (a box of zero width or height). It
 * inverts the matrix as `matrixAttr` writes it, so the two cancel exactly.
 */
function slideSpace(m: Matrix): string | undefined {
  const r = roundedMatrix(m);
  if (r[0] === 1 && r[1] === 0 && r[2] === 0 && r[3] === 1 && r[4] === 0 && r[5] === 0)
    return undefined;
  return invertedAttr(r);
}

/** A matrix as `matrixAttr` writes it. */
function roundedMatrix(m: Matrix): Matrix {
  return m.map((v) => Number(num(v))) as unknown as Matrix;
}

/** The inverse of `r` as a `transform` value, or undefined when `r` cannot be inverted. */
function invertedAttr(r: Matrix): string | undefined {
  const det = r[0] * r[3] - r[1] * r[2];
  if (!Number.isFinite(det) || Math.abs(det) < 1e-9) return undefined;
  const fine = (v: number): string => {
    const x = Math.round(v * 1e9) / 1e9;
    return Object.is(x, -0) ? '0' : String(x);
  };
  return `matrix(${invert(r).map(fine).join(' ')})`;
}

/**
 * How to draw a pattern fill of a box that the slide turns, flips, or scales: in slide space,
 * under a `transform` that undoes the box's matrix. A browser that prints a pattern under a turned
 * transform resamples its tiles; in slide space they print pixel for pixel. `matrix` is the box's
 * matrix as `matrixAttr` writes it, after `outer` (the turned face of a 3-D camera, which an
 * enclosing group applies): map the fill path through it. Undefined for a box that maps to the
 * slide by a translation, or that cannot be inverted.
 */
export function slideSpaceFill(
  m: Matrix,
  outer?: Matrix
): { matrix: Matrix; transform: string } | undefined {
  const r = outer ? multiply(roundedMatrix(outer), roundedMatrix(m)) : roundedMatrix(m);
  if (!outer && r[0] === 1 && r[1] === 0 && r[2] === 0 && r[3] === 1) return undefined;
  const transform = invertedAttr(r);
  return transform ? { matrix: r, transform } : undefined;
}

/** The smallest tile side, EMU (one pixel at 96 DPI): a file cannot ask for subpixel tiles. */
const MIN_TILE = 9525;
/**
 * The largest tile side, EMU: the largest slide side (56 in, `ST_SlideSizeCoordinate`). A larger
 * tile shows no more of the image on any slide, and a file cannot ask for a larger pattern.
 */
const MAX_TILE = 51_206_400;

/**
 * A tiled picture (`CT_TileInfoProperties`; design D4 of `fill-and-line-fidelity`): tiles at the
 * image's natural size × `sx`/`sy`, placed at the `algn` point plus `tx`/`ty`, mirrored by `flip`,
 * recolored by the blip effects. Without a size port, a tile is the box.
 */
function tilePaint(
  fill: Extract<ResolvedFill, { type: 'picture' }>,
  url: string,
  w: number,
  h: number,
  ctx: PaintContext
): Paint {
  const { defs } = ctx;
  const doc = defs.doc;
  const tile = fill.tile!;
  const natural = ctx.imageSize?.(ctx.sourcePart, fill.rId) ?? { cx: w, cy: h };
  const sw = Math.max(Math.abs(natural.cx * tile.sx) || w, MIN_TILE);
  const sh = Math.max(Math.abs(natural.cy * tile.sy) || h, MIN_TILE);
  // One factor for both sides, so a capped tile keeps its aspect ratio.
  const fit = Math.min(1, MAX_TILE / sw, MAX_TILE / sh);
  const tw = Math.max(sw * fit, MIN_TILE);
  const th = Math.max(sh * fit, MIN_TILE);
  // `algn` comes from the file: an own-key check, so `toString` cannot reach the prototype.
  const align = Object.hasOwn(ALIGN_POINTS, tile.algn) ? (tile.algn as EffectAlign) : 'tl';
  const [ax, ay] = ALIGN_POINTS[align];
  const fx = tile.flip === 'x' || tile.flip === 'xy';
  const fy = tile.flip === 'y' || tile.flip === 'xy';
  const image = (mx: boolean, my: boolean): SVGElement =>
    svg(doc, 'image', {
      href: url,
      width: num(tw),
      height: num(th),
      preserveAspectRatio: 'none',
      decoding: 'async',
      style: STORED_PIXELS_STYLE,
      transform:
        mx || my
          ? `translate(${num(mx ? 2 * tw : 0)} ${num(my ? 2 * th : 0)}) scale(${mx ? -1 : 1} ${my ? -1 : 1})`
          : undefined,
    });
  const images = [image(false, false)];
  if (fx) images.push(image(true, false));
  if (fy) images.push(image(false, true));
  if (fx && fy) images.push(image(true, true));
  const id = defs.id('tile');
  const alpha = fill.effects.reduce((a, e) => (e.kind === 'alphaModFix' ? a * e.amount : a), 1);
  const filter = blipFilter(doc, defs, fill.effects);
  const pattern = svg(
    doc,
    'pattern',
    {
      id,
      patternUnits: 'userSpaceOnUse',
      x: num(ax * w - ax * tw + tile.tx),
      y: num(ay * h - ay * th + tile.ty),
      width: num(fx ? 2 * tw : tw),
      height: num(fy ? 2 * th : th),
    },
    filter ? [svg(doc, 'g', { filter }, images)] : images
  );
  return { paint: defs.add(pattern), opacity: alpha };
}

/** Stroke attributes for a resolved outline. `scale` corrects for group scaling. */
export function strokeAttrs(
  line: ResolvedLine,
  w: number,
  h: number,
  ctx: FillContext,
  scale = 1
): Record<string, string | undefined> {
  const paint = fillPaint(line.fill, w, h, ctx);
  if (paint.paint === 'none') return { stroke: 'none' };
  // The width, dash pattern, and caps are the PDF export's too (`layout/line-strokes.ts`).
  const width = strokeWidth(line, scale);
  const dashes = dashArray(line, width);
  return {
    stroke: paint.paint,
    'stroke-opacity': paint.opacity < 1 ? num(paint.opacity) : undefined,
    'stroke-width': num(width),
    'stroke-linecap': strokeCap(line),
    'stroke-linejoin': line.join === 'miter' ? 'miter' : line.join === 'bevel' ? 'bevel' : 'round',
    'stroke-miterlimit': line.join === 'miter' ? num(Math.max(1, line.miterLimit)) : undefined,
    'stroke-dasharray': dashes?.map(num).join(' '),
  };
}
