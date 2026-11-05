// Fill and outline paint: solid, linear and path gradients (`paint-gradient.ts`), pictures
// (stretch with srcRect and fillRect, tile), patterns, dashes, and arrowheads. Paint servers are
// defined in the shape's local coordinate space (userSpaceOnUse over the shape box), so rotation
// and flips carry them.

import {
  toCss,
  type EffectAlign,
  type Rgba,
  type ResolvedFill,
  type ResolvedLine,
  type BlipEffect,
} from '../cascade/index.ts';
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
   * The slide position (EMU) of the fill box's top-left corner, for pattern tiles on the slide's
   * grid. Paint sets it for each shape box that maps to the slide by a translation only.
   */
  readonly fillOrigin?: { readonly x: number; readonly y: number };
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
      const url =
        ctx.images(ctx.sourcePart, fill.svgRId ?? fill.rId) ?? ctx.images(ctx.sourcePart, fill.rId);
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
      return patternPaint(fill, ctx);
  }
}

/**
 * A preset pattern as its 8×8 bitmap in the two colors (design D2 of `fill-and-line-fidelity`).
 * The tiles sit on a grid from the slide's top-left corner, not from the shape: the reference PDF
 * places every tile at a multiple of 8 pt (`fixtures/fills-lines.pdf`, slide 1).
 */
function patternPaint(fill: Extract<ResolvedFill, { type: 'pattern' }>, ctx: FillContext): Paint {
  const { defs } = ctx;
  const doc = defs.doc;
  const bits = patternBitmap(fill.preset) ?? patternBitmap('pct50')!;
  const size = 8 * PATTERN_PIXEL;
  const fg = toCss({ ...fill.fg, a: 1 });
  const phase = (v: number): number => -(((v % size) + size) % size);
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
  const origin = ctx.fillOrigin ?? { x: 0, y: 0 };
  return {
    paint: defs.add(
      svg(
        doc,
        'pattern',
        {
          id,
          patternUnits: 'userSpaceOnUse',
          x: num(phase(origin.x)),
          y: num(phase(origin.y)),
          width: num(size),
          height: num(size),
        },
        kids
      )
    ),
    opacity: 1,
  };
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

/** Preset dash patterns, in multiples of the line width (ECMA-376 ST_PresetLineDashVal). */
const DASHES: Record<string, readonly number[]> = {
  dot: [1, 3],
  dash: [4, 3],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1],
  sysDot: [1, 1],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

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
  // A zero-width line is a hairline: one device pixel at 96 dpi.
  const width = (line.width > 0 ? line.width : 9525) / (scale || 1);
  // `prstDash/@val` comes from the file: an own-key check keeps prototype names out.
  const pattern =
    line.dash === 'custom'
      ? line.dashes
      : Object.hasOwn(DASHES, line.dash)
        ? DASHES[line.dash]
        : undefined;
  const roundCap = line.cap === 'rnd';
  const dasharray = pattern
    ? pattern
        .map((v, i) => {
          // Round caps extend each dash by the width; this keeps the visible rhythm.
          const len = v * width + (roundCap ? (i % 2 === 0 ? -width : width) : 0);
          return num(Math.max(i % 2 === 0 ? 0.001 : 0, len));
        })
        .join(' ')
    : undefined;
  return {
    stroke: paint.paint,
    'stroke-opacity': paint.opacity < 1 ? num(paint.opacity) : undefined,
    'stroke-width': num(width),
    'stroke-linecap': line.cap === 'rnd' ? 'round' : line.cap === 'sq' ? 'square' : 'butt',
    'stroke-linejoin': line.join === 'miter' ? 'miter' : line.join === 'bevel' ? 'bevel' : 'round',
    'stroke-miterlimit': line.join === 'miter' ? num(Math.max(1, line.miterLimit)) : undefined,
    'stroke-dasharray': dasharray,
  };
}

const END_SIZE = { sm: 2, med: 3, lg: 5 } as const;

/**
 * How far the line stops short of its end point, in line widths, so it ends inside the head
 * rather than under its narrowing tip: a triangle at its base, a stealth arrow at its notch.
 * Open arrows, circles, and diamonds keep the line to the end point (the head covers it).
 */
export function lineEndInset(
  end: { type: string; w: 'sm' | 'med' | 'lg'; len: 'sm' | 'med' | 'lg' } | undefined
): number {
  if (!end) return 0;
  const len = END_SIZE[end.len];
  if (end.type === 'triangle') return len * 0.9;
  if (end.type === 'stealth') return len * 0.65;
  // An open arrow's mitered tip reaches 0.5 / sin(half angle) past its vertex; the line stops
  // half a width inside that point, so the tip lands on the end point and covers the line.
  if (end.type === 'arrow') return 0.5 / Math.sin(Math.atan2(END_SIZE[end.w] / 2, len)) + 0.5;
  return 0;
}

/** An SVG marker for a line end, or undefined for `none`. */
export function lineEndMarker(
  defs: Defs,
  end: { type: string; w: 'sm' | 'med' | 'lg'; len: 'sm' | 'med' | 'lg' } | undefined,
  color: string,
  opacity: number
): string | undefined {
  if (!end) return undefined;
  const doc = defs.doc;
  const mw = END_SIZE[end.w];
  const ml = END_SIZE[end.len];
  let shape: SVGElement;
  switch (end.type) {
    case 'triangle':
      shape = svg(doc, 'path', { d: `M0 0 L${ml} ${mw / 2} L0 ${mw} Z`, fill: color });
      break;
    case 'stealth':
      shape = svg(doc, 'path', {
        d: `M0 0 L${ml} ${mw / 2} L0 ${mw} L${ml * 0.3} ${mw / 2} Z`,
        fill: color,
      });
      break;
    case 'diamond':
      shape = svg(doc, 'path', {
        d: `M0 ${mw / 2} L${ml / 2} 0 L${ml} ${mw / 2} L${ml / 2} ${mw} Z`,
        fill: color,
      });
      break;
    case 'oval':
      shape = svg(doc, 'ellipse', { cx: ml / 2, cy: mw / 2, rx: ml / 2, ry: mw / 2, fill: color });
      break;
    case 'arrow':
      shape = svg(doc, 'path', {
        d: `M0 0 L${ml} ${mw / 2} L0 ${mw}`,
        fill: 'none',
        stroke: color,
        // The head's strokes are as wide as the line (marker units are line widths).
        'stroke-width': 1,
        'stroke-linejoin': 'miter',
      });
      break;
    default:
      return undefined;
  }
  const centered = end.type === 'oval' || end.type === 'diamond';
  const id = defs.id('end');
  return defs.add(
    svg(
      doc,
      'marker',
      {
        id,
        markerUnits: 'strokeWidth',
        markerWidth: ml,
        markerHeight: mw,
        // The tip sits on the end point; circles and diamonds center on it.
        // The path is shortened by `lineEndInset`, so the anchor moves back by the same amount.
        refX: centered
          ? ml / 2
          : end.type === 'arrow'
            ? ml + 0.5 / Math.sin(Math.atan2(mw / 2, ml)) - lineEndInset(end)
            : ml - lineEndInset(end),
        refY: mw / 2,
        orient: 'auto-start-reverse',
        overflow: 'visible',
        opacity: opacity < 1 ? num(opacity) : undefined,
      },
      [shape]
    )
  );
}
