// Paint a laid-out slide as SVG. The painter reads layout records only; it never resolves a
// property or re-derives geometry. The editor and the fidelity harness both call this, so there
// is no second renderer for tests.

import {
  toCss,
  type BlendMode,
  type ResolvedEffects,
  type ResolvedFill,
} from '../cascade/index.ts';
import {
  apply as applyMatrix,
  hasCharacters,
  multiply as multiplyMatrix,
  type GeometryPath,
  type LayoutItem,
  type Matrix,
  type SlideLayout,
} from '../layout/index.ts';
import {
  bodyFilter,
  glowFilter,
  reflectionCopy,
  shadowCopy,
  type EffectSource,
} from './paint-effects.ts';
import { boundsOf, cornersUnder, grow, PX } from './effect-transforms.ts';
import {
  fillPaint,
  lineEndMarker,
  pictureImage,
  pictureRect,
  strokeAttrs,
  lineEndInset,
  type ImageSizePort,
  type FillContext,
  type ImageUrlPort,
  type Paint,
  type PaintContext,
} from './paint-fill.ts';
import { compoundStroke } from './compound-lines.ts';
import { paintTable } from './paint-table.ts';
import { paintText } from './paint-text.ts';
import { MAX_SHAPE_BANDS } from './paint-gradient.ts';
import { Defs, matrixAttr, num, svg, SVG_NS } from './svg.ts';

/** Paint options. */
export interface PaintOptions {
  /** Resolves a relationship id from a source part to an image URL. */
  readonly images: ImageUrlPort;
  /** Resolves a relationship id to the image's natural size, for tiled pictures. */
  readonly imageSize?: ImageSizePort;
  /** Unique prefix for ids in this SVG (several slides can share a page). */
  readonly idPrefix: string;
  /** Label for an element the editor does not render (chart, diagram, OLE). */
  readonly frameLabel?: (kind: string) => string;
  /** Include `data-node-id` hooks for hit testing. Default true. */
  readonly hitTargets?: boolean;
}

/**
 * Whether a path is open: it has no `close`. Line ends (`a:headEnd`,
 * `a:tailEnd`) show only on open paths, so a rectangle or a callout box never gets an arrowhead.
 */
export function isOpenPath(path: GeometryPath): boolean {
  return path.commands.length >= 2 && !path.commands.some((c) => c.op === 'Z');
}

/**
 * An open path shortened by `start` at its first point and `end` at its last, along the
 * tangent there, so an arrowhead covers the line's end. Closed paths are unchanged.
 */
export function shortenPath(path: GeometryPath, start: number, end: number): GeometryPath {
  const cmds = path.commands;
  if (!isOpenPath(path)) return path;
  const out = cmds.slice();
  const move = (
    x: number,
    y: number,
    tx: number,
    ty: number,
    by: number
  ): { x: number; y: number } => {
    const len = Math.hypot(x - tx, y - ty);
    if (!len || by <= 0) return { x, y };
    const k = Math.min(by, len * 0.9) / len;
    return { x: x + (tx - x) * k, y: y + (ty - y) * k };
  };
  // The tangent at an end runs to the nearest control point, or the neighboring point.
  const toward = (c: (typeof cmds)[number], fallback: { x: number; y: number }) =>
    c.op === 'C' ? { x: c.x2, y: c.y2 } : c.op === 'Q' ? { x: c.x1, y: c.y1 } : fallback;
  const first = out[0];
  const second = out[1];
  if (start > 0 && first.op === 'M' && second.op !== 'Z') {
    const t =
      second.op === 'C'
        ? { x: second.x1, y: second.y1 }
        : second.op === 'Q'
          ? { x: second.x1, y: second.y1 }
          : second;
    out[0] = { ...first, ...move(first.x, first.y, t.x, t.y, start) };
  }
  const last = out[out.length - 1];
  const prev = out[out.length - 2];
  if (end > 0 && last.op !== 'Z' && last.op !== 'M' && prev.op !== 'Z') {
    const t = toward(last, prev);
    out[out.length - 1] = { ...last, ...move(last.x, last.y, t.x, t.y, end) };
  }
  return { ...path, commands: out };
}

/** Path data for an evaluated path in local EMU. */
export function pathData(path: GeometryPath): string {
  let d = '';
  for (const c of path.commands) {
    switch (c.op) {
      case 'M':
      case 'L':
        d += `${c.op}${num(c.x)} ${num(c.y)}`;
        break;
      case 'C':
        d += `C${num(c.x1)} ${num(c.y1)} ${num(c.x2)} ${num(c.y2)} ${num(c.x)} ${num(c.y)}`;
        break;
      case 'Q':
        d += `Q${num(c.x1)} ${num(c.y1)} ${num(c.x)} ${num(c.y)}`;
        break;
      case 'Z':
        d += 'Z';
        break;
    }
  }
  return d;
}

/** The uniform scale a matrix applies (for line widths inside scaled groups). */
function matrixScale(m: Matrix): number {
  return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1;
}

/**
 * The most paths of one shape that draw a `path="shape"` gradient (fill or overlay) as outline
 * bands; a shape with more paths uses the rectangle construction, so a file cannot multiply the
 * bands.
 */
const MAX_BANDED_PATHS = 16;
/** The most path commands of one shape that draw outline bands: bands repeat every command. */
const MAX_BANDED_COMMANDS = 2000;
/**
 * The band work (path commands × `MAX_SHAPE_BANDS`) one slide may spend. Later shapes use the
 * rectangle construction, so many shapes cannot multiply the bands either.
 */
const MAX_SLIDE_BAND_WORK = 512_000;
/** The band work left on each slide, by its `Defs`. */
const bandWork = new WeakMap<Defs, { left: number }>();
/**
 * The most paths of one shape that draw a fill overlay. Each overlay is an isolated group (an
 * offscreen layer); a shape with more paths draws no overlay, and the report says so.
 */
export const MAX_OVERLAY_PATHS = 64;

/** CSS blend modes for `ST_BlendMode`. */
const BLEND_MODES: Record<BlendMode, string> = {
  over: 'normal',
  mult: 'multiply',
  screen: 'screen',
  darken: 'darken',
  lighten: 'lighten',
};

/** Whether a fill's paint server depends on the path it fills (a shape gradient). */
function pathDependent(fill: ResolvedFill, groupFill: ResolvedFill | undefined): boolean {
  const painted = fill.type === 'group' ? groupFill : fill;
  return painted?.type === 'path' && painted.path === 'shape';
}

/**
 * Paint servers for one shape, made once per fill mode (and per path only for a shape gradient),
 * so a shape with many paths does not multiply gradient definitions.
 */
function shapePaints(
  item: Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>,
  ctx: FillContext
): (fill: ResolvedFill, d: string, mode?: string) => Paint {
  const { cx: w, cy: h } = item.element.xfrm;
  const paths = item.geometry.paths;
  const commands = paths.reduce((n, p) => n + p.commands.length, 0);
  let banded: boolean | undefined;
  // Decided at the first shape gradient, so only shapes that band spend the slide's budget.
  const bands = (): boolean => {
    if (banded === undefined) {
      let budget = bandWork.get(ctx.defs);
      if (!budget) bandWork.set(ctx.defs, (budget = { left: MAX_SLIDE_BAND_WORK }));
      const cost = commands * MAX_SHAPE_BANDS;
      banded =
        paths.length <= MAX_BANDED_PATHS && commands <= MAX_BANDED_COMMANDS && cost <= budget.left;
      if (banded) budget.left -= cost;
    }
    return banded;
  };
  const made = new Map<ResolvedFill, Map<string, Paint>>();
  return (fill, d, mode = 'norm') => {
    const band = pathDependent(fill, ctx.groupFill) && bands();
    const key = band ? `${mode}|${d}` : mode;
    let byKey = made.get(fill);
    if (!byKey) made.set(fill, (byKey = new Map()));
    let paint = byKey.get(key);
    if (!paint) {
      paint = fillPaint(fill, w, h, band ? { ...ctx, fillPath: d } : ctx, mode);
      byKey.set(key, paint);
    }
    return paint;
  };
}

/** The `a:fillOverlay` path for a filled outline, or null (design D5 of `fill-fidelity`). */
function overlayPath(
  doc: Document,
  d: string,
  effects: ResolvedEffects | undefined,
  paintFor: (fill: ResolvedFill, d: string) => Paint
): SVGElement | null {
  const overlay = effects?.fillOverlay;
  if (!overlay || !d) return null;
  const paint = paintFor(overlay.fill, d);
  if (paint.paint === 'none') return null;
  return svg(doc, 'path', {
    d,
    fill: paint.paint,
    'fill-opacity': paint.opacity < 1 ? num(paint.opacity) : undefined,
    'fill-rule': 'evenodd',
    style: `mix-blend-mode: ${BLEND_MODES[overlay.blend]}`,
  });
}

/** A group whose blend modes see only its own content, not the slide behind it. */
function isolated(doc: Document, children: readonly SVGElement[]): SVGElement {
  return svg(doc, 'g', { style: 'isolation: isolate' }, children);
}

/**
 * A mask that keeps a picture inside its geometry (design D7 of `fill-fidelity`). Chromium can
 * misplace a `clipPath` edge in EMU coordinates under a translation; a mask clips correctly.
 */
function geometryMask(
  defs: Defs,
  paths: readonly GeometryPath[],
  d: string,
  w: number,
  h: number
): string {
  let x0 = 0;
  let y0 = 0;
  let x1 = w;
  let y1 = h;
  for (const p of paths)
    for (const c of p.commands) {
      if (c.op === 'Z') continue;
      const pts =
        c.op === 'C'
          ? [c.x1, c.y1, c.x2, c.y2, c.x, c.y]
          : c.op === 'Q'
            ? [c.x1, c.y1, c.x, c.y]
            : [c.x, c.y];
      for (let i = 0; i < pts.length; i += 2) {
        x0 = Math.min(x0, pts[i]);
        x1 = Math.max(x1, pts[i]);
        y0 = Math.min(y0, pts[i + 1]);
        y1 = Math.max(y1, pts[i + 1]);
      }
    }
  const id = defs.id('clip');
  defs.add(
    svg(
      defs.doc,
      'mask',
      {
        id,
        maskUnits: 'userSpaceOnUse',
        x: num(x0 - 1),
        y: num(y0 - 1),
        width: num(x1 - x0 + 2),
        height: num(y1 - y0 + 2),
      },
      [svg(defs.doc, 'path', { d, fill: '#ffffff' })]
    )
  );
  return `url(#${id})`;
}

/**
 * A picture's outline. It lies outside the picture, not centered on its edge,
 * as reference PDFs show. So the stroke is twice as wide, and a mask hides the half inside the
 * geometry. A round or square cap of the wider stroke reaches twice as far along the path, so
 * dashes get butt caps and grow by one line width instead: the dash period of a centered line
 * stays, and round dots draw as squares. Dashes outside a picture are not checked against a
 * reference PDF (approximation `picture-border-dots`).
 */
function outsideOutline(
  doc: Document,
  ctx: PaintContext,
  d: string,
  stroke: Record<string, string | number | undefined>,
  w: number,
  h: number
): SVGElement {
  const width = Number(stroke['stroke-width']) || 0;
  const pad = width * 2 + Math.max(w, h);
  const id = ctx.defs.id('outside');
  ctx.defs.add(
    svg(
      doc,
      'mask',
      {
        id,
        maskUnits: 'userSpaceOnUse',
        x: -pad,
        y: -pad,
        width: w + 2 * pad,
        height: h + 2 * pad,
      },
      [
        svg(doc, 'rect', {
          x: -pad,
          y: -pad,
          width: w + 2 * pad,
          height: h + 2 * pad,
          fill: '#fff',
        }),
        svg(doc, 'path', { d, fill: '#000' }),
      ]
    )
  );
  const dashes = stroke['stroke-dasharray'];
  const capped = dashes !== undefined && stroke['stroke-linecap'] !== 'butt';
  return svg(doc, 'path', {
    d,
    fill: 'none',
    ...stroke,
    'stroke-width': num(width * 2),
    ...(capped
      ? {
          'stroke-linecap': 'butt',
          'stroke-dasharray': String(dashes)
            .split(' ')
            .map((v, i) => num(Math.max(0, Number(v) + (i % 2 === 0 ? width : -width))))
            .join(' '),
        }
      : {}),
    mask: `url(#${id})`,
  });
}

/**
 * The context with the slide position of a shape's box, for pattern tiles on the slide's grid.
 * Only a box that maps to the slide by a translation gets one: the tile grid of a rotated,
 * flipped, or scaled box is not measured, so its tiles start at the box's corner.
 */
function withFillOrigin(
  item: Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>,
  ctx: PaintContext
): FillContext {
  const [a, b, c, d, e, f] = item.matrix;
  return a === 1 && b === 0 && c === 0 && d === 1 ? { ...ctx, fillOrigin: { x: e, y: f } } : ctx;
}

function paintGeometry(
  doc: Document,
  item: Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>,
  shapeCtx: PaintContext,
  groupScale: number
): SVGElement[] {
  const ctx = withFillOrigin(item, shapeCtx);
  const el = item.element;
  const { cx: w, cy: h } = el.xfrm;
  const out: SVGElement[] = [];
  const line = el.line;
  const stroke = line ? strokeAttrs(line, w, h, ctx, groupScale) : { stroke: 'none' };
  const hasEnds =
    item.kind === 'shape' &&
    line !== undefined &&
    stroke.stroke !== 'none' &&
    item.geometry.paths.some((p) => p.stroke && isOpenPath(p));
  const head = hasEnds ? lineEndMarker(ctx.defs, line.head, stroke.stroke ?? '#000', 1) : undefined;
  const tail = hasEnds ? lineEndMarker(ctx.defs, line.tail, stroke.stroke ?? '#000', 1) : undefined;
  const paintFor = shapePaints(item, ctx);
  const effects = item.geometry.paths.length <= MAX_OVERLAY_PATHS ? el.effects : undefined;
  if (item.kind === 'picture') {
    const fill = item.element.fill;
    const url =
      fill.type === 'picture'
        ? (ctx.images(ctx.sourcePart, fill.svgRId ?? fill.rId) ??
          ctx.images(ctx.sourcePart, fill.rId))
        : undefined;
    const d = item.geometry.paths.map(pathData).join('');
    if (url && fill.type === 'picture') {
      const image = pictureImage(doc, ctx.defs, fill, url, w, h);
      const overlay = overlayPath(doc, d, effects, paintFor);
      const content = overlay ? isolated(doc, [image, overlay]) : image;
      // A rectangle picture inside its box needs no clip; the mask would soften its edges.
      const r = pictureRect(fill, w, h);
      const inside =
        el.geometry.kind === 'preset' &&
        el.geometry.preset === 'rect' &&
        r.x >= -1 &&
        r.y >= -1 &&
        r.x + r.width <= w + 1 &&
        r.y + r.height <= h + 1;
      out.push(
        inside
          ? content
          : svg(doc, 'g', { mask: geometryMask(ctx.defs, item.geometry.paths, d, w, h) }, [content])
      );
    } else
      out.push(
        svg(doc, 'rect', {
          width: num(w),
          height: num(h),
          fill: '#e8e8e8',
          stroke: '#b0b0b0',
          'stroke-width': 9525,
        })
      );
    if (stroke.stroke !== 'none') out.push(outsideOutline(doc, ctx, d, stroke, w, h));
    return out;
  }
  const width = Number(stroke['stroke-width']) || 0;
  for (const path of item.geometry.paths) {
    // Line ends belong to open, stroked paths only.
    const ends = path.stroke && isOpenPath(path);
    const insets =
      ends && line && stroke.stroke !== 'none'
        ? [lineEndInset(line.head) * width, lineEndInset(line.tail) * width]
        : [0, 0];
    const d = pathData(insets[0] || insets[1] ? shortenPath(path, insets[0], insets[1]) : path);
    if (!d) continue;
    const fill =
      path.fill === 'none' ? { paint: 'none', opacity: 1 } : paintFor(el.fill, d, path.fill);
    const overlay = fill.paint === 'none' ? null : overlayPath(doc, d, effects, paintFor);
    if (overlay) {
      // The overlay blends with this fill only (design D5 of `fill-fidelity`); the outline draws on top.
      out.push(
        isolated(doc, [
          svg(doc, 'path', {
            d,
            fill: fill.paint,
            'fill-opacity': fill.opacity < 1 ? num(fill.opacity) : undefined,
            'fill-rule': 'evenodd',
          }),
          overlay,
        ])
      );
    }
    const compound =
      path.stroke && line && stroke.stroke !== 'none'
        ? compoundStroke(ctx.defs, d, line.compound, stroke, w, h)
        : null;
    const own = overlay ? { paint: 'none', opacity: 1 } : fill;
    if (compound) {
      // The bands draw in their own element; arrowheads stay on the center line.
      out.push(
        svg(doc, 'path', {
          d,
          fill: own.paint,
          'fill-opacity': own.opacity < 1 ? num(own.opacity) : undefined,
          'fill-rule': 'evenodd',
          stroke: 'none',
        }),
        compound
      );
      if (ends && (head || tail))
        out.push(
          svg(doc, 'path', {
            d,
            fill: 'none',
            stroke: 'none',
            'stroke-width': stroke['stroke-width'],
            'marker-start': head,
            'marker-end': tail,
          })
        );
      continue;
    }
    out.push(
      svg(doc, 'path', {
        d,
        fill: own.paint,
        'fill-opacity': own.opacity < 1 ? num(own.opacity) : undefined,
        'fill-rule': 'evenodd',
        ...(path.stroke ? stroke : { stroke: 'none' }),
        'marker-start': ends ? head : undefined,
        'marker-end': ends ? tail : undefined,
      })
    );
  }
  return out;
}

function paintItem(
  doc: Document,
  item: LayoutItem,
  ctx: PaintContext,
  opts: PaintOptions,
  parentScale: number
): SVGElement | null {
  const el = item.element;
  const hit = opts.hitTargets !== false;
  const common = {
    'data-node-id': hit ? el.nodeId : undefined,
    'data-part': hit ? el.partName : undefined,
    'data-layer': hit ? el.layer : undefined,
  };
  const local: PaintContext = { ...ctx, sourcePart: el.partName };
  switch (item.kind) {
    case 'group': {
      const g = svg(doc, 'g', { ...common, 'data-kind': 'group' });
      const scale = matrixScale(item.matrix);
      // A group with `a:grpFill` passes its own group's fill on (ECMA-376 Part 1, 20.1.8.35).
      const inherits = item.element.fill.type === 'group';
      const groupCtx: PaintContext = {
        ...local,
        groupFill: inherits ? ctx.groupFill : item.element.fill,
        groupSourcePart: inherits ? ctx.groupSourcePart : el.partName,
      };
      for (const child of item.children) {
        const painted = paintItem(doc, child, groupCtx, opts, scale);
        if (painted) g.appendChild(painted);
      }
      return g;
    }
    case 'frame': {
      if (item.table)
        return svg(
          doc,
          'g',
          { ...common, 'data-kind': 'table', transform: matrixAttr(item.matrix) },
          paintTable(doc, item.table, item.element.table?.background ?? { type: 'none' }, local)
        );
      // An OLE object shows its picture (ECMA-376 Part 1, 19.3.1.36); the frame holds the hit id.
      if (item.poster) {
        const picture = paintItem(
          doc,
          item.poster,
          local,
          { ...opts, hitTargets: false },
          parentScale
        );
        return svg(doc, 'g', { ...common, 'data-kind': 'frame' }, picture ? [picture] : []);
      }
      const { cx: w, cy: h } = el.xfrm;
      const label = opts.frameLabel?.(item.element.frameKind) ?? item.element.frameKind;
      // Points under a scale: browsers clamp font sizes as large as EMU values.
      const t = svg(doc, 'text', {
        x: num(w / 2 / 12_700),
        y: num(h / 2 / 12_700),
        'text-anchor': 'middle',
        'dominant-baseline': 'middle',
        'font-size': 12,
        fill: '#5f6368',
        'font-family': 'sans-serif',
        transform: 'scale(12700)',
      });
      t.textContent = label;
      return svg(
        doc,
        'g',
        { ...common, 'data-kind': 'frame', transform: matrixAttr(item.matrix) },
        [
          svg(doc, 'rect', {
            width: num(w),
            height: num(h),
            fill: '#f1f3f4',
            stroke: '#9aa0a6',
            'stroke-width': 9525,
            'stroke-dasharray': '76200 38100',
          }),
          t,
        ]
      );
    }
    case 'shape':
    case 'picture': {
      const outer = svg(doc, 'g', { ...common, 'data-kind': item.kind });
      const effects = item.element.effects;
      const body = svg(doc, 'g', { transform: matrixAttr(item.matrix) });
      for (const node of paintGeometry(doc, item, local, parentScale)) body.appendChild(node);
      const bodyFx =
        effects && (effects.softEdge || effects.innerShadow)
          ? bodyFilter(ctx.defs, effects, boundsOf(boxCorners(item)))
          : undefined;
      const content = svg(doc, 'g', {}, [
        bodyFx ? svg(doc, 'g', { filter: bodyFx }, [body]) : body,
      ]);
      // An empty body's caret line draws nothing: no group for it.
      if (item.kind === 'shape' && item.text && item.textMatrix && hasCharacters(item.text))
        content.appendChild(
          svg(doc, 'g', { transform: matrixAttr(item.textMatrix) }, [
            paintText(doc, item.text, local),
          ])
        );
      if (!effects || !(effects.outerShadow || effects.reflection || effects.glow)) {
        outer.appendChild(content);
        return outer;
      }
      // Shadow, reflection, and glow use the content, text included (design D1 of `effects-fidelity`).
      const contentId = ctx.defs.id('content');
      content.setAttribute('id', contentId);
      const source: EffectSource = {
        contentId,
        matrix: item.matrix,
        width: el.xfrm.cx,
        height: el.xfrm.cy,
        corners: contentCorners(item),
      };
      const reflection = reflectionCopy(ctx.defs, effects, source);
      if (reflection) outer.appendChild(reflection);
      const shadow = shadowCopy(ctx.defs, effects, source);
      if (shadow) outer.appendChild(shadow);
      const glow = glowFilter(ctx.defs, effects, boundsOf(source.corners));
      outer.appendChild(glow ? svg(doc, 'g', { filter: glow }, [content]) : content);
      return outer;
    }
  }
}

/**
 * How far the outline can paint past the shape's geometry, in slide space: a picture's outline
 * lies outside it, a shape's is centered, a miter join reaches up to half the miter limit times
 * the width, and line ends reach further. Effect filter regions grow by this, so they never cut
 * the outline off.
 */
function outlineReach(item: Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>): number {
  const line = item.element.line;
  if (!line || line.fill.type === 'none') return 0;
  const width = Math.max(line.width, PX);
  const join = line.join === 'miter' ? Math.max(2, line.miterLimit / 2) : 2;
  const ends = item.kind === 'shape' && item.geometry.paths.some((p) => p.stroke && isOpenPath(p));
  return width * Math.max(join, ends ? 3 : 0);
}

/**
 * The slide-space corners of a shape's box and geometry, with its outline. Geometry can reach
 * past the box (a callout tail, `a:custGeom`); a Bezier curve lies inside its control points.
 */
function boxCorners(
  item: Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>
): { x: number; y: number }[] {
  const { cx, cy } = item.element.xfrm;
  const m = item.matrix;
  const points = cornersUnder(m, cx, cy);
  for (const path of item.geometry.paths)
    for (const c of path.commands) {
      if (c.op === 'Z') continue;
      if (c.op === 'C') points.push(applyMatrix(m, c.x1, c.y1), applyMatrix(m, c.x2, c.y2));
      else if (c.op === 'Q') points.push(applyMatrix(m, c.x1, c.y1));
      points.push(applyMatrix(m, c.x, c.y));
    }
  const b = grow(boundsOf(points), outlineReach(item));
  return [
    { x: b.x, y: b.y },
    { x: b.r, y: b.y },
    { x: b.r, y: b.b },
    { x: b.x, y: b.b },
  ];
}

/** Slide-space corners of a shape's box, its outline, and its text frame (text can overflow the box). */
function contentCorners(
  item: Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>
): { x: number; y: number }[] {
  const out = boxCorners(item);
  if (item.kind === 'shape' && item.text && item.textMatrix && hasCharacters(item.text)) {
    const b = item.text.box;
    const h = Math.max(b.height, item.text.contentHeight);
    const frame: Matrix = [1, 0, 0, 1, b.x, b.y];
    // Rotated text can extend either way: take the box as a square of its larger side.
    const side = item.text.rotation % 180 !== 0 ? Math.max(b.width, h) : 0;
    out.push(...cornersUnder(multiplyMatrix(item.textMatrix, frame), side || b.width, side || h));
  }
  return out;
}

/** Paint a slide into a new `<svg>` sized by viewBox in EMU. The caller sets CSS size. */
export function paintSlide(doc: Document, layout: SlideLayout, opts: PaintOptions): SVGSVGElement {
  const { width, height } = layout.slide;
  const root = doc.createElementNS(SVG_NS, 'svg') as SVGSVGElement;
  root.setAttribute('viewBox', `0 0 ${width} ${height}`);
  root.setAttribute('preserveAspectRatio', 'none');
  root.setAttribute('data-slide-part', layout.slide.partName);
  const defs = new Defs(doc, opts.idPrefix);
  root.appendChild(defs.element);
  const ctx: PaintContext = {
    defs,
    images: opts.images,
    ...(opts.imageSize ? { imageSize: opts.imageSize } : {}),
    sourcePart: layout.slide.backgroundPart,
  };
  const bg = fillPaint(layout.slide.background, width, height, ctx);
  root.appendChild(
    svg(doc, 'rect', {
      width,
      height,
      fill: bg.paint === 'none' ? '#ffffff' : bg.paint,
      'fill-opacity': bg.opacity < 1 ? num(bg.opacity) : undefined,
      'data-kind': 'background',
    })
  );
  const content = svg(doc, 'g', { 'data-kind': 'content' });
  for (const item of layout.items) {
    const painted = paintItem(doc, item, ctx, opts, 1);
    if (painted) content.appendChild(painted);
  }
  root.appendChild(content);
  return root;
}

export { toCss };
