// Paint a laid-out slide as SVG. The painter reads layout records only; it never resolves a
// property or re-derives geometry. The editor and the fidelity harness both call this, so there
// is no second renderer for tests.

import {
  toCss,
  type BlendMode,
  type ResolvedEffects,
  type ResolvedFill,
  type ResolvedLine,
} from '../cascade/index.ts';
import {
  apply as applyMatrix,
  ARROW_MITER_LIMIT,
  drawingRenderers,
  hasCharacters,
  pictureBorder,
  planStroke,
  type GeometryPath,
  type StrokePlan,
  type LayoutItem,
  type Matrix,
  type SlideLayout,
} from '../layout/index.ts';
import {
  effectDefs,
  effectElement,
  facetElements,
  groupPaintsEffects,
  ownText,
  paintGroupEffects,
  shadowCopy,
  type EffectSource,
} from './paint-effects.ts';
import { threeDColor, threeDFaceItem } from './painter-math.ts';
import { boundsOf, PX } from './effect-transforms.ts';
import { boxCorners, contentCorners } from './item-corners.ts';
import {
  fillPaint,
  pictureImage,
  pictureRect,
  pictureUrl,
  slideSpaceFill,
  strokeAttrs,
  pathBounds,
  tileArea,
  tiledPictureContent,
  type ImageSizePort,
  type FillContext,
  type ImageUrlPort,
  type Paint,
  type PaintContext,
} from './paint-fill.ts';
import { compoundStroke } from './compound-lines.ts';
import { paintTable } from './paint-table.ts';
import { paintText, type Shows } from './paint-text.ts';
import { paintBudgets } from './paint-budgets.ts';
import { Defs, matrixAttr, num, svg, SVG_NS } from './svg.ts';
import { commandData, mapCommands, pathData } from './path-data.ts';

export { pathData } from './path-data.ts';

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
  /**
   * The alt text of an element. With it, each picture is exposed to assistive technology: a
   * picture with a description is an image named by it, and a decorative picture or one with
   * no description is hidden. Without it, the paint has no ARIA (thumbnails and exports).
   */
  readonly altText?: AltTextPort;
  /**
   * The part of slide space that can show, in EMU. A text line or a table cell wholly outside
   * it does not paint, unless its shape casts copies of itself (a shadow, a reflection, a
   * glow). Without it, everything paints. A thumbnail shows the slide; the editing surface
   * shows what its view can reach.
   */
  readonly view?: ViewRect;
}

/** A rectangle of slide space, in EMU. */
export interface ViewRect {
  /** The left edge, from the slide's left edge. */
  readonly x: number;
  /** The top edge, from the slide's top edge. */
  readonly y: number;
  /** The width. */
  readonly width: number;
  /** The height. */
  readonly height: number;
}

/**
 * Whether a box can show in `view`: its corners map to the slide by `m`. A box whose corners
 * do not map to finite points always shows.
 */
function showsIn(view: ViewRect | undefined, m: Matrix | undefined): Shows | undefined {
  if (!view || !m) return undefined;
  return (corners) => {
    let x0 = Infinity;
    let y0 = Infinity;
    let x1 = -Infinity;
    let y1 = -Infinity;
    for (const c of corners) {
      const p = applyMatrix(m, c.x, c.y);
      if (!Number.isFinite(p.x) || !Number.isFinite(p.y)) return true;
      x0 = Math.min(x0, p.x);
      y0 = Math.min(y0, p.y);
      x1 = Math.max(x1, p.x);
      y1 = Math.max(y1, p.y);
    }
    return x1 >= view.x && x0 <= view.x + view.width && y1 >= view.y && y0 <= view.y + view.height;
  };
}

/** Resolves an element (`p:cNvPr/@descr` and the `adec:decorative` extension) to its alt text. */
export type AltTextPort = (
  part: string,
  nodeId: string
) => { readonly description: string; readonly decorative: boolean } | undefined;

/**
 * The ARIA of a picture (WAI-ARIA `img` role, after the WAI images tutorial): its description as
 * the name, or hidden when decorative or without a description. Pixels do not change.
 */
function pictureAria(
  el: { readonly partName: string; readonly nodeId: string },
  altText: AltTextPort
): Record<string, string> {
  const alt = altText(el.partName, el.nodeId);
  const name = alt && !alt.decorative ? alt.description.trim() : '';
  return name ? { role: 'img', 'aria-label': name } : { 'aria-hidden': 'true' };
}

/** The uniform scale a matrix applies (for line widths inside scaled groups). */
function matrixScale(m: Matrix): number {
  return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1;
}

/**
 * The most paths of one shape that draw a fill overlay. Each overlay is an isolated group (an
 * offscreen layer); a shape with more paths draws no overlay, and the report says so.
 * @internal `pptx-to-pdf` applies the same limit.
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
      banded = paintBudgets(ctx.defs).bands.allow(paths.length, commands);
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
  const { x0, y0, x1, y1 } = pathBounds(paths, w, h);
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
 * Picture content kept inside one filled path: the box of a rectangle as a nested viewport, which
 * a printed page clips with a path, and any other path with a mask (`geometryMask`).
 */
function keepInside(
  doc: Document,
  ctx: PaintContext,
  item: Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>,
  path: GeometryPath,
  d: string,
  content: SVGElement
): SVGElement {
  const { geometry, xfrm } = item.element;
  const box =
    geometry.kind === 'preset' && geometry.preset === 'rect' && item.geometry.paths.length === 1;
  return box
    ? svg(doc, 'svg', { width: num(xfrm.cx), height: num(xfrm.cy), overflow: 'hidden' }, [content])
    : svg(doc, 'g', { mask: geometryMask(ctx.defs, [path], d, xfrm.cx, xfrm.cy) }, [content]);
}

/**
 * The slide background: one rectangle, or the images of a tiled picture (`tiledPictureContent`)
 * in the slide's viewport. Both carry `data-kind="background"`.
 */
export function paintBackground(
  doc: Document,
  fill: ResolvedFill,
  width: number,
  height: number,
  ctx: PaintContext
): SVGElement {
  const content = tiledPictureContent(fill, width, height, ctx);
  if (content)
    return svg(doc, 'svg', { width, height, overflow: 'hidden', 'data-kind': 'background' }, [
      content,
    ]);
  const bg = fillPaint(fill, width, height, ctx);
  return svg(doc, 'rect', {
    width,
    height,
    fill: bg.paint === 'none' ? '#ffffff' : bg.paint,
    'fill-opacity': bg.opacity < 1 ? num(bg.opacity) : undefined,
    'data-kind': 'background',
  });
}

/**
 * A picture's outline. It lies outside the picture, as reference PDFs show: the outline of the
 * geometry grown by the line width, with a shape's dashes, caps, and joins, as in the PDF export
 * (`layout/line-strokes.ts`, `pictureBorder`). A solid border on a geometry other than a
 * rectangle is a stroke twice as wide, of which a mask hides the half inside the geometry.
 * Dashes outside a geometry other than a rectangle are not checked against a reference PDF
 * (approximation `picture-border-dots`).
 */
function outsideOutline(
  doc: Document,
  ctx: PaintContext,
  item: Extract<LayoutItem, { kind: 'picture' }>,
  d: string,
  stroke: Record<string, string | undefined>,
  line: ResolvedLine,
  groupScale: number
): SVGElement[] {
  const { geometry, xfrm } = item.element;
  const { cx: w, cy: h } = xfrm;
  const rect = geometry.kind === 'preset' && geometry.preset === 'rect';
  const paths = item.geometry.paths.map((path) => path.commands);
  const border = pictureBorder(paths, rect, line, w, h, groupScale);
  const width = Number(stroke['stroke-width']) || 0;
  const mask = (): string => {
    const pad = width * 2 + Math.max(w, h);
    const box = { x: -pad, y: -pad, width: w + 2 * pad, height: h + 2 * pad };
    const id = ctx.defs.id('outside');
    const shapes = [
      svg(doc, 'rect', { ...box, fill: '#fff' }),
      svg(doc, 'path', { d, fill: '#000' }),
    ];
    ctx.defs.add(svg(doc, 'mask', { id, maskUnits: 'userSpaceOnUse', ...box }, shapes));
    return `url(#${id})`;
  };
  if (border.kind === 'doubled')
    return [
      svg(doc, 'path', {
        d,
        fill: 'none',
        ...stroke,
        'stroke-width': num(width * 2),
        mask: mask(),
      }),
    ];
  const clip = border.clip ? mask() : undefined;
  return border.paths.map((plan) =>
    svg(doc, 'path', {
      d: pathData({ commands: plan.commands, fill: 'none', stroke: true }),
      fill: 'none',
      ...withDashPhase(stroke, plan),
      mask: clip,
    })
  );
}

/** The stroke attributes of a path whose dash pattern starts at `plan.dashPhase`. */
function withDashPhase(
  stroke: Record<string, string | undefined>,
  plan: StrokePlan | undefined
): Record<string, string | undefined> {
  return plan?.dashPhase && stroke['stroke-dasharray']
    ? { ...stroke, 'stroke-dashoffset': num(plan.dashPhase) }
    : stroke;
}

/** The context with a shape's box-to-slide matrix, for pattern tiles on the slide's grid. */
function withFillMatrix(
  item: Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>,
  ctx: PaintContext
): FillContext {
  return { ...ctx, fillMatrix: item.matrix };
}

/** The fill of the placeholder that a picture the paint cannot draw shows. */
const PICTURE_PLACEHOLDER = '#e8e8e8';

function paintGeometry(
  doc: Document,
  item: Extract<LayoutItem, { kind: 'shape' } | { kind: 'picture' }>,
  shapeCtx: PaintContext,
  groupScale: number,
  front?: Matrix
): SVGElement[] {
  const ctx = withFillMatrix(item, shapeCtx);
  const el = item.element;
  const { cx: w, cy: h } = el.xfrm;
  const out: SVGElement[] = [];
  const line = item.paintLine ?? el.line;
  const stroke = line ? strokeAttrs(line, w, h, ctx, groupScale) : { stroke: 'none' };
  const paintFor = shapePaints(item, ctx);
  // A pattern fill of a turned, flipped, or scaled box draws in slide space (`slideSpaceFill`),
  // also under the turned face of a 3-D camera (`front`), which the body's parent applies.
  const shown = el.fill.type === 'group' ? ctx.groupFill : el.fill;
  const slideFill =
    item.kind === 'shape' && shown?.type === 'pattern'
      ? slideSpaceFill(item.matrix, front)
      : undefined;
  const fillPaintFor = slideFill ? shapePaints(item, shapeCtx) : paintFor;
  const fillElement = (
    path: GeometryPath,
    d: string,
    attrs: Record<string, string | undefined>
  ): SVGElement =>
    slideFill
      ? svg(doc, 'g', { transform: slideFill.transform }, [
          svg(doc, 'path', {
            d: pathData({ ...path, commands: mapCommands(path.commands, slideFill.matrix) }),
            ...attrs,
          }),
        ])
      : svg(doc, 'path', { d, ...attrs });
  const effects = item.geometry.paths.length <= MAX_OVERLAY_PATHS ? el.effects : undefined;
  if (item.kind === 'picture') {
    const fill = item.element.fill;
    const url = fill.type === 'picture' ? pictureUrl(ctx, fill) : undefined;
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
          fill: PICTURE_PLACEHOLDER,
          stroke: '#b0b0b0',
          'stroke-width': 9525,
        })
      );
    if (line && stroke.stroke !== 'none')
      out.push(...outsideOutline(doc, ctx, item, d, stroke, line, groupScale));
    return out;
  }
  // A picture fill whose image does not resolve (a missing or damaged part) fills the geometry
  // with the picture placeholder of a `p:pic`, so the shape does not disappear.
  const lostPicture = el.fill.type === 'picture' && !pictureUrl(ctx, el.fill);
  for (const path of item.geometry.paths) {
    // Line ends and square dash caps belong to open, stroked paths only.
    const plan =
      path.stroke && line && stroke.stroke !== 'none'
        ? planStroke(path.commands, line, groupScale)
        : undefined;
    // The fill takes the whole path; only the stroke stops short under a line end.
    const d = pathData(path);
    if (!d) continue;
    const strokeD =
      plan && plan.commands !== path.commands ? pathData({ ...path, commands: plan.commands }) : d;
    const extras = plan ? strokeExtras(doc, plan, stroke) : [];
    const pathStroke = withDashPhase(stroke, plan);
    // A tiled picture draws as images inside the path (`tiledPictureContent`); the stroke follows.
    const content =
      path.fill === 'none' || lostPicture
        ? undefined
        : tiledPictureContent(el.fill, w, h, ctx, tileArea(path, w, h));
    if (content) {
      const kept = keepInside(doc, ctx, item, path, d, content);
      const over = overlayPath(doc, d, effects, paintFor);
      out.push(over ? isolated(doc, [kept, over]) : kept);
    }
    const fill =
      path.fill === 'none' || content
        ? { paint: 'none', opacity: 1 }
        : lostPicture
          ? { paint: PICTURE_PLACEHOLDER, opacity: 1 }
          : fillPaintFor(el.fill, d, path.fill);
    const overlay = fill.paint === 'none' ? null : overlayPath(doc, d, effects, paintFor);
    if (overlay) {
      // The overlay blends with this fill only (design D5 of `fill-fidelity`); the outline draws on top.
      out.push(
        isolated(doc, [
          fillElement(path, d, {
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
        ? compoundStroke(
            ctx.defs,
            strokeD,
            line.compound,
            pathStroke,
            w,
            h,
            plan?.commands ?? path.commands
          )
        : null;
    const own = overlay ? { paint: 'none', opacity: 1 } : fill;
    const filled = own.paint !== 'none';
    if (compound || (strokeD !== d && filled) || (slideFill && filled)) {
      // The fill and the stroke draw apart: the bands of a compound line, a stroke that stops
      // short under a line end, or a fill in slide space. Line ends stay on the center line.
      if (filled)
        out.push(
          fillElement(path, d, {
            fill: own.paint,
            'fill-opacity': own.opacity < 1 ? num(own.opacity) : undefined,
            'fill-rule': 'evenodd',
            stroke: 'none',
          })
        );
      if (compound) out.push(compound);
      else if (path.stroke) out.push(svg(doc, 'path', { d: strokeD, fill: 'none', ...pathStroke }));
      out.push(...extras);
      continue;
    }
    out.push(
      svg(doc, 'path', {
        d: strokeD,
        fill: own.paint,
        'fill-opacity': own.opacity < 1 ? num(own.opacity) : undefined,
        'fill-rule': 'evenodd',
        ...(path.stroke ? pathStroke : { stroke: 'none' }),
      }),
      ...extras
    );
  }
  return out;
}

/**
 * The line ends and square dash caps of a stroke, in the stroke's paint: filled heads, open
 * arrows stroked as wide as the line, and solid caps with butt ends.
 */
function strokeExtras(
  doc: Document,
  plan: StrokePlan,
  stroke: Record<string, string | undefined>
): SVGElement[] {
  const paint = stroke.stroke;
  const opacity = stroke['stroke-opacity'];
  const out: SVGElement[] = [];
  for (const end of plan.ends)
    out.push(
      end.kind === 'fill'
        ? svg(doc, 'path', {
            d: commandData(end.commands),
            fill: paint,
            'fill-opacity': opacity,
            stroke: 'none',
            'data-line-end': '',
          })
        : svg(doc, 'path', {
            d: commandData(end.commands),
            fill: 'none',
            stroke: paint,
            'stroke-opacity': opacity,
            'stroke-width': num(end.width),
            'stroke-linecap': 'round',
            'stroke-linejoin': 'miter',
            'stroke-miterlimit': ARROW_MITER_LIMIT,
            'data-line-end': '',
          })
    );
  for (const cap of plan.caps)
    out.push(
      svg(doc, 'path', {
        d: commandData(cap),
        fill: 'none',
        stroke: paint,
        'stroke-opacity': opacity,
        'stroke-width': stroke['stroke-width'],
        'stroke-linecap': 'butt',
      })
    );
  return out;
}

/** Paint one layout item. @internal */
export function paintItem(
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
      const g = groupNode(doc, item, opts);
      if (!groupPaintsEffects(item)) {
        paintChildren(doc, item, g, ctx, opts, 0, Infinity);
        return g;
      }
      // A copy of the content can carry a line from out of view into it: every line paints.
      const inner = svg(doc, 'g');
      paintChildren(doc, item, inner, ctx, { ...opts, view: undefined }, 0, Infinity);
      paintGroupEffects(doc, item, g, inner, ctx);
      return g;
    }
    case 'frame': {
      if (item.table)
        return svg(
          doc,
          'g',
          { ...common, 'data-kind': 'table', transform: matrixAttr(item.matrix) },
          paintTable(
            doc,
            item.table,
            item.element.table?.background ?? { type: 'none' },
            local,
            showsIn(opts.view, item.matrix)
          )
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
      // A diagram draws its saved drawing (ECMA-376 Part 1, 21.4); the frame holds the hit id,
      // so the diagram selects as one read-only object.
      if (item.drawing) {
        const inner = { ...opts, hitTargets: false };
        const scale = matrixScale(item.matrix);
        const g = svg(doc, 'g', { ...common, 'data-kind': 'frame' });
        for (const child of item.drawing) {
          const painted = paintItem(doc, child, ctx, inner, scale);
          if (painted) g.appendChild(painted);
        }
        return g;
      }
      // A chart drawn by an installed renderer: its shapes and labels, under the frame's hit id.
      if (item.chart) {
        const inner = { ...opts, hitTargets: false };
        const parts: SVGElement[] = [];
        for (const child of item.chart) {
          const painted = paintItem(doc, child, local, inner, parentScale);
          if (painted) parts.push(painted);
        }
        return svg(doc, 'g', { ...common, 'data-kind': 'chart' }, parts);
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
      const aria = item.kind === 'picture' && opts.altText ? pictureAria(el, opts.altText) : {};
      const outer = svg(doc, 'g', { ...common, 'data-kind': item.kind, ...aria });
      const effects = item.element.effects;
      // Glow, inner shadow, reflection, soft edge, blur, and 3-D paint only through the installed
      // effects renderer (`@pptx-editor.dev/pro/drawing`); outer and preset shadows are the
      // engine's.
      const pro = effects ? drawingRenderers().effects : undefined;
      const ids = ctx.defs;
      const threeD =
        pro?.threeD && effects?.threeD
          ? pro.threeD({
              effects,
              paths: item.geometry.paths,
              matrix: item.matrix,
              width: el.xfrm.cx,
              height: el.xfrm.cy,
              color: threeDColor(item.element.fill),
            })
          : null;
      const extrusion = threeD ? facetElements(doc, threeD.behind, SEAM) : [];
      // A face turned edge-on paints only its extrusion: it has no area to paint or to copy.
      if (threeD && !threeD.showsFace) {
        for (const node of extrusion) outer.appendChild(node);
        return outer;
      }
      const body = svg(doc, 'g', { transform: matrixAttr(item.matrix) });
      if (threeD) for (const node of facetElements(doc, threeD.under, SEAM)) body.appendChild(node);
      // A turned camera turns the face and its text about the shape's center, and the regions
      // of the effects that copy the face turn with it.
      const front = threeD?.front.some((v, i) => v !== IDENTITY[i]) ? threeD.front : undefined;
      // A solid face takes the scene's light (ECMA-376 Part 1, 20.1.5).
      const faceItem = threeDFaceItem(item, threeD);
      for (const node of paintGeometry(doc, faceItem, local, parentScale, front))
        body.appendChild(node);
      // The bevel shades the geometry, under the text (ECMA-376 Part 1, 20.1.5.1).
      if (threeD) for (const node of facetElements(doc, threeD.over, SEAM)) body.appendChild(node);
      const bodyFx =
        pro && effects && (effects.softEdge || effects.innerShadow)
          ? pro.body(effects, boundsOf(boxCorners(item)), ids)
          : null;
      if (bodyFx) effectDefs(ctx.defs, bodyFx.defs);
      const copies = effects?.outerShadow || (pro && (effects?.reflection || effects?.glow));
      let cornerCache: { x: number; y: number }[] | undefined;
      const cornersOf = (): { x: number; y: number }[] =>
        (cornerCache ??= front
          ? contentCorners(item).map((q) => applyMatrix(front, q.x, q.y))
          : contentCorners(item));
      const blurFx =
        pro?.blur && effects?.blur ? pro.blur(effects, boundsOf(cornersOf()), ids) : null;
      if (blurFx) effectDefs(ctx.defs, blurFx.defs);
      // The body and the text paint in a group of their own only for a filter or the copies to
      // refer to. Without copies, the extrusion lies under the face.
      if (!copies) for (const node of extrusion) outer.appendChild(node);
      const content = copies || blurFx ? svg(doc, 'g') : outer;
      const face = front ? svg(doc, 'g') : undefined;
      if (face && front)
        content.appendChild(svg(doc, 'g', { transform: matrixAttr(front) }, [face]));
      const parts = face ?? content;
      parts.appendChild(
        bodyFx ? svg(doc, 'g', { filter: `url(#${bodyFx.filter})` }, [body]) : body
      );
      // An empty body's caret line draws nothing: no group for it. A copy of the content (a
      // shadow, a reflection, a glow), a blur, or a turned face can carry a line from out of
      // view into it: every line of such a shape paints.
      if (item.kind === 'shape' && item.text && item.textMatrix && hasCharacters(item.text)) {
        const shows =
          effects?.outerShadow || effects?.reflection || effects?.glow || blurFx || threeD
            ? undefined
            : showsIn(opts.view, item.textMatrix);
        const text = paintText(doc, item.text, local, shows, matrixAttr(item.textMatrix));
        // `ownText` links the body to its shape, for the clips of a soft-edged group (`softEdged`).
        parts.appendChild(ownText(item, text));
      }
      const shown = blurFx
        ? svg(doc, 'g', { filter: `url(#${blurFx.filter})` }, [content])
        : content;
      if (!effects || !copies) {
        if (shown !== outer) outer.appendChild(shown);
        return outer;
      }
      const corners = cornersOf();
      // A turned extrusion casts its shadows and its glow from its back.
      const back = threeD?.backOffset;
      const castFrom = (m: Matrix): Matrix =>
        back ? [m[0], m[1], m[2], m[3], m[4] + back.x, m[5] + back.y] : m;
      // Shadow, reflection, and glow use the content, text included (design D1 of `effects-fidelity`).
      const contentId = ctx.defs.id('content');
      content.setAttribute('id', contentId);
      const faceId = face ? ctx.defs.id('face') : undefined;
      if (face && faceId) face.setAttribute('id', faceId);
      const source: EffectSource = {
        contentId,
        matrix: item.matrix,
        width: el.xfrm.cx,
        height: el.xfrm.cy,
        corners,
        ...(front && faceId
          ? { face: { id: faceId, corners: contentCorners(item), front: castFrom(front) } }
          : {}),
      };
      const reflection = pro && effects.reflection ? pro.reflection(effects, source, ids) : null;
      if (reflection) {
        effectDefs(ctx.defs, reflection.defs);
        outer.appendChild(effectElement(doc, reflection.node));
      }
      // A preset shadow's second copy lies under its first.
      const second = shadowCopy(ctx.defs, effects, source, 'secondShadow');
      if (second) outer.appendChild(second);
      const shadow = shadowCopy(ctx.defs, effects, source);
      if (shadow) outer.appendChild(shadow);
      // The glow is a copy that paints the glow alone, under the extrusion and the content. The
      // content itself paints with no filter, so its text stays text in a printed PDF.
      // The glow of a turned extrusion lies around the face and its back.
      const glowCorners = back
        ? [...corners, ...corners.map((q) => ({ x: q.x + back.x, y: q.y + back.y }))]
        : corners;
      const glow = pro && effects.glow ? pro.glow(effects, boundsOf(glowCorners), ids, true) : null;
      if (glow) {
        effectDefs(ctx.defs, glow.defs);
        let glowSource = contentId;
        if (shown !== content) {
          glowSource = ctx.defs.id('shown');
          shown.setAttribute('id', glowSource);
        }
        outer.appendChild(
          svg(doc, 'g', { filter: `url(#${glow.filter})`, 'data-effect': 'glow' }, [
            svg(doc, 'use', {
              href: `#${glowSource}`,
              transform: back ? matrixAttr([1, 0, 0, 1, back.x, back.y]) : undefined,
            }),
          ])
        );
      }
      for (const node of extrusion) outer.appendChild(node);
      outer.appendChild(shown);
      return outer;
    }
  }
}

/**
 * The number of items `item` paints: one, or for a group, a diagram, or a chart, the items in
 * it, and for a table, its cells. A paint budget (`SlidePainter.paint`) counts these, so a group
 * of thousands of shapes paints over several frames as thousands of shapes do. @internal
 */
export function paintWeight(item: LayoutItem): number {
  const known = weights.get(item);
  if (known !== undefined) return known;
  const children =
    item.kind === 'group'
      ? item.children
      : item.kind === 'frame'
        ? (item.drawing ?? item.chart)
        : undefined;
  let weight = 1 + (item.kind === 'frame' ? (item.table?.cells.length ?? 0) : 0);
  if (children) for (const child of children) weight += paintWeight(child);
  weights.set(item, weight);
  return weight;
}

const weights = new WeakMap<LayoutItem, number>();

/** The node of a group, without its children. */
export function groupNode(
  doc: Document,
  group: Extract<LayoutItem, { kind: 'group' }>,
  opts: PaintOptions
): SVGElement {
  const el = group.element;
  const hit = opts.hitTargets !== false;
  return svg(doc, 'g', {
    'data-node-id': hit ? el.nodeId : undefined,
    'data-part': hit ? el.partName : undefined,
    'data-layer': hit ? el.layer : undefined,
    'data-kind': 'group',
  });
}

/** What one `paintChildren` call painted of a group's children. */
export interface ChildPaint {
  /** The index of the first child left out: the group's `children.length` when none is. */
  readonly next: number;
  /** The weight (`paintWeight`) that it painted. */
  readonly spent: number;
}

/**
 * Paint the children of `group` into its node `g`, from child `from` on, while their weight
 * (`paintWeight`) fits in `budget`; at least one child. Answers the index of the first child
 * left out (`children.length` when none is) and the weight painted.
 */
export function paintChildren(
  doc: Document,
  group: Extract<LayoutItem, { kind: 'group' }>,
  g: SVGElement,
  ctx: PaintContext,
  opts: PaintOptions,
  from: number,
  budget: number
): ChildPaint {
  const el = group.element;
  const scale = matrixScale(group.matrix);
  // A group with `a:grpFill` passes its own group's fill on (ECMA-376 Part 1, 20.1.8.35).
  const inherits = el.fill.type === 'group';
  const groupCtx: PaintContext = {
    ...ctx,
    sourcePart: el.partName,
    groupFill: inherits ? ctx.groupFill : el.fill,
    groupSourcePart: inherits ? ctx.groupSourcePart : el.partName,
  };
  let spent = 0;
  let i = from;
  for (; i < group.children.length; i += 1) {
    const child = group.children[i]!;
    const weight = paintWeight(child);
    if (i > from && spent + weight > budget) break;
    const painted = paintItem(doc, child, groupCtx, opts, scale);
    if (painted) g.appendChild(painted);
    spent += weight;
  }
  return { next: i, spent };
}

/** The outline width that closes the antialiased seams between 3-D polygons: half a pixel. */
const SEAM = PX / 2;
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

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
  root.appendChild(paintBackground(doc, layout.slide.background, width, height, ctx));
  const content = svg(doc, 'g', { 'data-kind': 'content' });
  for (const item of layout.items) {
    const painted = paintItem(doc, item, ctx, opts, 1);
    if (painted) content.appendChild(painted);
  }
  root.appendChild(content);
  return root;
}

export { toCss };
