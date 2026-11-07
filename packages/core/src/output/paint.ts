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
  isOpenPath,
  lineEndReach,
  multiply as multiplyMatrix,
  outsideDashArray,
  planStroke,
  strokeWidth,
  type GeometryPath,
  type PathCommand,
  type StrokePlan,
  type LayoutItem,
  type Matrix,
  type SlideLayout,
} from '../layout/index.ts';
import { effectDefs, effectElement, shadowCopy, type EffectSource } from './paint-effects.ts';
import { boundsOf, cornersUnder, grow, PX } from './effect-transforms.ts';
import {
  fillPaint,
  pictureImage,
  pictureRect,
  pictureUrl,
  strokeAttrs,
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
  /**
   * The alt text of an element. With it, each picture is exposed to assistive technology: a
   * picture with a description is an image named by it, and a decorative picture or one with
   * no description is hidden. Without it, the paint has no ARIA (thumbnails and exports).
   */
  readonly altText?: AltTextPort;
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
 * geometry. Its dashes come from `outsideDashArray`, as in the PDF export: the dash period of a
 * centered line stays, and round dots draw as squares. Dashes outside a picture are not checked
 * against a reference PDF (approximation `picture-border-dots`).
 */
function outsideOutline(
  doc: Document,
  ctx: PaintContext,
  d: string,
  stroke: Record<string, string | number | undefined>,
  line: ResolvedLine,
  groupScale: number,
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
  const dashes = outsideDashArray(line, strokeWidth(line, groupScale));
  return svg(doc, 'path', {
    d,
    fill: 'none',
    ...stroke,
    'stroke-width': num(width * 2),
    ...(dashes ? { 'stroke-linecap': 'butt', 'stroke-dasharray': dashes.map(num).join(' ') } : {}),
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

/** The fill of the placeholder that a picture the paint cannot draw shows. */
const PICTURE_PLACEHOLDER = '#e8e8e8';

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
  const paintFor = shapePaints(item, ctx);
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
      out.push(outsideOutline(doc, ctx, d, stroke, line, groupScale, w, h));
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
    const fill =
      path.fill === 'none'
        ? { paint: 'none', opacity: 1 }
        : lostPicture
          ? { paint: PICTURE_PLACEHOLDER, opacity: 1 }
          : paintFor(el.fill, d, path.fill);
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
        ? compoundStroke(ctx.defs, strokeD, line.compound, stroke, w, h)
        : null;
    const own = overlay ? { paint: 'none', opacity: 1 } : fill;
    const filled = own.paint !== 'none';
    if (compound || (strokeD !== d && filled)) {
      // The fill and the stroke draw apart: the bands of a compound line, or a stroke that
      // stops short under a line end. Line ends stay on the center line.
      if (filled)
        out.push(
          svg(doc, 'path', {
            d,
            fill: own.paint,
            'fill-opacity': own.opacity < 1 ? num(own.opacity) : undefined,
            'fill-rule': 'evenodd',
            stroke: 'none',
          })
        );
      out.push(compound ?? svg(doc, 'path', { d: strokeD, fill: 'none', ...stroke }), ...extras);
      continue;
    }
    out.push(
      svg(doc, 'path', {
        d: strokeD,
        fill: own.paint,
        'fill-opacity': own.opacity < 1 ? num(own.opacity) : undefined,
        'fill-rule': 'evenodd',
        ...(path.stroke ? stroke : { stroke: 'none' }),
      }),
      ...extras
    );
  }
  return out;
}

/** SVG path data for path commands. */
function commandData(commands: readonly PathCommand[]): string {
  return pathData({ commands, fill: 'none', stroke: true });
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
      // Glow, inner shadow, reflection, and soft edge paint only through the installed effects
      // renderer (`@pptx-editor.dev/pro/drawing`); the outer shadow is the engine's.
      const pro = effects ? drawingRenderers().effects : undefined;
      const ids = ctx.defs;
      const body = svg(doc, 'g', { transform: matrixAttr(item.matrix) });
      for (const node of paintGeometry(doc, item, local, parentScale)) body.appendChild(node);
      const bodyFx =
        pro && effects && (effects.softEdge || effects.innerShadow)
          ? pro.body(effects, boundsOf(boxCorners(item)), ids)
          : null;
      if (bodyFx) effectDefs(ctx.defs, bodyFx.defs);
      const content = svg(doc, 'g', {}, [
        bodyFx ? svg(doc, 'g', { filter: `url(#${bodyFx.filter})` }, [body]) : body,
      ]);
      // An empty body's caret line draws nothing: no group for it.
      if (item.kind === 'shape' && item.text && item.textMatrix && hasCharacters(item.text))
        content.appendChild(
          svg(doc, 'g', { transform: matrixAttr(item.textMatrix) }, [
            paintText(doc, item.text, local),
          ])
        );
      const copies = effects?.outerShadow || (pro && (effects?.reflection || effects?.glow));
      if (!effects || !copies) {
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
      const reflection = pro && effects.reflection ? pro.reflection(effects, source, ids) : null;
      if (reflection) {
        effectDefs(ctx.defs, reflection.defs);
        outer.appendChild(effectElement(doc, reflection.node));
      }
      const shadow = shadowCopy(ctx.defs, effects, source);
      if (shadow) outer.appendChild(shadow);
      const glow = pro && effects.glow ? pro.glow(effects, boundsOf(source.corners), ids) : null;
      if (glow) effectDefs(ctx.defs, glow.defs);
      outer.appendChild(
        glow ? svg(doc, 'g', { filter: `url(#${glow.filter})` }, [content]) : content
      );
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
  const ends =
    item.kind === 'shape' && item.geometry.paths.some((p) => p.stroke && isOpenPath(p.commands));
  return Math.max(width * join, ends ? lineEndReach(line) : 0);
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

/** The painted node of one item and the definitions it added. */
interface PaintedItem {
  readonly node: SVGElement | null;
  readonly defs: readonly SVGElement[];
}

/**
 * Paints the slides of one view again and again, and reuses the nodes of the items it painted
 * before. A layout item is an immutable value, and the layout keeps an item's object while its
 * shape is unchanged (`layoutSlide`), so the same item paints the same nodes: an edit paints the
 * items it changed, not the whole slide. `paintSlide` paints from nothing, for exports,
 * thumbnails, and the fidelity check; both paint each item with `paintItem`.
 *
 * Nothing else may change the nodes it returns, except for a while: a drag preview hides nodes
 * and must put them back (`PaintedPreview.clear`) before the next `paint`. @internal
 */
export class SlidePainter {
  readonly #doc: Document;
  readonly #opts: PaintOptions;
  readonly #defs: Defs;
  #root: SVGSVGElement | null = null;
  #content: SVGElement | null = null;
  #background: PaintedItem | null = null;
  #items = new Map<LayoutItem, PaintedItem>();
  /** Items that a layout held twice: their second paint is never reused. */
  #twice: PaintedItem[] = [];
  #slide = '';

  constructor(doc: Document, opts: PaintOptions) {
    this.#doc = doc;
    this.#opts = opts;
    this.#defs = new Defs(doc, opts.idPrefix);
  }

  /** The prefix of the ids this painter writes. */
  get idPrefix(): string {
    return this.#opts.idPrefix;
  }

  /** Whether this painter paints with `opts`: the same ports and id prefix. */
  paints(opts: PaintOptions): boolean {
    const a = this.#opts;
    return (
      a.images === opts.images &&
      a.imageSize === opts.imageSize &&
      a.idPrefix === opts.idPrefix &&
      a.frameLabel === opts.frameLabel &&
      a.hitTargets === opts.hitTargets &&
      a.altText === opts.altText
    );
  }

  /**
   * Whether the last `paint` left items out because they passed its budget: a later `paint` of
   * the same layout adds them.
   */
  get pending(): boolean {
    return this.#pending;
  }

  #pending = false;

  /**
   * Paint `layout`, reusing the nodes of items painted before. Returns the same `<svg>`. With a
   * `budget`, it paints at most that many new items, in paint order, and keeps every item it
   * painted before (`pending` tells when some are left): a slide of 24,000 shapes then goes on
   * screen over several frames, and no frame blocks for seconds.
   */
  paint(layout: SlideLayout, budget = Infinity): SVGSVGElement {
    const doc = this.#doc;
    const { width, height } = layout.slide;
    const slide = `${layout.slide.partName} ${width}x${height}`;
    if (!this.#root || !this.#content || this.#slide !== slide) {
      this.#root = doc.createElementNS(SVG_NS, 'svg') as SVGSVGElement;
      this.#root.setAttribute('viewBox', `0 0 ${width} ${height}`);
      this.#root.setAttribute('preserveAspectRatio', 'none');
      this.#root.setAttribute('data-slide-part', layout.slide.partName);
      this.#defs.element.replaceChildren();
      this.#root.appendChild(this.#defs.element);
      this.#content = svg(doc, 'g', { 'data-kind': 'content' });
      this.#root.appendChild(this.#content);
      this.#background = null;
      this.#items.clear();
      this.#twice = [];
      this.#slide = slide;
    }
    const root = this.#root;
    const content = this.#content;
    const ctx: PaintContext = {
      defs: this.#defs,
      images: this.#opts.images,
      ...(this.#opts.imageSize ? { imageSize: this.#opts.imageSize } : {}),
      sourcePart: layout.slide.backgroundPart,
    };
    // The background is one rectangle: it paints each time.
    const drop = (painted: PaintedItem): void => {
      painted.node?.remove();
      for (const d of painted.defs) d.remove();
    };
    if (this.#background) drop(this.#background);
    for (const painted of this.#twice.splice(0)) drop(painted);
    const background = this.#defs.capture(() => {
      const bg = fillPaint(layout.slide.background, width, height, ctx);
      return svg(doc, 'rect', {
        width,
        height,
        fill: bg.paint === 'none' ? '#ffffff' : bg.paint,
        'fill-opacity': bg.opacity < 1 ? num(bg.opacity) : undefined,
        'data-kind': 'background',
      });
    });
    this.#background = { node: background.value, defs: background.defs };
    root.insertBefore(background.value, content);
    const next = new Map<LayoutItem, PaintedItem>();
    const nodes: SVGElement[] = [];
    const paintFresh = (item: LayoutItem): PaintedItem => {
      const fresh = this.#defs.capture(() => paintItem(doc, item, ctx, this.#opts, 1));
      return { node: fresh.value, defs: fresh.defs };
    };
    let left = budget;
    this.#pending = false;
    // The earlier nodes of the shapes that a budget leaves out (`itemKey`): each stays on
    // screen in its place until its shape paints again (a font load or a theme edit makes
    // every item new).
    let earlier: Map<string, [LayoutItem, PaintedItem]> | null = null;
    let current: Set<LayoutItem> | null = null;
    for (const item of layout.items) {
      let painted: PaintedItem;
      const kept = next.has(item) ? undefined : this.#items.get(item);
      if (!kept && left <= 0) {
        this.#pending = true;
        earlier ??= new Map([...this.#items].map(([i, p]) => [itemKey(i), [i, p]] as const));
        current ??= new Set(layout.items);
        const stand = earlier.get(itemKey(item));
        if (stand && !next.has(stand[0]) && !current.has(stand[0])) {
          next.set(stand[0], stand[1]);
          if (stand[1].node) nodes.push(stand[1].node);
        }
        continue;
      }
      if (next.has(item)) {
        // An item twice in one layout paints twice: one node cannot stand in two places.
        painted = paintFresh(item);
        this.#twice.push(painted);
        left -= 1;
      } else {
        painted = kept ?? paintFresh(item);
        if (!kept) left -= 1;
        next.set(item, painted);
      }
      if (painted.node) nodes.push(painted.node);
    }
    for (const [item, painted] of this.#items) if (!next.has(item)) drop(painted);
    this.#items = next;
    // Put the nodes in paint order, moving only the ones out of place.
    nodes.forEach((node, i) => {
      const at = content.childNodes[i] ?? null;
      if (at !== node) content.insertBefore(node, at);
    });
    while (content.childNodes.length > nodes.length) content.lastChild?.remove();
    return root;
  }
}

/**
 * An item's place in the document: its part, its layer, and its node id. Node ids are numbered
 * per part, so a master shape and a slide shape can share one.
 */
const itemKey = (item: LayoutItem): string =>
  `${item.element.partName}\u0000${item.element.layer}\u0000${item.element.nodeId}`;

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
