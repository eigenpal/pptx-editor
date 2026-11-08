// Fill, line, effect, and blip resolution. Reads DrawingML property elements (direct or from a
// theme style reference) and returns plain resolved values. Layout and paint never look at a
// fill element again: they read these records.

import {
  attr,
  boolAttr,
  child,
  children,
  elements,
  intAttr,
  NS,
  q,
  type Element,
} from '../store/index.ts';
import { colorChild, resolveColor, type ColorContext, type Rgba } from './color.ts';
import type { Theme } from './theme.ts';
import { presetShadows, type ResolvedThreeD } from './three-d.ts';

/** A rectangle inset as fractions of the bounds (`a:srcRect`, `a:fillToRect`). */
export interface RelRect {
  readonly l: number;
  readonly t: number;
  readonly r: number;
  readonly b: number;
}

/** One gradient stop. `pos` is 0–1. */
export interface GradientStop {
  readonly pos: number;
  readonly color: Rgba;
}

/** A blip (image) effect, in document order. */
export type BlipEffect =
  | { readonly kind: 'alphaModFix'; readonly amount: number }
  | { readonly kind: 'grayscale' }
  | { readonly kind: 'duotone'; readonly dark: Rgba; readonly light: Rgba }
  | { readonly kind: 'lum'; readonly bright: number; readonly contrast: number }
  | { readonly kind: 'biLevel'; readonly threshold: number }
  | { readonly kind: 'unsupported'; readonly name: string };

/** A resolved fill. */
export type ResolvedFill =
  | { readonly type: 'none' }
  | { readonly type: 'solid'; readonly color: Rgba }
  | {
      readonly type: 'linear';
      readonly stops: readonly GradientStop[];
      /** Degrees, clockwise from the +x axis (`a:lin/@ang`). */
      readonly angle: number;
      readonly scaled: boolean;
      readonly rotWithShape: boolean;
      /** The gradient's box as insets from the fill box (`a:tileRect`); negatives extend it. */
      readonly tileRect?: RelRect;
    }
  | {
      readonly type: 'path';
      readonly path: 'circle' | 'rect' | 'shape';
      readonly stops: readonly GradientStop[];
      readonly fillToRect: RelRect;
      readonly rotWithShape: boolean;
      /** The gradient's box as insets from the fill box (`a:tileRect`); negatives extend it. */
      readonly tileRect?: RelRect;
    }
  | {
      readonly type: 'picture';
      readonly rId: string;
      /** An `asvg:svgBlip` alternative, when present. */
      readonly svgRId?: string;
      readonly srcRect: RelRect;
      readonly mode: 'stretch' | 'tile';
      readonly fillRect: RelRect;
      readonly tile?: {
        readonly tx: number;
        readonly ty: number;
        readonly sx: number;
        readonly sy: number;
        readonly flip: string;
        readonly algn: string;
      };
      readonly effects: readonly BlipEffect[];
      readonly rotWithShape: boolean;
    }
  | { readonly type: 'pattern'; readonly preset: string; readonly fg: Rgba; readonly bg: Rgba }
  | { readonly type: 'group' };

/** Arrowhead description. */
export interface LineEnd {
  readonly type: string;
  readonly w: 'sm' | 'med' | 'lg';
  readonly len: 'sm' | 'med' | 'lg';
}

/** A resolved outline. Width in EMU. */
export interface ResolvedLine {
  readonly fill: ResolvedFill;
  readonly width: number;
  readonly cap: 'flat' | 'rnd' | 'sq';
  readonly join: 'round' | 'bevel' | 'miter';
  readonly miterLimit: number;
  /** Preset dash name, or `custom` with `dashes` in multiples of line width. */
  readonly dash: string;
  readonly dashes?: readonly number[];
  readonly compound: string;
  readonly head?: LineEnd;
  readonly tail?: LineEnd;
}

/** The point of a shape's box an effect scales and skews about (`a:outerShdw/@algn`). */
export type EffectAlign = 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';

/** Scale, skew, alignment, and offset shared by outer shadows and reflections. */
export interface EffectPlacement {
  /** Blur radius, EMU. */
  readonly blur: number;
  /** Offset distance, EMU, at `dir` degrees (clockwise from the x axis). */
  readonly dist: number;
  readonly dir: number;
  readonly sx: number;
  readonly sy: number;
  /** Skew angles, degrees. */
  readonly kx: number;
  readonly ky: number;
  readonly align: EffectAlign;
  /** True: the effect turns with the shape (the schema default). */
  readonly rotWithShape: boolean;
}

/** A blend mode (`ST_BlendMode`), as `a:fillOverlay` and `a:blend` use it. */
export type BlendMode = 'over' | 'mult' | 'screen' | 'darken' | 'lighten';

/** Resolved shape effects. Distances in EMU, angles in degrees. */
export interface ResolvedEffects {
  readonly outerShadow?: EffectPlacement & { readonly color: Rgba };
  /**
   * A second outer shadow, under the first: the far or opposite copy of a preset shadow
   * (`a:prstShdw`), or a preset shadow next to an `a:outerShdw`.
   */
  readonly secondShadow?: EffectPlacement & { readonly color: Rgba };
  /** A blur of the whole shape (`a:blur`): radius in EMU; `grow` lets it spread past the edge. */
  readonly blur?: { readonly radius: number; readonly grow: boolean };
  /** Bevels, extrusion, contour, and the camera (`a:scene3d`, `a:sp3d`). */
  readonly threeD?: ResolvedThreeD;
  readonly innerShadow?: {
    readonly blur: number;
    readonly dist: number;
    readonly dir: number;
    readonly color: Rgba;
  };
  readonly glow?: { readonly radius: number; readonly color: Rgba };
  readonly softEdge?: { readonly radius: number };
  /** A fill drawn over the shape's fill with a blend mode (`a:fillOverlay`). */
  readonly fillOverlay?: {
    readonly blend: BlendMode;
    readonly fill: Exclude<ResolvedFill, { type: 'group' }>;
  };
  readonly reflection?: EffectPlacement & {
    readonly startAlpha: number;
    /** Fade start and end, fractions of the reflection's extent along `fadeDir`. */
    readonly startPos: number;
    readonly endAlpha: number;
    readonly endPos: number;
    /** Fade direction, degrees. */
    readonly fadeDir: number;
  };
  /** Effect elements present but not rendered, for the fidelity report. */
  readonly unsupported: readonly string[];
}

const ALIGNS = new Set<EffectAlign>(['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br']);
const BLEND_MODES = new Set<BlendMode>(['over', 'mult', 'screen', 'darken', 'lighten']);
const MAX_EFFECT_LENGTH = 51_206_400;

/**
 * An effect length (`ST_PositiveCoordinate`: blur radii, distances, glow and soft-edge radii),
 * clamped to 0 … `MAX_EFFECT_LENGTH` so a hostile value never reaches a filter unbounded.
 */
function effectLength(e: Element, name: string, d: number): number {
  return Math.max(0, Math.min(intAttr(e, name) ?? d, MAX_EFFECT_LENGTH));
}

/** An `ST_PositiveFixedAngle` in degrees, 0 … 360. */
function positiveAngle(e: Element, name: string): number {
  const deg = ((intAttr(e, name) ?? 0) / 60_000) % 360;
  return deg < 0 ? deg + 360 : deg;
}

/** Placement attributes of `a:outerShdw` or `a:reflection`. */
function placementOf(e: Element, defaultDist: number): EffectPlacement {
  const length = (name: string, d: number): number => effectLength(e, name, d);
  const angle = (name: string): number => ((intAttr(e, name) ?? 0) / 60_000) % 360;
  // Skew angles are limited to ±90° exclusive in the schema.
  const skew = (name: string): number => Math.max(-89.9, Math.min(angle(name), 89.9));
  const scale = (name: string): number => Math.max(-100, Math.min(pct(intAttr(e, name), 1), 100));
  const align = attr(e, 'algn') as EffectAlign | undefined;
  return {
    blur: length('blurRad', 0),
    dist: length('dist', defaultDist),
    dir: angle('dir'),
    sx: scale('sx'),
    sy: scale('sy'),
    kx: skew('kx'),
    ky: skew('ky'),
    align: align && ALIGNS.has(align) ? align : 'b',
    rotWithShape: boolAttr(e, 'rotWithShape') ?? true,
  };
}

const pct = (v: number | undefined, d: number): number => (v === undefined ? d : v / 100_000);

function relRect(el: Element | undefined): RelRect {
  return {
    l: pct(intAttr(el, 'l'), 0),
    t: pct(intAttr(el, 't'), 0),
    r: pct(intAttr(el, 'r'), 0),
    b: pct(intAttr(el, 'b'), 0),
  };
}

/** The most gradient stops kept (a common gradient editor allows 10): a resource limit. */
const MAX_GRADIENT_STOPS = 64;

/** The most `a:ds` dash and space pairs kept from an `a:custDash`: a resource limit. */
const MAX_CUSTOM_DASHES = 64;

/**
 * The shortest custom dash period kept, in line widths. A shorter pattern draws as a solid line:
 * it looks solid, and it would cost a renderer a dash for every hundredth of a width.
 */
const MIN_CUSTOM_DASH_PERIOD = 0.01;

/**
 * The dash and space lengths of an `a:custDash`, in line widths, or undefined for a solid line.
 * Only the first `MAX_CUSTOM_DASHES` pairs count, a negative length is 0, and a pattern with no
 * length, or one shorter than `MIN_CUSTOM_DASH_PERIOD`, draws solid.
 */
function customDashes(custDash: Element): number[] | undefined {
  const len = (v: number | undefined): number => Math.max(0, pct(v, 1));
  const out = children(custDash, 'a:ds')
    .slice(0, MAX_CUSTOM_DASHES)
    .flatMap((ds) => [len(intAttr(ds, 'd')), len(intAttr(ds, 'sp'))]);
  const period = out.reduce((sum, v) => sum + v, 0);
  return Number.isFinite(period) && period >= MIN_CUSTOM_DASH_PERIOD ? out : undefined;
}

function stopsOf(grad: Element, ctx: ColorContext): GradientStop[] {
  const out: GradientStop[] = [];
  for (const gs of children(child(grad, 'a:gsLst'), 'a:gs')) {
    const color = resolveColor(colorChild(gs), ctx);
    if (color) out.push({ pos: Math.min(1, Math.max(0, pct(intAttr(gs, 'pos'), 0))), color });
  }
  out.sort((a, b) => a.pos - b.pos);
  if (out.length <= MAX_GRADIENT_STOPS) return out;
  // Keep an even sample that includes the first and the last stop.
  const step = (out.length - 1) / (MAX_GRADIENT_STOPS - 1);
  return Array.from({ length: MAX_GRADIENT_STOPS }, (_, i) => out[Math.round(i * step)]);
}

function blipEffects(blip: Element | undefined, ctx: ColorContext): BlipEffect[] {
  const out: BlipEffect[] = [];
  for (const e of elements(blip)) {
    if (e.ns !== NS.a) continue;
    switch (e.local) {
      case 'alphaModFix':
        out.push({ kind: 'alphaModFix', amount: pct(intAttr(e, 'amt'), 1) });
        break;
      case 'grayscl':
        out.push({ kind: 'grayscale' });
        break;
      case 'duotone': {
        const cs = elements(e).map((c) => resolveColor(c, ctx));
        if (cs[0] && cs[1]) out.push({ kind: 'duotone', dark: cs[0], light: cs[1] });
        break;
      }
      case 'lum':
        out.push({
          kind: 'lum',
          bright: pct(intAttr(e, 'bright'), 0),
          contrast: pct(intAttr(e, 'contrast'), 0),
        });
        break;
      case 'biLevel':
        out.push({ kind: 'biLevel', threshold: pct(intAttr(e, 'thresh'), 0.5) });
        break;
      case 'extLst':
        break;
      default:
        out.push({ kind: 'unsupported', name: e.local });
    }
  }
  return out;
}

/** Find the `asvg:svgBlip/@r:embed` inside a blip's extension list. */
function svgBlipId(blip: Element | undefined): string | undefined {
  for (const ext of children(child(blip, 'a:extLst'), 'a:ext')) {
    const svg = child(ext, { ns: NS.asvg, local: 'svgBlip' });
    const id = attr(svg, 'r:embed');
    if (id) return id;
  }
  return undefined;
}

/** Resolve one fill element (`a:solidFill`, `a:gradFill`, …). */
export function resolveFillElement(
  el: Element | undefined,
  ctx: ColorContext
): ResolvedFill | undefined {
  if (!el || (el.ns !== NS.a && el.ns !== NS.p)) return undefined;
  switch (el.local) {
    case 'noFill':
      return { type: 'none' };
    case 'solidFill': {
      const color = resolveColor(colorChild(el), ctx);
      return color ? { type: 'solid', color } : { type: 'none' };
    }
    case 'gradFill': {
      const stops = stopsOf(el, ctx);
      if (stops.length === 0) return { type: 'none' };
      const rotWithShape = boolAttr(el, 'rotWithShape') ?? true;
      // CT_GradientFillProperties/tileRect: the rectangle the gradient is computed in.
      const tileEl = child(el, 'a:tileRect');
      const tileRect = tileEl ? { tileRect: relRect(tileEl) } : {};
      const pathEl = child(el, 'a:path');
      if (pathEl) {
        const kind = attr(pathEl, 'path');
        return {
          type: 'path',
          path: kind === 'rect' || kind === 'shape' ? kind : 'circle',
          stops,
          fillToRect: relRect(child(pathEl, 'a:fillToRect')),
          rotWithShape,
          ...tileRect,
        };
      }
      const lin = child(el, 'a:lin');
      return {
        type: 'linear',
        stops,
        angle: (intAttr(lin, 'ang') ?? 0) / 60_000,
        scaled: boolAttr(lin, 'scaled') ?? false,
        rotWithShape,
        ...tileRect,
      };
    }
    case 'blipFill': {
      const blip = child(el, 'a:blip');
      const rId = attr(blip, 'r:embed');
      if (!rId) return { type: 'none' };
      const tile = child(el, 'a:tile');
      const stretch = child(el, 'a:stretch');
      const svgRId = svgBlipId(blip);
      return {
        type: 'picture',
        rId,
        ...(svgRId ? { svgRId } : {}),
        srcRect: relRect(child(el, 'a:srcRect')),
        mode: tile && !stretch ? 'tile' : 'stretch',
        fillRect: relRect(child(stretch, 'a:fillRect')),
        ...(tile && !stretch
          ? {
              tile: {
                tx: intAttr(tile, 'tx') ?? 0,
                ty: intAttr(tile, 'ty') ?? 0,
                sx: pct(intAttr(tile, 'sx'), 1),
                sy: pct(intAttr(tile, 'sy'), 1),
                flip: attr(tile, 'flip') ?? 'none',
                algn: attr(tile, 'algn') ?? 'tl',
              },
            }
          : {}),
        effects: blipEffects(blip, ctx),
        rotWithShape: boolAttr(el, 'rotWithShape') ?? true,
      };
    }
    case 'pattFill': {
      const fg = resolveColor(colorChild(child(el, 'a:fgClr')), ctx) ?? { r: 0, g: 0, b: 0, a: 1 };
      const bg = resolveColor(colorChild(child(el, 'a:bgClr')), ctx) ?? {
        r: 255,
        g: 255,
        b: 255,
        a: 1,
      };
      return { type: 'pattern', preset: attr(el, 'prst') ?? 'pct5', fg, bg };
    }
    case 'grpFill':
      return { type: 'group' };
  }
  return undefined;
}

const FILL_NAMES = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);

/** The fill element among a property container's children (`p:spPr`, `a:rPr`, `p:bgPr`). */
export function fillElementOf(props: Element | undefined): Element | undefined {
  for (const c of elements(props)) if (c.ns === NS.a && FILL_NAMES.has(c.local)) return c;
  return undefined;
}

/**
 * Resolve a theme style-matrix reference (`a:fillRef`, `a:lnRef`, `a:effectRef`, `p:bgRef`).
 * The reference's own color becomes `phClr` inside the referenced style.
 */
export function styleRefContext(ref: Element | undefined, ctx: ColorContext): ColorContext {
  const color = resolveColor(colorChild(ref), ctx);
  return color ? { ...ctx, placeholderColor: color } : ctx;
}

/** The theme fill a `fillRef`/`bgRef` index points at. 0 and 1000 mean no fill. */
export function themeFill(idx: number, theme: Theme): Element | undefined {
  if (idx >= 1001) return theme.backgroundFillStyles[idx - 1001];
  if (idx >= 1) return theme.fillStyles[idx - 1];
  return undefined;
}

const LINE_END_SIZES = new Set(['sm', 'med', 'lg']);

function lineEnd(el: Element | undefined): LineEnd | undefined {
  const type = attr(el, 'type');
  if (!el || !type || type === 'none') return undefined;
  const w = attr(el, 'w') ?? 'med';
  const len = attr(el, 'len') ?? 'med';
  return {
    type,
    w: (LINE_END_SIZES.has(w) ? w : 'med') as LineEnd['w'],
    len: (LINE_END_SIZES.has(len) ? len : 'med') as LineEnd['len'],
  };
}

/** Default outline width when neither the shape nor its style gives one: 0.75 pt. */
export const DEFAULT_LINE_WIDTH = 9525;

/**
 * Resolve an outline from the direct `a:ln` and the style `a:ln` it overrides, property by
 * property: direct attributes and children win; anything absent comes from the style.
 */
export function resolveLine(
  direct: Element | undefined,
  style: Element | undefined,
  directCtx: ColorContext,
  styleCtx: ColorContext
): ResolvedLine | undefined {
  if (!direct && !style) return undefined;
  const pick = (name: string): string | undefined => attr(direct, name) ?? attr(style, name);
  const childOf = (name: string): [Element | undefined, ColorContext] => {
    const d = child(direct, name);
    return d ? [d, directCtx] : [child(style, name), styleCtx];
  };
  const fillDirect = fillElementOf(direct);
  const fill = fillDirect
    ? resolveFillElement(fillDirect, directCtx)
    : resolveFillElement(fillElementOf(style), styleCtx);
  const joinEl =
    elements(direct).find((e) => e.ns === NS.a && ['round', 'bevel', 'miter'].includes(e.local)) ??
    elements(style).find((e) => e.ns === NS.a && ['round', 'bevel', 'miter'].includes(e.local));
  const [prstDash] = childOf('a:prstDash');
  const [custDash] = childOf('a:custDash');
  const width = intAttr(direct, 'w') ?? intAttr(style, 'w') ?? DEFAULT_LINE_WIDTH;
  const cap = pick('cap');
  const dashes = custDash ? customDashes(custDash) : undefined;
  const head = lineEnd(childOf('a:headEnd')[0]);
  const tail = lineEnd(childOf('a:tailEnd')[0]);
  return {
    fill: fill ?? { type: 'none' },
    width: Math.max(0, Math.min(width, 20_116_800)),
    cap: cap === 'rnd' || cap === 'sq' ? cap : 'flat',
    join: (joinEl?.local as ResolvedLine['join']) ?? 'round',
    miterLimit: pct(intAttr(joinEl, 'lim'), 8),
    dash: custDash ? (dashes ? 'custom' : 'solid') : (attr(prstDash, 'val') ?? 'solid'),
    ...(dashes ? { dashes } : {}),
    compound: pick('cmpd') ?? 'sng',
    ...(head ? { head } : {}),
    ...(tail ? { tail } : {}),
  };
}

/** Resolve an effect list (`a:effectLst`). `a:effectDag` is reported unsupported. */
export function resolveEffects(
  container: Element | undefined,
  ctx: ColorContext
): ResolvedEffects | undefined {
  if (!container) return undefined;
  const dag = child(container, 'a:effectDag');
  const list = child(container, 'a:effectLst');
  if (!list) return dag ? { unsupported: ['effectDag'] } : undefined;
  const unsupported: string[] = [];
  let out: Omit<ResolvedEffects, 'unsupported'> = {};
  for (const e of elements(list)) {
    const color = (): Rgba => resolveColor(colorChild(e), ctx) ?? { r: 0, g: 0, b: 0, a: 1 };
    switch (e.local) {
      case 'outerShdw': {
        // An outer shadow written before a preset shadow takes the first place.
        const shadow = { ...placementOf(e, 0), color: color() };
        out = out.outerShadow
          ? { ...out, outerShadow: shadow, secondShadow: out.outerShadow }
          : { ...out, outerShadow: shadow };
        break;
      }
      case 'prstShdw': {
        const [first, second] = presetShadows(e, ctx);
        const extra = out.outerShadow ? first : second;
        out = {
          ...out,
          ...(out.outerShadow ? {} : { outerShadow: first }),
          ...(extra && !out.secondShadow ? { secondShadow: extra } : {}),
        };
        break;
      }
      case 'blur':
        out = {
          ...out,
          blur: { radius: effectLength(e, 'rad', 0), grow: boolAttr(e, 'grow') ?? true },
        };
        break;
      case 'innerShdw':
        out = {
          ...out,
          innerShadow: {
            blur: effectLength(e, 'blurRad', 0),
            dist: effectLength(e, 'dist', 0),
            dir: positiveAngle(e, 'dir'),
            color: color(),
          },
        };
        break;
      case 'glow':
        out = { ...out, glow: { radius: effectLength(e, 'rad', 0), color: color() } };
        break;
      case 'softEdge':
        out = { ...out, softEdge: { radius: effectLength(e, 'rad', 0) } };
        break;
      case 'fillOverlay': {
        // CT_FillOverlayEffect: one fill (EG_FillProperties) and `blend` (ST_BlendMode).
        const fill = resolveFillElement(fillElementOf(e), ctx);
        const blend = attr(e, 'blend') as BlendMode | undefined;
        if (fill && fill.type !== 'group' && blend && BLEND_MODES.has(blend))
          out = { ...out, fillOverlay: { blend, fill } };
        else unsupported.push(e.local);
        break;
      }
      case 'reflection': {
        const fraction = (name: string, d: number): number =>
          Math.max(0, Math.min(pct(intAttr(e, name), d), 1));
        out = {
          ...out,
          reflection: {
            ...placementOf(e, 0),
            startAlpha: fraction('stA', 1),
            startPos: fraction('stPos', 0),
            endAlpha: fraction('endA', 0),
            endPos: fraction('endPos', 1),
            fadeDir: ((intAttr(e, 'fadeDir') ?? 5_400_000) / 60_000) % 360,
          },
        };
        break;
      }
      default:
        unsupported.push(e.local);
    }
  }
  return { ...out, unsupported };
}

/** `p:spPr`-like containers keep the effect list as a direct child. */
export function effectContainer(spPr: Element | undefined): Element | undefined {
  return child(spPr, 'a:effectLst') || child(spPr, 'a:effectDag') ? spPr : undefined;
}

export { q };
