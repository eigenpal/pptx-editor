// Shape Effects and Shape Styles: shadow, reflection, glow, and soft edges in `a:effectLst`, and
// the theme styles of the Shape Styles gallery in `p:style`. Each writes children in schema
// order (CT_ShapeProperties, CT_EffectList, CT_Shape), so saved files stay valid.
//
// The preset values are the gallery output recorded in `reference/quick-ref/drawingml.md`
// ("Effect presets and Shape Styles"), checked against reference files; the rendering gaps
// are listed on the public fidelity page (`docs/site/content/pptx-fidelity.mdx`).

import {
  attr,
  child,
  elements,
  intAttr,
  NS,
  REL,
  relTargetByType,
  specOf,
  xmlPartOf,
  type Element,
  type NodeSpec,
} from '../../store/index.ts';
import { inheritedPlaceholders, parseTheme, type ResolvedEffects } from '../../cascade/index.ts';
import type {
  ColorChoice,
  ReflectionPreset,
  ShadowPreset,
  ShapeEffectChoice,
  ShapeEffectsState,
  ShapeStyleChoice,
  ShapeStyleColor,
  ShapeStyleRow,
} from '../../contracts/neutral.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { specFromXml } from './shape-commands.ts';
import { colorXml, isColorChoice } from './style-commands.ts';
import {
  bevelStateOf,
  isThreeDChoice,
  rotationStateOf,
  setThreeD,
  type InheritedThreeD,
} from './three-d-commands.ts';

/** CT_EffectList children, in schema order. */
const EFFECT_ORDER = [
  'blur',
  'fillOverlay',
  'glow',
  'innerShdw',
  'outerShdw',
  'prstShdw',
  'reflection',
  'softEdge',
];
/** The `p:spPr` children that come before `a:effectLst` (CT_ShapeProperties). */
const BEFORE_EFFECTS = new Set([
  'xfrm',
  'custGeom',
  'prstGeom',
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'ln',
]);
/** The effect elements each effect kind owns: a new shadow replaces any shadow. */
const OWNED: Record<ShapeEffectChoice['kind'], ReadonlySet<string>> = {
  shadow: new Set(['innerShdw', 'outerShdw', 'prstShdw']),
  reflection: new Set(['reflection']),
  glow: new Set(['glow']),
  softEdge: new Set(['softEdge']),
  blur: new Set(['blur']),
  bevel: new Set<string>(),
  rotation3d: new Set<string>(),
};
/** The largest effect radius the chrome writes, in points (the Format pane's limits). */
const MAX_EFFECT_POINTS = 150;
const EMU_PER_POINT = 12_700;

interface ShadowSpec {
  readonly inner: boolean;
  readonly blur: number;
  readonly dist: number;
  /** Direction, 60000ths of a degree. */
  readonly dir: number;
  readonly algn?: string;
  /** Scale, 1000ths of a percent. */
  readonly scale?: number;
  /** Alpha, 1000ths of a percent; absent is opaque. */
  readonly alpha?: number;
}

const outer = (dir: number, algn: string): ShadowSpec => ({
  inner: false,
  blur: 50_800,
  dist: 38_100,
  dir,
  algn,
  alpha: 40_000,
});
const inner = (dir: number): ShadowSpec => ({
  inner: true,
  blur: 63_500,
  dist: 50_800,
  dir,
  alpha: 50_000,
});

/** The Shadow gallery: Outer ("Offset: …") and Inner ("Inside: …") presets. */
export const SHADOW_PRESETS: Readonly<Record<ShadowPreset, ShadowSpec>> = {
  outerBottomRight: outer(2_700_000, 'tl'),
  outerBottom: outer(5_400_000, 't'),
  outerBottomLeft: outer(8_100_000, 'tr'),
  outerRight: outer(0, 'l'),
  outerCenter: {
    inner: false,
    blur: 63_500,
    dist: 0,
    dir: 0,
    algn: 'ctr',
    scale: 102_000,
    alpha: 40_000,
  },
  outerLeft: outer(10_800_000, 'r'),
  outerTopRight: outer(18_900_000, 'bl'),
  outerTop: outer(16_200_000, 'b'),
  outerTopLeft: outer(13_500_000, 'br'),
  innerTopLeft: inner(13_500_000),
  innerTop: inner(16_200_000),
  innerTopRight: inner(18_900_000),
  innerLeft: inner(10_800_000),
  innerCenter: { inner: true, blur: 114_300, dist: 0, dir: 0 },
  innerRight: inner(0),
  innerBottomLeft: inner(8_100_000),
  innerBottom: inner(5_400_000),
  innerBottomRight: inner(2_700_000),
};

/** The Reflection gallery: how far the fade reaches (`endPos`) and the gap (`dist`), EMU. */
export const REFLECTION_PRESETS: Readonly<
  Record<ReflectionPreset, { readonly endPos: number; readonly dist: number }>
> = {
  tightTouching: { endPos: 35_000, dist: 0 },
  halfTouching: { endPos: 55_500, dist: 0 },
  fullTouching: { endPos: 90_000, dist: 0 },
  tight4pt: { endPos: 35_000, dist: 50_800 },
  half4pt: { endPos: 55_500, dist: 50_800 },
  full4pt: { endPos: 90_000, dist: 50_800 },
  tight8pt: { endPos: 35_000, dist: 101_600 },
  half8pt: { endPos: 55_500, dist: 101_600 },
  full8pt: { endPos: 90_000, dist: 101_600 },
};

/** The sizes the Glow and Soft Edges galleries offer, in points. */
export const GLOW_SIZES: readonly number[] = [5, 8, 11, 18];
export const SOFT_EDGE_SIZES: readonly number[] = [1, 2.5, 5, 10, 25, 50];

function shadowXml(s: ShadowSpec): string {
  const name = s.inner ? 'innerShdw' : 'outerShdw';
  const attrs = [
    `blurRad="${s.blur}"`,
    s.dist ? `dist="${s.dist}"` : '',
    s.dist ? `dir="${s.dir}"` : '',
    s.scale ? `sx="${s.scale}" sy="${s.scale}"` : '',
    s.algn ? `algn="${s.algn}"` : '',
    // The Shadow gallery writes rotWithShape="0" on outer shadows (r53-1, 2 November 2026):
    // a turned shape's shadow keeps its slide direction. CT_InnerShadowEffect has no such
    // attribute.
    s.inner ? '' : 'rotWithShape="0"',
  ]
    .filter(Boolean)
    .join(' ');
  const color = s.alpha
    ? `<a:prstClr val="black"><a:alpha val="${s.alpha}"/></a:prstClr>`
    : '<a:prstClr val="black"/>';
  return `<a:${name} ${attrs}>${color}</a:${name}>`;
}

function reflectionXml(preset: ReflectionPreset): string {
  const r = REFLECTION_PRESETS[preset];
  const dist = r.dist ? ` dist="${r.dist}"` : '';
  return `<a:reflection blurRad="6350" stA="50000" endA="300" endPos="${r.endPos}"${dist} dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>`;
}

/** A size in points as an effect radius in EMU, or a refusal for a size out of range. */
function radius(points: number): number {
  if (!Number.isFinite(points) || points <= 0 || points > MAX_EFFECT_POINTS)
    throw new CommandRefusal('errors.unsupported', 'effectSize');
  return Math.round(points * EMU_PER_POINT);
}

/**
 * A glow color: the pick with the 40% opacity of the Glow Variations, and the saturation
 * boost (`a:satMod`) that a theme color gets.
 */
function glowColorSpec(color: ColorChoice): NodeSpec {
  if (!isColorChoice(color)) throw new CommandRefusal('errors.unsupported', 'color');
  const base = specFromXml(colorXml(color));
  const extra = [
    ...(typeof color === 'string' ? [] : [specFromXml('<a:satMod val="175000"/>')]),
    specFromXml('<a:alpha val="40000"/>'),
  ];
  return { ...base, children: [...(base.children ?? []), ...extra] };
}

const DEFAULT_GLOW_COLOR: ColorChoice = { theme: 'accent1' };

/** The spec of the effect element a choice writes, or null to remove the effect. */
function effectSpec(choice: ShapeEffectChoice, current: Element | undefined): NodeSpec | null {
  switch (choice.kind) {
    case 'shadow': {
      if (choice.preset === null) return null;
      const s = Object.prototype.hasOwnProperty.call(SHADOW_PRESETS, choice.preset)
        ? SHADOW_PRESETS[choice.preset]
        : undefined;
      if (!s) throw new CommandRefusal('errors.unsupported', 'shadow');
      return specFromXml(shadowXml(s));
    }
    case 'reflection': {
      if (choice.preset === null) return null;
      if (!Object.prototype.hasOwnProperty.call(REFLECTION_PRESETS, choice.preset))
        throw new CommandRefusal('errors.unsupported', 'reflection');
      return specFromXml(reflectionXml(choice.preset));
    }
    case 'glow': {
      if (choice.size === null) return null;
      // A color change sends the size the chrome shows (rounded to 0.01 pt): keep the rad.
      const oldRad = intAttr(current, 'rad');
      const rad =
        oldRad !== undefined && Math.abs(oldRad / EMU_PER_POINT - choice.size) < 0.006
          ? oldRad
          : radius(choice.size);
      // Keep the color of the glow the shape has, as Glow Size does.
      const old = current && elements(current)[0];
      const color = choice.color
        ? glowColorSpec(choice.color)
        : old
          ? specOf(old)
          : glowColorSpec(DEFAULT_GLOW_COLOR);
      return {
        ns: NS.a,
        local: 'glow',
        prefix: 'a',
        attrs: [{ ns: '', local: 'rad', value: String(rad) }],
        children: [color],
      };
    }
    case 'softEdge':
      if (choice.size === null) return null;
      return specFromXml(`<a:softEdge rad="${radius(choice.size)}"/>`);
    case 'blur':
      if (choice.size === null) return null;
      return specFromXml(`<a:blur rad="${radius(choice.size)}"/>`);
    case 'bevel':
    case 'rotation3d':
      // `setThreeD` writes them.
      throw new CommandRefusal('errors.unsupported', 'effect');
  }
}

/** The `a:` children of `parent` named in `names`. */
const named = (parent: Element, names: ReadonlySet<string>): Element[] =>
  elements(parent).filter((c) => c.ns === NS.a && names.has(c.local));

/**
 * Copy a spec, replacing each `a:schemeClr val="phClr"` with the style reference's own color:
 * inside a theme style, the placeholder color is the color the `a:effectRef` names
 * (ECMA-376 Part 1, 20.1.4.2.8 effectRef). The placeholder's transforms follow the ref's own.
 */
export function withRefColor(spec: NodeSpec, ref: NodeSpec | undefined): NodeSpec {
  const isPh =
    spec.ns === NS.a &&
    spec.local === 'schemeClr' &&
    spec.attrs?.some((a) => a.local === 'val' && a.value === 'phClr');
  if (isPh && ref) return { ...ref, children: [...(ref.children ?? []), ...(spec.children ?? [])] };
  // A reference without a color: the placeholder draws black, as the cascade resolves it, so
  // the shape keeps the look it had and no `phClr` reaches its own properties.
  if (isPh)
    return {
      ns: NS.a,
      local: 'srgbClr',
      prefix: 'a',
      attrs: [{ ns: '', local: 'val', value: '000000' }],
      children: spec.children ?? [],
    };
  if (!spec.children) return spec;
  return {
    ...spec,
    children: spec.children.map((c) => ('text' in c ? c : withRefColor(c, ref))),
  };
}

/**
 * The effects a theme style gives (`p:style/a:effectRef` into the theme's `a:effectStyleLst`),
 * as specs with the placeholder color resolved. Empty without a style and for index 0. An
 * effect style that uses `a:effectDag` is refused, as an inherited one is: an own effect list
 * would replace its effects, and the editor does not write a DAG.
 */
function themeEffects(batch: OpBatch, part: string, style: Element | undefined): NodeSpec[] {
  const ref = child(style, 'a:effectRef');
  const idx = intAttr(ref, 'idx') ?? 0;
  if (idx < 1) return [];
  const effectStyle = themeEffectStyle(batch, part, idx);
  if (child(effectStyle, 'a:effectDag'))
    throw new CommandRefusal('errors.unsupported', 'effectDag');
  const list = child(effectStyle, 'a:effectLst');
  if (!list) return [];
  const color = ref ? elements(ref)[0] : undefined;
  const refColor = color ? specOf(color) : undefined;
  return elements(list).map((e) => withRefColor(specOf(e), refColor));
}

/**
 * The effects a shape has before its own `a:effectLst`, as `resolveSlide` reads them: the
 * nearest effect list of the placeholders it inherits from, else the theme style of the nearest
 * `p:style` (its own first). A shape that edits one effect keeps the others.
 */
export function inheritedEffects(batch: OpBatch, part: string, el: Element): NodeSpec[] {
  const chain = [el, ...inheritedPlaceholders(batch.pkg, part, el)];
  for (const e of chain.slice(1)) {
    const spPr = child(e, 'p:spPr');
    if (child(spPr, 'a:effectDag')) throw new CommandRefusal('errors.unsupported', 'effectDag');
    const list = child(spPr, 'a:effectLst');
    if (list) return elements(list).map(specOf);
  }
  const style = chain.map((e) => child(e, 'p:style')).find((s) => s !== undefined);
  return themeEffects(batch, part, style);
}

/** The shapes, connectors, and pictures among `ids`: the elements that carry effects. */
function effectTargets(batch: OpBatch, part: string, ids: readonly string[]): Element[] {
  if (ids.length === 0) throw new CommandRefusal('disabled.noSelection');
  const out = ids
    .map((id) => batch.element(part, id))
    .filter((el) => el.kind === 'sp' || el.kind === 'cxnSp' || el.kind === 'pic');
  if (out.length === 0) throw new CommandRefusal('disabled.noShapeEffects');
  return out;
}

/** The shape's `a:effectLst`, written from its inherited effects when the shape has none. */
export function effectListOf(
  batch: OpBatch,
  part: string,
  el: Element,
  inherited: readonly NodeSpec[]
): Element {
  const spPr = child(el, 'p:spPr');
  if (!spPr) throw new CommandRefusal('errors.malformed');
  if (child(spPr, 'a:effectDag')) throw new CommandRefusal('errors.unsupported', 'effectDag');
  const list = child(spPr, 'a:effectLst');
  if (list) return list;
  let index = 0;
  spPr.children.forEach((c, i) => {
    if (c.type === 'element' && c.ns === NS.a && BEFORE_EFFECTS.has(c.local)) index = i + 1;
  });
  const id = batch.insert({
    op: 'insertNode',
    part,
    parent: spPr.id,
    index,
    spec: {
      ns: NS.a,
      local: 'effectLst',
      prefix: 'a',
      children: [...inherited],
    },
  });
  return batch.element(part, id);
}

/**
 * Make an existing element equal to `spec` in place: its attributes, then its children by
 * position, recursing where the names match. Two replicas that change the same element then
 * set attributes (one wins) instead of each inserting a copy, which would leave two `p:style`
 * or two effects of one kind in a shape.
 */
function morph(batch: OpBatch, part: string, el: Element, spec: NodeSpec): void {
  const want = new Map((spec.attrs ?? []).map((a) => [`${a.ns}|${a.local}`, a]));
  for (const a of el.attrs)
    if (!want.has(`${a.ns}|${a.local}`))
      batch.push({
        op: 'setAttribute',
        part,
        node: el.id,
        name: { ns: a.ns, local: a.local },
        value: null,
      });
  for (const a of spec.attrs ?? []) {
    const have = el.attrs.find((x) => x.ns === a.ns && x.local === a.local)?.value;
    if (have !== a.value)
      batch.push({
        op: 'setAttribute',
        part,
        node: el.id,
        name: { ns: a.ns, local: a.local, ...(a.prefix ? { prefix: a.prefix } : {}) },
        value: a.value,
      });
  }
  const kids = (spec.children ?? []).filter((c): c is NodeSpec => !('text' in c));
  const old = elements(el);
  kids.forEach((k, i) => {
    const o = old[i];
    if (o && o.ns === k.ns && o.local === k.local) morph(batch, part, o, k);
    else {
      if (o) batch.push({ op: 'removeNode', part, node: o.id });
      batch.insert({ op: 'insertNode', part, parent: el.id, index: i, spec: k });
    }
  });
  for (const o of old.slice(kids.length)) batch.push({ op: 'removeNode', part, node: o.id });
}

/**
 * The 3-D properties an element inherits: from the placeholders it inherits from, else from the
 * theme effect style of its nearest `p:style` (CT_EffectStyleItem holds `a:scene3d` and
 * `a:sp3d`), as `resolveSlide` reads them.
 */
function inheritedThreeD(batch: OpBatch, part: string, el: Element): InheritedThreeD {
  const chain = [el, ...inheritedPlaceholders(batch.pkg, part, el)];
  const spPrs = chain.slice(1).map((e) => child(e, 'p:spPr'));
  // Only a turned camera has a rotation to remove.
  const turned = (scene: Element | undefined): boolean => {
    const camera = child(scene, 'a:camera');
    const rot = child(camera, 'a:rot');
    const angles = ['lat', 'lon', 'rev'].map((n) => intAttr(rot, n) ?? 0);
    return (
      scene !== undefined &&
      (!/^(orthographicFront|perspectiveFront)$/.test(attr(camera, 'prst') ?? '') ||
        angles.some((a) => a % 21_600_000 !== 0))
    );
  };
  let scene = spPrs.some((s) => turned(child(s, 'a:scene3d')));
  let anyScene = spPrs.some((s) => child(s, 'a:scene3d') !== undefined);
  let bevel = spPrs.some((s) => child(child(s, 'a:sp3d'), 'a:bevelT') !== undefined);
  const style = chain.map((e) => child(e, 'p:style')).find((s) => s !== undefined);
  const idx = intAttr(child(style, 'a:effectRef'), 'idx') ?? 0;
  if (idx > 0) {
    const effectStyle = themeEffectStyle(batch, part, idx);
    anyScene ||= child(effectStyle, 'a:scene3d') !== undefined;
    scene ||= turned(child(effectStyle, 'a:scene3d'));
    bevel ||= child(child(effectStyle, 'a:sp3d'), 'a:bevelT') !== undefined;
  }
  return { scene, anyScene, bevel };
}

/**
 * The theme part of a slide, layout, or master part: through the slide's layout, the layout's
 * master, and the master's theme relationship.
 */
export function themePartFor(batch: OpBatch, part: string): string | undefined {
  const layout = relTargetByType(batch.pkg, part, REL.slideLayout) ?? part;
  const master = relTargetByType(batch.pkg, layout, REL.slideMaster) ?? layout;
  return relTargetByType(batch.pkg, master, REL.theme);
}

/** Effect style `idx` (1-based) of the theme that `part` uses, or undefined. */
function themeEffectStyle(batch: OpBatch, part: string, idx: number): Element | undefined {
  const themePart = themePartFor(batch, part);
  const theme = themePart ? parseTheme(xmlPartOf(batch.pkg, themePart)) : undefined;
  return theme?.effectStyles[idx - 1];
}

/** Set or remove one effect on the selected shapes, connectors, and pictures. */
export function setShapeEffect(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  choice: ShapeEffectChoice
): void {
  const kind = choice?.kind;
  const owned = Object.prototype.hasOwnProperty.call(OWNED, kind) ? OWNED[kind] : undefined;
  if (!owned) throw new CommandRefusal('errors.unsupported', 'effect');
  if (isThreeDChoice(choice)) {
    setThreeD(batch, part, effectTargets(batch, part, ids), choice, (el) =>
      inheritedThreeD(batch, part, el)
    );
    return;
  }
  let changed = 0;
  for (const el of effectTargets(batch, part, ids)) {
    const spPr = child(el, 'p:spPr');
    // An own effect DAG is refused before the "nothing to remove" check, so the reason is true.
    if (child(spPr, 'a:effectDag')) throw new CommandRefusal('errors.unsupported', 'effectDag');
    // Removing an effect the shape does not have, neither its own nor an inherited one,
    // changes nothing: no empty effect list and no undo step.
    const removing =
      (choice.kind === 'shadow' || choice.kind === 'reflection' ? choice.preset : choice.size) ===
      null;
    const own = child(spPr, 'a:effectLst');
    const inherited = own ? undefined : inheritedEffects(batch, part, el);
    if (
      removing &&
      (own ? named(own, owned).length === 0 : !inherited!.some((e) => owned.has(e.local)))
    )
      continue;
    changed += 1;
    const list = own ?? effectListOf(batch, part, el, inherited!);
    const old = named(list, owned);
    const spec = effectSpec(choice, choice.kind === 'glow' ? old[0] : undefined);
    // The same kind of element: change it in place (see `morph`).
    if (spec && old.length === 1 && old[0]!.local === spec.local) {
      morph(batch, part, old[0]!, spec);
      continue;
    }
    for (const c of old) batch.push({ op: 'removeNode', part, node: c.id });
    if (!spec) continue;
    const rank = EFFECT_ORDER.indexOf(spec.local);
    const current = batch.element(part, list.id);
    let index = 0;
    current.children.forEach((c, i) => {
      if (c.type === 'element' && EFFECT_ORDER.indexOf(c.local) < rank) index = i + 1;
    });
    batch.insert({ op: 'insertNode', part, parent: list.id, index, spec });
  }
  if (changed === 0) throw new CommandRefusal('disabled.noEffectChange');
}

/** One style-matrix reference: the theme index and its color (`C` is the column color). */
type StyleRef = readonly [idx: number, color: 'C' | 'Cshade' | 'lt1' | 'dk1'];

/**
 * The Shape Styles gallery rows as style-matrix references (`a:lnRef`, `a:fillRef`,
 * `a:effectRef`, and the `a:fontRef` color), from the default theme styles.
 */
export const SHAPE_STYLE_ROWS: Readonly<
  Record<
    ShapeStyleRow,
    {
      readonly ln: StyleRef;
      readonly fill: StyleRef;
      readonly effect: number;
      readonly font: 'lt1' | 'dk1';
    }
  >
> = {
  coloredOutline: { ln: [2, 'C'], fill: [1, 'lt1'], effect: 0, font: 'dk1' },
  coloredFill: { ln: [2, 'Cshade'], fill: [1, 'C'], effect: 0, font: 'lt1' },
  lightOutline: { ln: [3, 'lt1'], fill: [1, 'C'], effect: 1, font: 'lt1' },
  subtleEffect: { ln: [1, 'C'], fill: [2, 'C'], effect: 1, font: 'dk1' },
  moderateEffect: { ln: [1, 'C'], fill: [3, 'C'], effect: 2, font: 'lt1' },
  intenseEffect: { ln: [0, 'C'], fill: [3, 'C'], effect: 3, font: 'lt1' },
};
export const SHAPE_STYLE_COLORS: readonly ShapeStyleColor[] = [
  'dk1',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
];

function refColorXml(ref: StyleRef[1], color: ShapeStyleColor): string {
  if (ref === 'Cshade') return `<a:schemeClr val="${color}"><a:shade val="15000"/></a:schemeClr>`;
  return `<a:schemeClr val="${ref === 'C' ? color : ref}"/>`;
}

/** The `p:style` of a gallery cell. */
export function shapeStyleXml(style: ShapeStyleChoice): string {
  const row = Object.prototype.hasOwnProperty.call(SHAPE_STYLE_ROWS, style?.row)
    ? SHAPE_STYLE_ROWS[style.row]
    : undefined;
  if (!row || !SHAPE_STYLE_COLORS.includes(style.color))
    throw new CommandRefusal('errors.unsupported', 'shapeStyle');
  const c = style.color;
  return (
    `<p:style><a:lnRef idx="${row.ln[0]}">${refColorXml(row.ln[1], c)}</a:lnRef>` +
    `<a:fillRef idx="${row.fill[0]}">${refColorXml(row.fill[1], c)}</a:fillRef>` +
    `<a:effectRef idx="${row.effect}">${refColorXml('C', c)}</a:effectRef>` +
    `<a:fontRef idx="minor">${refColorXml(row.font, c)}</a:fontRef></p:style>`
  );
}

const DIRECT_STYLE = new Set([
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'ln',
  'effectLst',
  'effectDag',
]);

/**
 * Apply a Shape Styles gallery cell to the selected shapes: write `p:style` (after `p:spPr`,
 * CT_Shape) and remove the shape's own fill, outline, and effects, so the style shows.
 */
export function setShapeStyle(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  style: ShapeStyleChoice
): void {
  if (ids.length === 0) throw new CommandRefusal('disabled.noSelection');
  const xml = shapeStyleXml(style);
  const shapes = ids.map((id) => batch.element(part, id)).filter((el) => el.kind === 'sp');
  if (shapes.length === 0) throw new CommandRefusal('disabled.noShapeStyle');
  for (const el of shapes) {
    const spPr = child(el, 'p:spPr');
    if (!spPr) throw new CommandRefusal('errors.malformed');
    for (const c of named(spPr, DIRECT_STYLE)) batch.push({ op: 'removeNode', part, node: c.id });
    const old = child(batch.element(part, el.id), 'p:style');
    if (old) {
      morph(batch, part, old, specFromXml(xml));
      continue;
    }
    const current = batch.element(part, el.id);
    const index = current.children.findIndex((c) => c.id === spPr.id) + 1;
    batch.insert({ op: 'insertNode', part, parent: el.id, index, spec: specFromXml(xml) });
  }
}

/** The gallery cell a shape's `p:style` matches, or undefined. Transforms are not compared. */
export function shapeStyleOf(el: Element | undefined): ShapeStyleChoice | undefined {
  const style = child(el, 'p:style');
  if (!style) return undefined;
  const ref = (name: string): { idx: number; val: string | undefined } => {
    const r = child(style, name);
    return { idx: intAttr(r, 'idx') ?? -1, val: attr(r ? elements(r)[0] : undefined, 'val') };
  };
  const ln = ref('a:lnRef');
  const fill = ref('a:fillRef');
  const effect = ref('a:effectRef');
  const font = attr(elements(child(style, 'a:fontRef'))[0], 'val');
  const want = (r: StyleRef[1], c: ShapeStyleColor): string =>
    r === 'C' || r === 'Cshade' ? c : r;
  for (const [row, s] of Object.entries(SHAPE_STYLE_ROWS) as [
    ShapeStyleRow,
    (typeof SHAPE_STYLE_ROWS)[ShapeStyleRow],
  ][])
    for (const color of SHAPE_STYLE_COLORS)
      if (
        ln.idx === s.ln[0] &&
        ln.val === want(s.ln[1], color) &&
        fill.idx === s.fill[0] &&
        fill.val === want(s.fill[1], color) &&
        effect.idx === s.effect &&
        font === s.font
      )
        return { row, color };
  return undefined;
}

const near = (a: number, b: number, tolerance: number): boolean => Math.abs(a - b) <= tolerance;
const points = (emu: number): number => Math.round((emu / EMU_PER_POINT) * 100) / 100;

/** The shadow preset resolved effects match, `none`, or `other`. */
function shadowState(fx: ResolvedEffects | undefined): ShapeEffectsState['shadow'] {
  const shadow = fx?.outerShadow ?? fx?.innerShadow;
  if (!fx || !shadow) return 'none';
  // An outer and an inner shadow together are no one gallery choice; a pick replaces both.
  if (fx.outerShadow && fx.innerShadow) return 'other';
  const isInner = !fx.outerShadow;
  // Every preset is black: a colored shadow is not a gallery choice.
  const c = shadow.color;
  if (c.r > 1 || c.g > 1 || c.b > 1) return 'other';
  const o = fx.outerShadow;
  // Outer presets do not skew and scale evenly. `rotWithShape` is not compared: the gallery
  // writes `rotWithShape="0"`, and other writers write none.
  if (o && (o.kx !== 0 || o.ky !== 0 || !near(o.sx, o.sy, 0.005))) return 'other';
  for (const [id, s] of Object.entries(SHADOW_PRESETS) as [ShadowPreset, ShadowSpec][]) {
    if (s.inner !== isInner) continue;
    if (!near(shadow.blur, s.blur, 100) || !near(shadow.dist, s.dist, 100)) continue;
    if (s.dist && !near(shadow.dir, s.dir / 60_000, 0.5)) continue;
    if (!near(c.a, (s.alpha ?? 100_000) / 100_000, 0.02)) continue;
    if (o && (o.align !== (s.algn ?? 'b') || !near(o.sx, (s.scale ?? 100_000) / 100_000, 0.005)))
      continue;
    return id;
  }
  return 'other';
}

function reflectionState(fx: ResolvedEffects | undefined): ShapeEffectsState['reflection'] {
  const r = fx?.reflection;
  if (!r) return 'none';
  // The fixed part of every preset (reflectionXml): blur, fade, direction, and mirror.
  const fixed =
    near(r.blur, 6350, 100) &&
    near(r.startAlpha, 0.5, 0.01) &&
    near(r.endAlpha, 0.003, 0.002) &&
    near(r.startPos, 0, 0.005) &&
    near(r.dir, 90, 0.5) &&
    near(r.sy, -1, 0.005) &&
    near(r.sx, 1, 0.005) &&
    r.align === 'bl' &&
    !r.rotWithShape;
  if (!fixed) return 'other';
  for (const [id, p] of Object.entries(REFLECTION_PRESETS) as [
    ReflectionPreset,
    { endPos: number; dist: number },
  ][])
    if (near(r.endPos, p.endPos / 100_000, 0.005) && near(r.dist, p.dist, 100)) return id;
  return 'other';
}

/**
 * The theme accent of an element's own glow when it is a Glow Variations color: `a:schemeClr`
 * `accent1` to `accent6` at the 40% opacity `glowColorSpec` writes.
 */
function glowAccentOf(el: Element | undefined): ShapeEffectsState['glowAccent'] {
  const glow = child(child(child(el, 'p:spPr'), 'a:effectLst'), 'a:glow');
  const color = glow ? elements(glow)[0] : undefined;
  if (color?.local !== 'schemeClr') return undefined;
  const val = attr(color, 'val');
  const alpha = intAttr(child(color, 'a:alpha'), 'val');
  // Only the transforms `glowColorSpec` writes: any other changes the color.
  const extra = elements(color).some(
    (t) => t.local !== 'alpha' && !(t.local === 'satMod' && attr(t, 'val') === '175000')
  );
  return !extra && alpha === 40_000 && /^accent[1-6]$/.test(val ?? '')
    ? (val as ShapeEffectsState['glowAccent'])
    : undefined;
}

/** The effects of an element as the chrome shows them; `el` is its XML, for the glow accent. */
export function effectsStateOf(fx: ResolvedEffects | undefined, el?: Element): ShapeEffectsState {
  const glow = fx?.glow;
  const accent = glow ? glowAccentOf(el) : undefined;
  const hex = glow
    ? `#${[glow.color.r, glow.color.g, glow.color.b]
        .map((v) =>
          Math.round(Math.max(0, Math.min(255, v)))
            .toString(16)
            .padStart(2, '0')
        )
        .join('')
        .toUpperCase()}`
    : undefined;
  return {
    shadow: shadowState(fx),
    reflection: reflectionState(fx),
    glow: glow ? points(glow.radius) : 0,
    ...(hex ? { glowColor: hex } : {}),
    ...(accent ? { glowAccent: accent } : {}),
    softEdge: fx?.softEdge ? points(fx.softEdge.radius) : 0,
    bevel: bevelStateOf(fx),
    rotation3d: rotationStateOf(fx),
    blur: fx?.blur ? points(fx.blur.radius) : 0,
  };
}
