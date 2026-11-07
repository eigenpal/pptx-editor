// The background of a slide, a layout, or a master (`p:cSld/p:bg`, ECMA-376 Part 1, 19.3.1.1):
// read the fill the part shows, and write an own `p:bgPr` (CT_BackgroundProperties), the
// inheritance flag, and `showMasterSp`.
//
// A slide reads its background from the cascade (`resolveSlide`). A layout or a master reads it
// with the same rule over its own chain: the first `p:bg` of the part, then of its master, wins.

import { PATTERN_PRESETS } from './commands/fill-commands.ts';
import {
  DEFAULT_COLOR_MAP,
  fillElementOf,
  parseTheme,
  readColorMap,
  resolveFillElement,
  styleRefContext,
  themeFill,
  toHex,
  type ColorContext,
  type ResolvedFill,
  type Rgba,
} from '../cascade/index.ts';
import {
  attr,
  child,
  elements,
  intAttr,
  NS,
  path,
  REL,
  relationshipsOf,
  specOf,
  xmlPartOf,
  type Element,
  type NodeSpec,
} from '../store/index.ts';
import { addRelationship } from './commands/package-edits.ts';
import { specFromXml } from './commands/shape-commands.ts';
import type { AutomationBackgroundWrite } from './operations.ts';
import { resolvedSlide } from './plan-reads.ts';
import { imageRelationship } from './plan-shapes.ts';
import { colorSpec } from './plan-text.ts';
import type { AutomationBackgroundFillRead, AutomationRef, AutomationValue } from './protocol.ts';
import {
  AutomationFailure,
  badArgument,
  edit,
  masterOfLayout,
  numberIn,
  ownerPart,
  slideOf,
  targetOf,
  type PlanContext,
} from './targets.ts';

/** A slide, a layout, or a master, with the parts its background comes from. */
export interface BackgroundOwner {
  readonly kind: 'slide' | 'slideLayout' | 'slideMaster';
  readonly part: string;
  /** The own part first, then the layout and the master it inherits from. */
  readonly chain: readonly string[];
  readonly themePart: string | undefined;
  /** The slide index, for the cascade. */
  readonly slideIndex?: number;
}

/** The theme part of a master. */
function themeOfMaster(ctx: PlanContext, master: string | undefined): string | undefined {
  if (!master) return undefined;
  return relationshipsOf(ctx.batch.pkg, master).find((r) => r.type === REL.theme && !r.external)
    ?.target;
}

/** The owner a ref names: a slide, a layout, or a master. */
export function backgroundOwner(ctx: PlanContext, ref: AutomationRef): BackgroundOwner {
  const target = targetOf(ctx, ref, 'slide', 'slideLayout', 'slideMaster');
  const pkg = ctx.batch.pkg;
  if (target.kind === 'slide') {
    const { ref: slide, index } = slideOf(pkg, target.sldId);
    const chain = [slide.partName, slide.layoutPart, slide.masterPart];
    return {
      kind: 'slide',
      part: slide.partName,
      chain: chain.filter((p): p is string => p !== undefined),
      themePart: slide.themePart,
      slideIndex: index,
    };
  }
  const part = ownerPart(pkg, { kind: target.kind, part: target.part });
  const master = target.kind === 'slideLayout' ? masterOfLayout(pkg, part) : part;
  const chain = target.kind === 'slideLayout' && master ? [part, master] : [part];
  return { kind: target.kind, part, chain, themePart: themeOfMaster(ctx, master) };
}

/** The `p:bg` element of a part, if it has its own. */
function ownBackground(ctx: PlanContext, part: string): Element | undefined {
  return path(xmlPartOf(ctx.batch.pkg, part)?.root, 'p:cSld', 'p:bg');
}

/** The color context of a part: the master's `p:clrMap` with each override below it. */
function colorContext(ctx: PlanContext, owner: BackgroundOwner): ColorContext {
  const theme = parseTheme(owner.themePart ? xmlPartOf(ctx.batch.pkg, owner.themePart) : undefined);
  const roots = owner.chain.map((p) => xmlPartOf(ctx.batch.pkg, p)?.root);
  const master = roots[roots.length - 1];
  let colorMap = readColorMap(child(master, 'p:clrMap')) ?? DEFAULT_COLOR_MAP;
  for (const over of roots.slice(0, -1).reverse()) {
    const read = readColorMap(path(over, 'p:clrMapOvr', 'a:overrideClrMapping'));
    if (read) colorMap = read;
  }
  return { scheme: theme.colors, colorMap };
}

/** The background a part shows: where it comes from, its fill element, and the resolved fill. */
interface ShownBackground {
  /** The part of the `p:bg` that wins, or undefined when no part has one. */
  readonly part: string | undefined;
  readonly bg: Element | undefined;
  /** The fill element: in `p:bgPr`, or the theme's background fill style that `p:bgRef` names. */
  readonly fillElement: Element | undefined;
  readonly fill: ResolvedFill;
}

const WHITE: ResolvedFill = { type: 'solid', color: { r: 255, g: 255, b: 255, a: 1 } };

/** The background a part shows, by the rule of `resolveSlide` (cascade/slide.ts). */
function shownBackground(ctx: PlanContext, owner: BackgroundOwner): ShownBackground {
  const colors = colorContext(ctx, owner);
  const theme = parseTheme(owner.themePart ? xmlPartOf(ctx.batch.pkg, owner.themePart) : undefined);
  for (const part of owner.chain) {
    const bg = ownBackground(ctx, part);
    if (!bg) continue;
    const bgPr = child(bg, 'p:bgPr');
    if (bgPr) {
      const fillElement = fillElementOf(bgPr);
      const fill = resolveFillElement(fillElement, colors);
      if (fill) return { part, bg, fillElement, fill };
    }
    const bgRef = child(bg, 'p:bgRef');
    if (bgRef) {
      const fillElement = themeFill(intAttr(bgRef, 'idx') ?? 0, theme);
      const fill = fillElement
        ? resolveFillElement(fillElement, styleRefContext(bgRef, colors))
        : undefined;
      if (fill) return { part, bg, fillElement, fill };
    }
  }
  return { part: undefined, bg: undefined, fillElement: undefined, fill: WHITE };
}

/** The shown background, with the fill a slide takes from the cascade. */
function shown(ctx: PlanContext, owner: BackgroundOwner): ShownBackground {
  const found = shownBackground(ctx, owner);
  if (owner.slideIndex === undefined) return found;
  const resolved = resolvedSlide(ctx.batch.pkg, owner.slideIndex);
  return resolved ? { ...found, fill: resolved.background } : found;
}

const hex = (c: Rgba): string => toHex({ ...c, a: 1 }).toUpperCase();
const transparencyOf = (a: number): number => Math.round((1 - a) * 1000) / 1000;

const NO_FILL: AutomationBackgroundFillRead = {
  type: 'unsupported',
  color: null,
  transparency: null,
  gradientType: null,
  pattern: null,
  foregroundColor: null,
  backgroundColor: null,
};

/** The object model's view of a shown background fill. */
function fillRead(found: ShownBackground): AutomationBackgroundFillRead {
  const fill = found.fill;
  switch (fill.type) {
    case 'solid':
      return {
        ...NO_FILL,
        type: 'solid',
        color: hex(fill.color),
        transparency: transparencyOf(fill.color.a),
      };
    case 'linear':
    case 'path': {
      const title = attr(child(found.bg, 'p:bgPr'), 'shadeToTitle');
      return {
        ...NO_FILL,
        type: 'gradient',
        gradientType:
          title === '1' || title === 'true'
            ? 'shadeToTitle'
            : fill.type === 'linear'
              ? 'linear'
              : fill.path,
      };
    }
    case 'pattern':
      return {
        ...NO_FILL,
        type: 'pattern',
        pattern: fill.preset,
        foregroundColor: hex(fill.fg),
        backgroundColor: hex(fill.bg),
      };
    case 'picture': {
      const amount = fill.effects.find((e) => e.kind === 'alphaModFix');
      const alpha = amount?.kind === 'alphaModFix' ? amount.amount : 1;
      return { ...NO_FILL, type: 'picture', transparency: transparencyOf(alpha) };
    }
    default:
      return NO_FILL;
  }
}

export function readBackground(ctx: PlanContext, ref: AutomationRef): AutomationValue {
  const owner = backgroundOwner(ctx, ref);
  const root = xmlPartOf(ctx.batch.pkg, owner.part)?.root;
  const master = owner.kind === 'slideMaster';
  return {
    kind: 'background',
    background: {
      followsMaster: !master && !ownBackground(ctx, owner.part),
      graphicsHidden: !master && attr(root, 'showMasterSp') === '0',
      fill: fillRead(shown(ctx, owner)),
    },
  };
}

/** The gradient shadings a write accepts. */
const GRADIENTS: ReadonlySet<string> = new Set(['linear', 'circle', 'rect', 'shape']);

/** The own `p:bgPr` fill element of a part, when it is a `local` element. */
function ownFill(ctx: PlanContext, part: string, local: string): Element | undefined {
  const el = fillElementOf(child(ownBackground(ctx, part), 'p:bgPr'));
  return el?.local === local ? el : undefined;
}

/** A spec of a color element with its `a:alpha` set to `transparency` (0 removes it). */
function withAlpha(color: Element, transparency: number): NodeSpec {
  const spec = specOf(color);
  const kept = (spec.children ?? []).filter(
    (c) => !('local' in c) || c.ns !== NS.a || c.local !== 'alpha'
  );
  const alpha: NodeSpec = {
    ns: NS.a,
    local: 'alpha',
    prefix: 'a',
    attrs: [{ ns: '', local: 'val', value: String(Math.round((1 - transparency) * 100_000)) }],
  };
  return { ...spec, children: transparency === 0 ? kept : [...kept, alpha] };
}

/** The first color child of a fill element. */
function colorElement(el: Element | undefined): Element | undefined {
  return elements(el).find((c) => c.ns === NS.a && c.local.endsWith('Clr'));
}

/** Insert `p:bg` with `p:bgPr` over `fill` as the first child of `p:cSld`. */
function insertBackground(ctx: PlanContext, part: string, fill: NodeSpec): void {
  const cSld = path(xmlPartOf(ctx.batch.pkg, part)?.root, 'p:cSld');
  if (!cSld) throw new AutomationFailure('unsupported-content', 'the part has no p:cSld');
  const effects = specFromXml('<a:effectLst/>');
  ctx.batch.insert({
    op: 'insertNode',
    part,
    parent: cSld.id,
    index: 0,
    spec: {
      ns: NS.p,
      local: 'bg',
      prefix: 'p',
      children: [{ ns: NS.p, local: 'bgPr', prefix: 'p', children: [fill, effects] }],
    },
  });
}

/**
 * Give a part `fill` as its own background fill. An own `p:bgPr` keeps its effects and
 * extensions and loses `shadeToTitle`; a `p:bgRef`, or no own `p:bg`, becomes a new `p:bgPr`.
 */
function writeFill(ctx: PlanContext, part: string, fill: NodeSpec): void {
  const bgPr = child(ownBackground(ctx, part), 'p:bgPr');
  if (!bgPr) {
    const old = ownBackground(ctx, part);
    if (old) ctx.batch.push({ op: 'removeNode', part, node: old.id });
    insertBackground(ctx, part, fill);
    return;
  }
  const old = fillElementOf(bgPr);
  if (old) ctx.batch.push({ op: 'removeNode', part, node: old.id });
  if (attr(bgPr, 'shadeToTitle') !== undefined)
    ctx.batch.push({
      op: 'setAttribute',
      part,
      node: bgPr.id,
      name: { ns: '', local: 'shadeToTitle' },
      value: null,
    });
  // EG_FillProperties is the first group of CT_BackgroundProperties.
  ctx.batch.insert({ op: 'insertNode', part, parent: bgPr.id, index: 0, spec: fill });
}

/** A transparency from 0 to 1, or the current one. */
function transparencyArg(value: unknown, current: number | null): number {
  return value === undefined ? (current ?? 0) : numberIn(value, 'transparency', 0, 1);
}

function setSolid(
  ctx: PlanContext,
  owner: BackgroundOwner,
  write: Extract<AutomationBackgroundWrite, { kind: 'solid' }>
): void {
  const now = fillRead(shown(ctx, owner));
  const transparency = transparencyArg(
    write.transparency,
    now.type === 'solid' ? now.transparency : 0
  );
  if (write.color !== undefined) {
    const xml = `<a:solidFill>${colorSpec(write.color, 'color', transparency)}</a:solidFill>`;
    return writeFill(ctx, owner.part, specFromXml(xml));
  }
  // An own solid color keeps its theme color and transforms; only its alpha changes.
  const own = colorElement(ownFill(ctx, owner.part, 'solidFill'));
  if (own) {
    const color = withAlpha(own, transparency);
    return writeFill(ctx, owner.part, {
      ns: NS.a,
      local: 'solidFill',
      prefix: 'a',
      children: [color],
    });
  }
  const color = now.type === 'solid' && now.color ? now.color : '#FFFFFF';
  writeFill(
    ctx,
    owner.part,
    specFromXml(`<a:solidFill>${colorSpec(color, 'color', transparency)}</a:solidFill>`)
  );
}

/** The shading element of a gradient: `a:lin`, or `a:path` with a centered `a:fillToRect`. */
function shadeXml(kind: string, angle: number): string {
  if (kind === 'linear') return `<a:lin ang="${angle}" scaled="0"/>`;
  return `<a:path path="${kind}"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>`;
}

function setGradient(
  ctx: PlanContext,
  owner: BackgroundOwner,
  write: Extract<AutomationBackgroundWrite, { kind: 'gradient' }>
): void {
  const found = shown(ctx, owner);
  const fill = found.fill;
  const current = fill.type === 'linear' ? 'linear' : fill.type === 'path' ? fill.path : undefined;
  const kind = write.gradientType ?? current ?? 'linear';
  if (typeof kind !== 'string' || !GRADIENTS.has(kind)) badArgument('gradientType');
  // `a:lin/@ang` in 60 000ths of a degree (ST_PositiveFixedAngle); 90° runs top to bottom.
  const turn = 21_600_000;
  const angle =
    fill.type === 'linear' ? ((Math.round(fill.angle * 60_000) % turn) + turn) % turn : 5_400_000;
  const own = ownFill(ctx, owner.part, 'gradFill');
  const gsLst = child(own, 'a:gsLst');
  let stops: NodeSpec;
  if (gsLst) stops = specOf(gsLst);
  else if ((fill.type === 'linear' || fill.type === 'path') && fill.stops.length > 0) {
    const gs = fill.stops
      .map(
        (s) =>
          `<a:gs pos="${Math.round(s.pos * 100_000)}">${colorSpec(hex(s.color), 'color', transparencyOf(s.color.a))}</a:gs>`
      )
      .join('');
    stops = specFromXml(`<a:gsLst>${gs}</a:gsLst>`);
  } else
    stops = specFromXml(
      '<a:gsLst><a:gs pos="0"><a:schemeClr val="bg1"/></a:gs><a:gs pos="100000"><a:schemeClr val="bg2"/></a:gs></a:gsLst>'
    );
  // CT_GradientFillProperties: gsLst, then the shading, then tileRect.
  const tileRect = child(own, 'a:tileRect');
  const children = [
    stops,
    specFromXml(shadeXml(kind, angle)),
    ...(tileRect ? [specOf(tileRect)] : []),
  ];
  writeFill(ctx, owner.part, {
    ns: NS.a,
    local: 'gradFill',
    prefix: 'a',
    attrs: [{ ns: '', local: 'rotWithShape', value: '1' }],
    children,
  });
}

function setPattern(
  ctx: PlanContext,
  owner: BackgroundOwner,
  write: Extract<AutomationBackgroundWrite, { kind: 'pattern' }>
): void {
  const now = fillRead(shown(ctx, owner));
  const own = ownFill(ctx, owner.part, 'pattFill');
  const pattern = write.pattern ?? now.pattern ?? 'pct5';
  if (typeof pattern !== 'string' || !PATTERN_PRESETS.has(pattern)) badArgument('pattern');
  /** The color element of one side: the written color, the own one, or the shown one. */
  const side = (
    name: 'fgClr' | 'bgClr',
    value: string | undefined,
    shownColor: string
  ): NodeSpec => {
    if (value !== undefined)
      return {
        ns: NS.a,
        local: name,
        prefix: 'a',
        children: [specFromXml(colorSpec(value, name))],
      };
    const kept = child(own, `a:${name}`);
    if (kept) return specOf(kept);
    return {
      ns: NS.a,
      local: name,
      prefix: 'a',
      children: [specFromXml(colorSpec(shownColor, name))],
    };
  };
  const isPattern = now.type === 'pattern';
  writeFill(ctx, owner.part, {
    ns: NS.a,
    local: 'pattFill',
    prefix: 'a',
    attrs: [{ ns: '', local: 'prst', value: pattern }],
    children: [
      side('fgClr', write.foregroundColor, (isPattern && now.foregroundColor) || '#000000'),
      side('bgClr', write.backgroundColor, (isPattern && now.backgroundColor) || '#FFFFFF'),
    ],
  });
}

/** An `a:alphaModFix` spec for a picture transparency, or none when opaque. */
function alphaModFix(transparency: number): NodeSpec[] {
  if (transparency === 0) return [];
  const amt = String(Math.round((1 - transparency) * 100_000));
  return [
    { ns: NS.a, local: 'alphaModFix', prefix: 'a', attrs: [{ ns: '', local: 'amt', value: amt }] },
  ];
}

function setPicture(
  ctx: PlanContext,
  owner: BackgroundOwner,
  write: Extract<AutomationBackgroundWrite, { kind: 'picture' }>
): void {
  const now = fillRead(shown(ctx, owner));
  const transparency = transparencyArg(
    write.transparency,
    now.type === 'picture' ? now.transparency : 0
  );
  if (write.base64 === undefined) {
    if (now.type !== 'picture') badArgument('imageBase64');
    // A picture the part shows from its layout or master becomes its own first.
    makeOwn(ctx, owner);
    const blip = child(ownFill(ctx, owner.part, 'blipFill'), 'a:blip');
    if (!attr(blip, 'r:embed'))
      throw new AutomationFailure('unsupported-content', 'the picture is a theme fill', 'picture');
    // Only the alpha changes: the crop, the tile or stretch mode, and other effects stay.
    for (const c of elements(blip))
      if (c.ns === NS.a && c.local === 'alphaModFix')
        ctx.batch.push({ op: 'removeNode', part: owner.part, node: c.id });
    for (const spec of alphaModFix(transparency))
      ctx.batch.insert({ op: 'insertNode', part: owner.part, parent: blip!.id, index: 0, spec });
    return;
  }
  const rId = imageRelationship(ctx, owner.part, write.base64);
  const blip: NodeSpec = {
    ns: NS.a,
    local: 'blip',
    prefix: 'a',
    attrs: [{ ns: NS.r, local: 'embed', prefix: 'r', value: rId }],
    children: alphaModFix(transparency),
  };
  writeFill(ctx, owner.part, {
    ns: NS.a,
    local: 'blipFill',
    prefix: 'a',
    attrs: [
      { ns: '', local: 'dpi', value: '0' },
      { ns: '', local: 'rotWithShape', value: '1' },
    ],
    children: [
      blip,
      specFromXml('<a:srcRect/>'),
      specFromXml('<a:stretch><a:fillRect/></a:stretch>'),
    ],
  });
}

/**
 * A copy of `spec` whose relationship ids (`r:embed`, `r:link`, `r:id`) name new relationships
 * of `part` to the same targets as those of `source`. An external target is refused: a copy
 * never adds a link to a file outside the package.
 */
function rebase(ctx: PlanContext, spec: NodeSpec, source: string, part: string): NodeSpec {
  const rels = relationshipsOf(ctx.batch.pkg, source);
  const visit = (node: NodeSpec, depth: number): NodeSpec => {
    if (depth > 64)
      throw new AutomationFailure('unsupported-content', 'the background nests too deep');
    const attrs = (node.attrs ?? []).map((a) => {
      if (a.ns !== NS.r) return a;
      const rel = rels.find((r) => r.id === a.value);
      if (!rel?.target || rel.external)
        throw new AutomationFailure(
          'unsupported-content',
          'the background links outside the package',
          'picture'
        );
      return { ...a, value: addRelationship(ctx.batch, part, rel.type, rel.target) };
    });
    const children = (node.children ?? []).map((c) => ('local' in c ? visit(c, depth + 1) : c));
    return { ...node, attrs, children };
  };
  return visit(spec, 0);
}

/** Copy the background a part shows into an own `p:bg`; nothing when it has one. */
function makeOwn(ctx: PlanContext, owner: BackgroundOwner): void {
  if (ownBackground(ctx, owner.part)) return;
  const found = shownBackground(ctx, owner);
  if (!found.bg || !found.part) {
    insertBackground(
      ctx,
      owner.part,
      specFromXml('<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>')
    );
    return;
  }
  const cSld = path(xmlPartOf(ctx.batch.pkg, owner.part)?.root, 'p:cSld');
  if (!cSld) throw new AutomationFailure('unsupported-content', 'the part has no p:cSld');
  const spec = rebase(ctx, specOf(found.bg), found.part, owner.part);
  ctx.batch.insert({ op: 'insertNode', part: owner.part, parent: cSld.id, index: 0, spec });
}

export function setBackground(
  ctx: PlanContext,
  ref: AutomationRef,
  write: AutomationBackgroundWrite
): AutomationValue {
  if (typeof write !== 'object' || write === null) badArgument('background');
  const owner = backgroundOwner(ctx, ref);
  const inherits = owner.kind !== 'slideMaster';
  edit(() => {
    switch (write.kind) {
      case 'follow': {
        if (!inherits) badArgument('background');
        const own = ownBackground(ctx, owner.part);
        if (own) ctx.batch.push({ op: 'removeNode', part: owner.part, node: own.id });
        return;
      }
      case 'own':
        if (!inherits) badArgument('background');
        return makeOwn(ctx, owner);
      case 'graphicsHidden': {
        if (!inherits || typeof write.hidden !== 'boolean')
          badArgument('areBackgroundGraphicsHidden');
        const root = xmlPartOf(ctx.batch.pkg, owner.part)?.root;
        if (!root) throw new AutomationFailure('invalid-handle', 'that part is gone', 'gone');
        // `showMasterSp` defaults to true (CT_Slide, CT_SlideLayout).
        ctx.batch.push({
          op: 'setAttribute',
          part: owner.part,
          node: root.id,
          name: { ns: '', local: 'showMasterSp' },
          value: write.hidden ? '0' : null,
        });
        return;
      }
      case 'solid':
        return setSolid(ctx, owner, write);
      case 'gradient':
        return setGradient(ctx, owner, write);
      case 'pattern':
        return setPattern(ctx, owner, write);
      case 'picture':
        return setPicture(ctx, owner, write);
      default:
        badArgument('background');
    }
  });
  return { kind: 'applied' };
}
