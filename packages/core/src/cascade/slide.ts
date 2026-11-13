// Resolve one slide: master → layout → slide inheritance, color map, background, and every
// element's final geometry, fill, line, effects, and text. The only module that knows how
// PresentationML inheritance works; layout and paint read its output only.

import {
  attr,
  binaryPartOf,
  boolAttr,
  child,
  elements,
  imageSize,
  metafileFormat,
  intAttr,
  isEmptiedTable,
  NS,
  path,
  presentationInfo,
  REL,
  relTarget,
  relTargetByType,
  isTemplateRef,
  slideRefAt,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../store/index.ts';
import { contentElements } from './alternate-content.ts';
import {
  collectPlaceholders,
  findLayoutPlaceholder,
  findMasterPlaceholder,
  partElementsOf,
  placeholderOf,
  treeOf,
  type PlaceholderEntry,
  type PlaceholderKey,
} from './placeholders.ts';
export { inheritedPlaceholders, layoutPlaceholderOf } from './placeholders.ts';
import { resolveChartPart } from './chart-part.ts';
import { colorChild, resolveColor, type ColorContext } from './color.ts';
import {
  effectContainer,
  fillElementOf,
  resolveEffects,
  resolveFillElement,
  resolveLine,
  styleRefContext,
  themeFill,
  type ResolvedEffects,
  type ResolvedFill,
} from './paint-props.ts';
import { entryParts } from './picture-bullet.ts';
import { layerKey, partContent, relsKey, sameKey } from './resolve-keys.ts';
import { resolveFrameTable } from './table.ts';
import { threeDOf } from './three-d.ts';
import { readTextWarp } from './text-warp.ts';
import {
  diagramParts,
  MAX_DIAGRAM_GROUP_DEPTH,
  MAX_DIAGRAM_SHAPES,
  MAX_SLIDE_DIAGRAM_SHAPES,
} from './diagram.ts';
import { textNodesOf } from './diagram-model.ts';
import { readGeometry, readXfrm, ZERO_XFRM } from './shape-xfrm.ts';
import { DEFAULT_COLOR_MAP, parseTheme, readColorMap, type Theme } from './theme.ts';
import {
  fontRefEntry,
  levelEntries,
  readKinsoku,
  resolveParagraph,
  type FontRef,
} from './text-style.ts';
import type {
  Autofit,
  BodyProps,
  ConnectionRef,
  DiagramFallback,
  GeometrySpec,
  KinsokuRules,
  Layer,
  ResolvedElement,
  ResolvedPicture,
  ResolvedShape,
  ResolvedSlide,
  ResolvedTextBody,
  Xfrm,
} from './types.ts';

/** Everything element resolution needs about the slide it is on. */
interface SlideScope {
  readonly pkg: PptxPackage;
  readonly theme: Theme;
  readonly ctx: ColorContext;
  readonly slideNumber: number;
  readonly defaultTextStyle?: Element;
  readonly txStyles?: Element;
  readonly kinsoku?: KinsokuRules;
  readonly layoutPlaceholders: readonly PlaceholderEntry[];
  readonly masterPlaceholders: readonly PlaceholderEntry[];
  /** Shapes and groups resolved from diagram drawings on this slide so far. */
  readonly diagramShapes: { count: number };
  /** The color map, the slide number, and the kinsoku rules as text: part of `textKey`. */
  readonly textContext: string;
  /** The presentation, master, and layout parts, whose relationships picture bullets name. */
  readonly presentationPart: string;
  readonly masterPart?: string;
  readonly layoutPart?: string;
}

/**
 * Resolved paragraphs by their XML element, with what else they read (`resolveText`): a
 * keystroke in a long text body resolves the edited paragraph, not every paragraph.
 */
const resolvedParagraphs = new WeakMap<
  Element,
  { readonly key: readonly unknown[]; readonly value: ReturnType<typeof resolveParagraph> }
>();

const MAX_GROUP_DEPTH = 32;

/** The namespace of a shape's own children: `p:` on a slide, `dsp:` in a diagram drawing. */
function ownNs(el: Element | undefined): string {
  return el?.ns === NS.dsp ? NS.dsp : NS.p;
}

/** A child in the element's own namespace (`p:spPr`, or `dsp:spPr` in a diagram drawing). */
function ownChild(el: Element | undefined, local: string): Element | undefined {
  return child(el, { ns: ownNs(el), local });
}

function nonVisualProps(el: Element): Element | undefined {
  const ns = ownNs(el);
  const nv = elements(el).find((c) => c.ns === ns && c.local.startsWith('nv'));
  return child(nv, { ns, local: 'cNvPr' });
}

function readBodyProps(chain: readonly Element[]): BodyProps {
  const pick = (name: string): string | undefined => {
    for (let i = chain.length - 1; i >= 0; i -= 1) {
      const v = attr(chain[i], name);
      if (v !== undefined) return v;
    }
    return undefined;
  };
  const num = (name: string, d: number): number => {
    const raw = pick(name);
    const v = raw === undefined ? NaN : Number(raw);
    return Number.isFinite(v) ? Math.max(-51_206_400, Math.min(v, 51_206_400)) : d;
  };
  let autofit: Autofit = { kind: 'none' };
  for (let i = chain.length - 1; i >= 0; i -= 1) {
    const norm = child(chain[i], 'a:normAutofit');
    if (norm) {
      // ECMA-376 Part 1, CT_TextNormalAutofit: fontScale is 1 %–100 %. A reduction above 100 %
      // has no meaning, so both clamp here, at the parse boundary.
      const scale = intAttr(norm, 'fontScale') ?? 100_000;
      const reduction = intAttr(norm, 'lnSpcReduction') ?? 0;
      autofit = {
        kind: 'normal',
        fontScale: Math.max(1_000, Math.min(scale, 100_000)) / 100_000,
        lineSpaceReduction: Math.max(0, Math.min(reduction, 100_000)) / 100_000,
      };
      break;
    }
    if (child(chain[i], 'a:spAutoFit')) {
      autofit = { kind: 'shape' };
      break;
    }
    if (child(chain[i], 'a:noAutofit')) break;
  }
  const anchor = pick('anchor');
  const warp = readTextWarp(chain);
  return {
    lIns: num('lIns', 91_440),
    tIns: num('tIns', 45_720),
    rIns: num('rIns', 91_440),
    bIns: num('bIns', 45_720),
    anchor:
      anchor === 'ctr' || anchor === 'b' || anchor === 'just' || anchor === 'dist' ? anchor : 't',
    anchorCtr: pick('anchorCtr') === '1',
    wrap: pick('wrap') === 'none' ? 'none' : 'square',
    vert: pick('vert') ?? 'horz',
    rot: num('rot', 0) / 60_000,
    upright: pick('upright') === '1',
    // ST_TextColumnCount is an integer from 1 to 16.
    numCol: Math.max(1, Math.min(Math.trunc(num('numCol', 1)), 16)),
    spcCol: Math.max(0, num('spcCol', 0)),
    autofit,
    spcFirstLastPara: pick('spcFirstLastPara') === '1',
    ...(warp ? { warp } : {}),
  };
}

function txStyleFor(type: string | undefined, txStyles: Element | undefined): Element | undefined {
  if (!type) return undefined;
  if (type === 'title' || type === 'ctrTitle') return child(txStyles, 'p:titleStyle');
  if (type === 'body' || type === 'subTitle' || type === 'obj')
    return child(txStyles, 'p:bodyStyle');
  return child(txStyles, 'p:otherStyle');
}

function resolveText(
  el: Element,
  partName: string,
  scope: SlideScope,
  ph: PlaceholderKey | undefined,
  inherited: readonly (readonly [Element, string | undefined])[],
  fontRef: FontRef | undefined
): ResolvedTextBody | undefined {
  const txBody = ownChild(el, 'txBody');
  if (!txBody) return undefined;
  const bodyChain: Element[] = [];
  const lstChain: Element[] = [];
  // Each list style with its part, for the relationships of picture bullets.
  const lists: (readonly [Element | undefined, string | undefined])[] = [];
  for (const [source, part] of inherited) {
    const tb = ownChild(source, 'txBody');
    const bp = child(tb, 'a:bodyPr');
    if (bp) bodyChain.push(bp);
    const ls = child(tb, 'a:lstStyle');
    if (ls) lstChain.push(ls);
    lists.push([ls, part]);
  }
  const ownBody = child(txBody, 'a:bodyPr');
  if (ownBody) bodyChain.push(ownBody);
  const ownList = child(txBody, 'a:lstStyle');
  if (ownList) lstChain.push(ownList);
  const masterStyle = txStyleFor(ph?.type, scope.txStyles);
  const pictures = {
    pkg: scope.pkg,
    partOf: entryParts(
      [
        [scope.defaultTextStyle, scope.presentationPart],
        [masterStyle, scope.masterPart],
        ...lists,
        [ownList, partName],
      ],
      partName
    ),
  };
  const refEntry = fontRefEntry(fontRef);
  const chainFor = (level: number): Element[] => [
    ...levelEntries(scope.defaultTextStyle, level),
    ...levelEntries(masterStyle, level),
    ...(refEntry ? [refEntry] : []),
    ...lstChain.flatMap((ls) => levelEntries(ls, level)),
  ];
  // Everything a paragraph reads besides its own XML: the level chains come from these.
  const key: readonly unknown[] = [
    scope.theme,
    scope.textContext,
    scope.defaultTextStyle,
    masterStyle,
    JSON.stringify(fontRef ?? null),
    ...lstChain,
  ];
  const cached = (p: Element) => {
    const hit = resolvedParagraphs.get(p);
    return hit && sameKey(hit.key, key) ? hit.value : undefined;
  };
  const remember = (p: Element, value: ReturnType<typeof resolveParagraph>) => {
    resolvedParagraphs.set(p, { key, value });
    return value;
  };
  const paragraphs = elements(txBody)
    .filter((p) => p.kind === 'paragraph')
    .map(
      (p) =>
        cached(p) ??
        remember(
          p,
          resolveParagraph({
            paragraph: p,
            chainFor,
            ctx: scope.ctx,
            theme: scope.theme,
            ...(fontRef ? { fontRef } : {}),
            slideNumber: scope.slideNumber,
            ...(scope.kinsoku ? { kinsoku: scope.kinsoku } : {}),
            pictures,
          })
        )
    );
  return { nodeId: txBody.id, body: readBodyProps(bodyChain), paragraphs };
}

function styleFontRef(style: Element | undefined, scope: SlideScope): FontRef | undefined {
  const ref = child(style, 'a:fontRef');
  if (!ref) return undefined;
  const idx = attr(ref, 'idx');
  const color = resolveColor(colorChild(ref), scope.ctx);
  return {
    ...(idx === 'major' || idx === 'minor' ? { font: idx } : {}),
    ...(color ? { color } : {}),
  };
}

function resolveShapeLike(
  el: Element,
  kind: 'shape' | 'connector' | 'picture',
  layer: Layer,
  partName: string,
  scope: SlideScope
): ResolvedElement {
  const ph = placeholderOf(el);
  const inherited: Element[] = [];
  // The inherited placeholders with their parts, for picture bullets.
  const sources: (readonly [Element, string | undefined])[] = [];
  if (ph && layer === 'slide') {
    const layoutPh = findLayoutPlaceholder(ph, scope.layoutPlaceholders);
    const masterPh = findMasterPlaceholder(
      layoutPh ? (placeholderOf(layoutPh)?.type ?? ph.type) : ph.type,
      scope.masterPlaceholders
    );
    if (masterPh) sources.push([masterPh, scope.masterPart]);
    if (layoutPh) sources.push([layoutPh, scope.layoutPart]);
  } else if (ph && layer === 'layout') {
    const masterPh = findMasterPlaceholder(ph.type, scope.masterPlaceholders);
    if (masterPh) sources.push([masterPh, scope.masterPart]);
  }
  for (const [source] of sources) inherited.push(source);
  const spPrOf = (e: Element): Element | undefined => ownChild(e, 'spPr');
  const chain = [...inherited, el];
  let xfrm: Xfrm | undefined;
  let geometry: GeometrySpec | undefined;
  let fillEl: Element | undefined;
  let lnEl: Element | undefined;
  let effects: Element | undefined;
  let style: Element | undefined;
  for (let i = chain.length - 1; i >= 0; i -= 1) {
    const spPr = spPrOf(chain[i]);
    xfrm ??= readXfrm(child(spPr, 'a:xfrm'));
    geometry ??= readGeometry(spPr);
    fillEl ??= fillElementOf(spPr);
    lnEl ??= child(spPr, 'a:ln');
    effects ??= effectContainer(spPr);
    style ??= ownChild(chain[i], 'style');
  }
  const ctx = scope.ctx;
  let fill: ResolvedFill = { type: 'none' };
  if (kind === 'picture') {
    const blip = resolveFillElement(child(el, 'p:blipFill'), ctx);
    fill = blip && blip.type === 'picture' ? blip : { type: 'none' };
  } else if (fillEl) fill = resolveFillElement(fillEl, ctx) ?? { type: 'none' };
  else {
    const ref = child(style, 'a:fillRef');
    const idx = intAttr(ref, 'idx') ?? 0;
    const themed = themeFill(idx, scope.theme);
    if (themed) fill = resolveFillElement(themed, styleRefContext(ref, ctx)) ?? { type: 'none' };
  }
  const lnRef = child(style, 'a:lnRef');
  const lnIdx = intAttr(lnRef, 'idx') ?? 0;
  const styleLine = lnIdx > 0 ? scope.theme.lineStyles[lnIdx - 1] : undefined;
  const line = resolveLine(lnEl, styleLine, ctx, styleRefContext(lnRef, ctx));
  const effectRef = child(style, 'a:effectRef');
  const effIdx = intAttr(effectRef, 'idx') ?? 0;
  const effectStyle = effIdx > 0 ? scope.theme.effectStyles[effIdx - 1] : undefined;
  const styleCtx = styleRefContext(effectRef, ctx);
  const listed = effects
    ? resolveEffects(effects, ctx)
    : effectStyle && resolveEffects(effectStyle, styleCtx);
  const threeD = threeDOf(chain.map(spPrOf), effectStyle, ctx, styleCtx);
  const resolvedEffects = threeD ? { unsupported: [], ...listed, threeD } : listed;
  const cNvPr = nonVisualProps(el);
  const drawingId = el.ns === NS.p ? attr(cNvPr, 'id') : undefined;
  const base = {
    nodeId: el.id,
    partName,
    layer,
    name: attr(cNvPr, 'name') ?? '',
    ...(drawingId !== undefined ? { drawingId } : {}),
    xfrm: xfrm ?? ZERO_XFRM,
    hidden: boolAttr(cNvPr, 'hidden') ?? false,
    ...(ph ? { placeholder: ph } : {}),
    ...(line && line.fill.type !== 'none' ? { line } : {}),
    ...(resolvedEffects ? { effects: resolvedEffects } : {}),
    geometry: geometry ?? ({ kind: 'preset', preset: 'rect', adjust: new Map() } as GeometrySpec),
  };
  if (kind === 'picture') {
    const media = path(el, 'p:nvPicPr', 'p:nvPr');
    const hasMedia = elements(media).some(
      (m) => m.local === 'videoFile' || m.local === 'audioFile'
    );
    return {
      ...base,
      kind: 'picture',
      fill: fill.type === 'picture' ? fill : { type: 'none' },
      ...(hasMedia ? { media: 'media' } : {}),
    };
  }
  const text =
    kind === 'shape'
      ? resolveText(el, partName, scope, ph, sources, styleFontRef(style, scope))
      : undefined;
  // ECMA-376 Part 1, 19.3.1.13 cNvSpPr: txBox marks a text box.
  const textBox = kind === 'shape' && boolAttr(path(el, 'p:nvSpPr', 'p:cNvSpPr'), 'txBox') === true;
  // A diagram shape names its own text rectangle (`dsp:txXfrm`).
  const textXfrm = el.ns === NS.dsp ? readXfrm(child(el, 'dsp:txXfrm')) : undefined;
  const connections = kind === 'connector' ? connectionsOf(el) : undefined;
  return {
    ...base,
    kind,
    fill,
    ...(connections ? { connections } : {}),
    ...(text ? { text } : {}),
    ...(textBox ? { textBox } : {}),
    ...(text && textXfrm ? { textXfrm } : {}),
  };
}

/** A connector's glued ends (`p:nvCxnSpPr/p:cNvCxnSpPr/a:stCxn` and `a:endCxn`). */
function connectionsOf(el: Element): ResolvedShape['connections'] {
  const props = path(el, 'p:nvCxnSpPr', 'p:cNvCxnSpPr');
  const end = (name: string): ConnectionRef | undefined => {
    const c = child(props, name);
    const id = attr(c, 'id');
    const idx = intAttr(c, 'idx');
    return id !== undefined && idx !== undefined && idx >= 0 ? { id, idx } : undefined;
  };
  const start = end('a:stCxn');
  const finish = end('a:endCxn');
  return start || finish
    ? { ...(start ? { start } : {}), ...(finish ? { end: finish } : {}) }
    : undefined;
}

function frameKind(uri: string): 'chart' | 'diagram' | 'table' | 'ole' | 'unknown' {
  if (uri.endsWith('/chart')) return 'chart';
  if (uri.endsWith('/diagram')) return 'diagram';
  if (uri.endsWith('/table')) return 'table';
  if (uri.endsWith('/ole')) return 'ole';
  return 'unknown';
}

/** The most pixels an OLE poster may have; the PDF writer's picture limit, so both agree. */
const OLE_POSTER_MAX_PIXELS = 25_000_000;

/**
 * Whether a picture relationship names an internal package part whose bytes are a PNG, JPEG,
 * GIF, or BMP image of a sane size, or an EMF or WMF picture (its PNG has a bounded size,
 * `output/metafile-png.ts`). An external target (`TargetMode="External"`) is never fetched, so
 * it fails, as do a missing part and any other content.
 */
function drawableImage(pkg: PptxPackage, sourcePart: string, rId: string): boolean {
  const bytes = binaryPartOf(pkg, relTarget(pkg, sourcePart, rId));
  if (bytes && metafileFormat(bytes)) return true;
  const size = bytes ? imageSize(bytes) : null;
  return (
    !!size && size.width > 0 && size.height > 0 && size.width * size.height <= OLE_POSTER_MAX_PIXELS
  );
}

/**
 * The picture that shows an OLE object: `p:oleObj/p:pic` (ECMA-376 Part 1, 19.3.1.36; required
 * by `CT_OleObject`). The object is never run. An `mc:AlternateContent` picks its branch as
 * the shape tree does (`contentElements`). The picture fills the frame's box and carries the
 * frame's node id, so a hit on the picture selects the frame. A picture that cannot be drawn
 * (see `drawableImage`) gives no poster, so the frame keeps its labeled placeholder.
 */
function resolveOlePoster(
  frame: Element,
  graphicData: Element | undefined,
  xfrm: Xfrm,
  layer: Layer,
  partName: string,
  scope: SlideScope
): ResolvedPicture | undefined {
  const ole = contentElements(graphicData).find((e) => e.ns === NS.p && e.local === 'oleObj');
  const pic = elements(ole).find((e) => e.ns === NS.p && e.local === 'pic');
  if (!pic) return undefined;
  const resolved = resolveShapeLike(pic, 'picture', layer, partName, scope);
  if (resolved.kind !== 'picture' || resolved.fill.type !== 'picture') return undefined;
  if (!drawableImage(scope.pkg, partName, resolved.fill.rId)) return undefined;
  return { ...resolved, nodeId: frame.id, xfrm, hidden: false };
}

/** A shape of a diagram drawing that fails the supported subset: the whole diagram falls back. */
const UNSUPPORTED = Symbol('unsupported');

/** The shapes one diagram drawing has resolved, and whether a limit stopped it. */
interface DiagramBudget {
  shapes: number;
  overLimit: boolean;
  readonly nodes: ReadonlyMap<string, string>;
}

/** Stop a diagram at a shape, depth, or slide limit, and remember why. */
function overLimit(budget: DiagramBudget): typeof UNSUPPORTED {
  budget.overLimit = true;
  return UNSUPPORTED;
}

/**
 * One element of a diagram drawing's shape tree (`dsp:spTree`), in its parent's space. Shapes
 * resolve as slide shapes do. Undefined for the tree's own properties; {@link UNSUPPORTED} for
 * a shape with no box, a group past the depth limit, an element past the frame's or the
 * slide's shape budget, or any other element.
 */
function resolveDiagramElement(
  el: Element,
  layer: Layer,
  drawingPart: string,
  scope: SlideScope,
  depth: number,
  budget: DiagramBudget
): ResolvedElement | undefined | typeof UNSUPPORTED {
  if (el.ns !== NS.dsp) return UNSUPPORTED;
  switch (el.local) {
    case 'nvGrpSpPr':
    case 'grpSpPr':
    case 'extLst':
      return undefined;
    case 'sp': {
      budget.shapes += 1;
      scope.diagramShapes.count += 1;
      if (budget.shapes > MAX_DIAGRAM_SHAPES) return overLimit(budget);
      if (scope.diagramShapes.count > MAX_SLIDE_DIAGRAM_SHAPES) return overLimit(budget);
      if (!path(el, 'dsp:spPr', 'a:xfrm')) return UNSUPPORTED;
      const shape = resolveShapeLike(el, 'shape', layer, drawingPart, scope);
      const node = budget.nodes.get(attr(el, 'modelId') ?? '');
      return node && shape.kind === 'shape' ? { ...shape, diagramNode: node } : shape;
    }
    case 'grpSp': {
      scope.diagramShapes.count += 1;
      if (depth >= MAX_DIAGRAM_GROUP_DEPTH) return overLimit(budget);
      if (scope.diagramShapes.count > MAX_SLIDE_DIAGRAM_SHAPES) return overLimit(budget);
      const kids = resolveDiagramTree(el, layer, drawingPart, scope, depth + 1, budget);
      if (!kids) return UNSUPPORTED;
      return diagramGroup(el, kids, layer, drawingPart, scope);
    }
  }
  return UNSUPPORTED;
}

function resolveDiagramTree(
  tree: Element,
  layer: Layer,
  drawingPart: string,
  scope: SlideScope,
  depth: number,
  budget: DiagramBudget
): ResolvedElement[] | undefined {
  const out: ResolvedElement[] = [];
  for (const c of elements(tree)) {
    const r = resolveDiagramElement(c, layer, drawingPart, scope, depth, budget);
    if (r === UNSUPPORTED) return undefined;
    if (r) out.push(r);
  }
  return out;
}

/** A `dsp:grpSp`, or the `dsp:spTree` root when its `dsp:grpSpPr` places its children. */
function diagramGroup(
  el: Element,
  children: readonly ResolvedElement[],
  layer: Layer,
  drawingPart: string,
  scope: SlideScope
): ResolvedElement {
  const grpSpPr = child(el, 'dsp:grpSpPr');
  const x = child(grpSpPr, 'a:xfrm');
  const xfrm = readXfrm(x) ?? ZERO_XFRM;
  const chOff = child(x, 'a:chOff');
  const chExt = child(x, 'a:chExt');
  const cNvPr = nonVisualProps(el);
  return {
    kind: 'group',
    nodeId: el.id,
    partName: drawingPart,
    layer,
    name: attr(cNvPr, 'name') ?? '',
    xfrm,
    hidden: boolAttr(cNvPr, 'hidden') ?? false,
    childOffset: { x: intAttr(chOff, 'x') ?? xfrm.x, y: intAttr(chOff, 'y') ?? xfrm.y },
    childExtent: { cx: intAttr(chExt, 'cx') ?? xfrm.cx, cy: intAttr(chExt, 'cy') ?? xfrm.cy },
    children,
    fill: resolveFillElement(fillElementOf(grpSpPr), scope.ctx) ?? { type: 'none' },
    ...groupEffects(grpSpPr, scope.ctx),
  };
}

/**
 * The effects of a group (`p:grpSpPr` or `dsp:grpSpPr`, ECMA-376 Part 1, 19.3.1.22): an effect
 * list applies to the group's content as one picture. A group has no style, so it has no theme
 * effect. A group's inner shadow draws nothing, so it is left out; it stays in the file. A
 * group's 3-D scene (`a:scene3d`) is reported unsupported.
 */
function groupEffects(
  grpSpPr: Element | undefined,
  ctx: ColorContext
): { effects?: ResolvedEffects } {
  const resolved = resolveEffects(effectContainer(grpSpPr), ctx);
  let listed = resolved;
  if (resolved?.innerShadow) {
    const { innerShadow: _drawsNothing, ...rest } = resolved;
    listed = rest;
  }
  if (!child(grpSpPr, 'a:scene3d')) return listed ? { effects: listed } : {};
  return { effects: { ...listed, unsupported: [...(listed?.unsupported ?? []), 'scene3d'] } };
}

/**
 * The diagram background (`dgm:bg`) and outline (`dgm:whole/a:ln`) of the data model, as one
 * rectangle over the frame's box. Undefined when the data model sets neither.
 */
function diagramBackground(
  frame: Element,
  dataModel: Element,
  xfrm: Xfrm,
  layer: Layer,
  dataPart: string,
  scope: SlideScope
): ResolvedElement | undefined {
  const fill = resolveFillElement(fillElementOf(child(dataModel, 'dgm:bg')), scope.ctx);
  const line = resolveLine(path(dataModel, 'dgm:whole', 'a:ln'), undefined, scope.ctx, scope.ctx);
  const drawn = line && line.fill.type !== 'none' ? line : undefined;
  if ((!fill || fill.type === 'none') && !drawn) return undefined;
  return {
    kind: 'shape',
    nodeId: `${frame.id}#background`,
    partName: dataPart,
    layer,
    name: '',
    xfrm: { ...ZERO_XFRM, cx: xfrm.cx, cy: xfrm.cy },
    hidden: false,
    geometry: { kind: 'preset', preset: 'rect', adjust: new Map() },
    fill: fill ?? { type: 'none' },
    ...(drawn ? { line: drawn } : {}),
  };
}

/**
 * The saved drawing of a diagram frame (ECMA-376 Part 1, 21.4; see `diagram.ts`): its shapes in
 * the frame's local space, behind the data model's background. Otherwise the reason it shows
 * its placeholder: no saved drawing (or a part that does not resolve), an element outside the
 * supported subset, or a shape, depth, or slide limit. A diagram renders whole or not at all.
 */
function resolveDiagram(
  frame: Element,
  graphicData: Element | undefined,
  xfrm: Xfrm,
  layer: Layer,
  partName: string,
  scope: SlideScope
): { drawing: readonly ResolvedElement[] } | { drawingFallback: DiagramFallback } {
  const parts = diagramParts(scope.pkg, partName, graphicData);
  const tree = child(parts?.drawing, 'dsp:spTree');
  if (!parts || !tree) return { drawingFallback: 'no-drawing' };
  // `nodes`: the node whose text each drawing shape shows, by the shape's `modelId`.
  const nodes = textNodesOf(parts.dataModel);
  const budget: DiagramBudget = { shapes: 0, overLimit: false, nodes };
  const shapes = resolveDiagramTree(tree, layer, parts.drawingPart, scope, 0, budget);
  if (!shapes) return { drawingFallback: budget.overLimit ? 'limit' : 'unsupported' };
  if (shapes.length === 0) return { drawingFallback: 'no-drawing' };
  const root = child(child(tree, 'dsp:grpSpPr'), 'a:xfrm')
    ? [diagramGroup(tree, shapes, layer, parts.drawingPart, scope)]
    : shapes;
  const background = diagramBackground(frame, parts.dataModel, xfrm, layer, parts.dataPart, scope);
  return { drawing: background ? [background, ...root] : root };
}

function resolveElement(
  el: Element,
  layer: Layer,
  partName: string,
  scope: SlideScope,
  depth: number
): ResolvedElement | undefined {
  switch (el.kind) {
    case 'sp':
      return resolveShapeLike(el, 'shape', layer, partName, scope);
    case 'cxnSp':
      return resolveShapeLike(el, 'connector', layer, partName, scope);
    case 'pic':
      return resolveShapeLike(el, 'picture', layer, partName, scope);
    case 'grpSp': {
      if (depth >= MAX_GROUP_DEPTH) return undefined;
      const grpSpPr = child(el, 'p:grpSpPr');
      const x = child(grpSpPr, 'a:xfrm');
      const xfrm = readXfrm(x) ?? ZERO_XFRM;
      const chOff = child(x, 'a:chOff');
      const chExt = child(x, 'a:chExt');
      const kids: ResolvedElement[] = [];
      for (const c of contentElements(el, xmlPartOf(scope.pkg, partName)?.root)) {
        const r = resolveElement(c, layer, partName, scope, depth + 1);
        if (r) kids.push(r);
      }
      const cNvPr = nonVisualProps(el);
      return {
        kind: 'group',
        nodeId: el.id,
        partName,
        layer,
        name: attr(cNvPr, 'name') ?? '',
        xfrm,
        hidden: boolAttr(cNvPr, 'hidden') ?? false,
        childOffset: { x: intAttr(chOff, 'x') ?? xfrm.x, y: intAttr(chOff, 'y') ?? xfrm.y },
        childExtent: { cx: intAttr(chExt, 'cx') ?? xfrm.cx, cy: intAttr(chExt, 'cy') ?? xfrm.cy },
        children: kids,
        fill: resolveFillElement(fillElementOf(grpSpPr), scope.ctx) ?? { type: 'none' },
        ...groupEffects(grpSpPr, scope.ctx),
      };
    }
    case 'graphicFrame': {
      const graphicData = path(el, 'a:graphic', 'a:graphicData');
      const uri = attr(graphicData, 'uri') ?? '';
      const kind = frameKind(uri);
      // A table that concurrent deletes left with no column or row is not drawn: the save
      // leaves it out (`isEmptiedTable`).
      if (kind === 'table' && isEmptiedTable(el)) return undefined;
      const table = kind === 'table' ? resolveFrameTable(el, { ...scope, partName }) : undefined;
      const cNvPr = nonVisualProps(el);
      const xfrm = readXfrm(child(el, 'p:xfrm')) ?? ZERO_XFRM;
      const poster =
        kind === 'ole'
          ? resolveOlePoster(el, graphicData, xfrm, layer, partName, scope)
          : undefined;
      const diagram =
        kind === 'diagram'
          ? resolveDiagram(el, graphicData, xfrm, layer, partName, scope)
          : undefined;
      const chart =
        kind === 'chart'
          ? resolveChartPart(scope.pkg, partName, graphicData, scope.ctx, scope.theme)
          : undefined;
      const frameId = attr(cNvPr, 'id');
      return {
        kind: 'graphicFrame',
        nodeId: el.id,
        partName,
        layer,
        name: attr(cNvPr, 'name') ?? '',
        ...(frameId !== undefined ? { drawingId: frameId } : {}),
        xfrm,
        hidden: boolAttr(cNvPr, 'hidden') ?? false,
        uri,
        frameKind: kind,
        ...(table ? { table } : {}),
        ...(poster ? { poster } : {}),
        ...diagram,
        ...(chart ? { chart } : {}),
      };
    }
  }
  return undefined;
}

function resolveBackground(
  sources: readonly (Element | undefined)[],
  scope: SlideScope
): { fill: ResolvedFill; index: number } {
  for (let i = 0; i < sources.length; i += 1) {
    const bg = path(sources[i], 'p:cSld', 'p:bg');
    if (!bg) continue;
    const bgPr = child(bg, 'p:bgPr');
    if (bgPr) {
      const fill = resolveFillElement(fillElementOf(bgPr), scope.ctx);
      if (fill) return { fill, index: i };
    }
    const bgRef = child(bg, 'p:bgRef');
    if (bgRef) {
      const themed = themeFill(intAttr(bgRef, 'idx') ?? 0, scope.theme);
      const fill = themed
        ? resolveFillElement(themed, styleRefContext(bgRef, scope.ctx))
        : undefined;
      if (fill) return { fill, index: i };
    }
  }
  return { fill: { type: 'solid', color: { r: 255, g: 255, b: 255, a: 1 } }, index: -1 };
}

/** An element of the edited template, as the editable layer of the slide master view. */
function asSlideLayer(el: ResolvedElement): ResolvedElement {
  const children = 'children' in el ? el.children.map(asSlideLayer) : undefined;
  return { ...el, layer: 'slide', ...(children ? { children } : {}) } as ResolvedElement;
}

/**
 * Resolved elements by their XML element, for the next resolution of the same slide. The
 * package is persistent: an edit replaces only the elements on the path it changes, so an
 * element that keeps its object resolves the same, when everything else it reads is the same
 * too (`layerKey`). A keystroke then resolves the edited shape, not the whole slide.
 */
const resolvedByElement = new WeakMap<
  Element,
  { readonly key: readonly unknown[]; readonly value: ResolvedElement }
>();

/** Resolve slide `index` (0-based, presentation order). */
export function resolveSlide(pkg: PptxPackage, index: number): ResolvedSlide | undefined {
  const info = presentationInfo(pkg);
  const ref = slideRefAt(pkg, index);
  if (!ref) return undefined;
  // In the slide master view, a layout or a master is the slide: its own shapes, placeholders
  // too, are the content, and a layout shows its master's shapes under them.
  const template = isTemplateRef(ref) ? ref.template : undefined;
  const theme = parseTheme(xmlPartOf(pkg, ref.themePart));
  const slide = xmlPartOf(pkg, ref.partName)?.root;
  // A master in the view is its own master: its color map and text styles apply.
  const master = template === 'master' ? slide : xmlPartOf(pkg, ref.masterPart)?.root;
  const layout = xmlPartOf(pkg, ref.layoutPart)?.root;
  let colorMap = readColorMap(child(master, 'p:clrMap')) ?? DEFAULT_COLOR_MAP;
  for (const over of [layout, template === 'master' ? undefined : slide]) {
    const mapping = path(over, 'p:clrMapOvr', 'a:overrideClrMapping');
    const read = readColorMap(mapping);
    if (read) colorMap = read;
  }
  const presentation = xmlPartOf(pkg, info.presentationPart)?.root;
  const kinsoku = readKinsoku(presentation);
  const scope: SlideScope = {
    pkg,
    theme,
    ctx: { scheme: theme.colors, colorMap },
    // A template has no slide number (0): its field shows its saved text.
    slideNumber: template ? 0 : index + 1,
    ...(child(presentation, 'p:defaultTextStyle')
      ? { defaultTextStyle: child(presentation, 'p:defaultTextStyle') }
      : {}),
    ...(child(master, 'p:txStyles') ? { txStyles: child(master, 'p:txStyles') } : {}),
    ...(kinsoku ? { kinsoku } : {}),
    layoutPlaceholders: collectPlaceholders(pkg, ref.layoutPart),
    masterPlaceholders: collectPlaceholders(
      pkg,
      template === 'master' ? ref.partName : ref.masterPart
    ),
    diagramShapes: { count: 0 },
    textContext: JSON.stringify([[...colorMap], template ? 0 : index + 1, kinsoku ?? null]),
    presentationPart: info.presentationPart,
    ...(ref.masterPart ? { masterPart: ref.masterPart } : {}),
    ...(ref.layoutPart ? { layoutPart: ref.layoutPart } : {}),
  };
  const bg = resolveBackground(
    template === 'master' ? [slide, undefined, undefined] : [slide, layout, master],
    scope
  );
  const backgroundPart = [ref.partName, ref.layoutPart, ref.masterPart][bg.index] ?? ref.partName;
  const showMaster = (el: Element | undefined): boolean => attr(el, 'showMasterSp') !== '0';
  const out: ResolvedElement[] = [];
  // Everything the elements of every layer read from outside their part (see `layerKey`).
  const shared: unknown[] = [
    theme,
    master,
    layout,
    index,
    info.width,
    info.height,
    child(slide, 'p:clrMapOvr'),
    child(presentation, 'p:defaultTextStyle'),
    JSON.stringify(kinsoku ?? null),
    partContent(pkg, relTargetByType(pkg, info.presentationPart, REL.tableStyles)),
    // A slide placeholder inherits pictures and links through the layout and the master.
    ...relsKey(pkg, ref.layoutPart),
    ...relsKey(pkg, ref.masterPart),
  ];
  const addLayer = (layer: Layer, partName: string | undefined, own = layer === 'slide'): void => {
    const { tree, root } = treeOf(pkg, partName);
    if (!tree || !partName) return;
    const key = layerKey(pkg, layer, partName, root, shared);
    // Layout and master placeholders are templates, not content: they never paint on a slide.
    // The template being edited shows them, as the slide master view does.
    const content = own
      ? contentElements(tree, root)
      : (partElementsOf(pkg, partName)?.others ?? []);
    for (const el of content) {
      const hit = resolvedByElement.get(el);
      if (hit && sameKey(hit.key, key)) {
        out.push(hit.value);
        continue;
      }
      const diagrams = scope.diagramShapes.count;
      const resolved = resolveElement(el, layer, partName, scope, 0);
      if (!resolved) continue;
      // The edited template's shapes resolve as their own layer (a layout placeholder inherits
      // from the master) and are the editable layer of the view.
      const r = own && layer !== 'slide' ? asSlideLayer(resolved) : resolved;
      out.push(r);
      // A diagram's shapes count across the slide: an element with diagrams always resolves.
      if (scope.diagramShapes.count === diagrams) resolvedByElement.set(el, { key, value: r });
    }
  };
  if (template === 'master') addLayer('master', ref.partName, true);
  else if (template === 'layout') {
    if (showMaster(slide)) addLayer('master', ref.masterPart);
    addLayer('layout', ref.partName, true);
  } else {
    if (showMaster(slide) && showMaster(layout)) addLayer('master', ref.masterPart);
    if (showMaster(slide)) addLayer('layout', ref.layoutPart);
    addLayer('slide', ref.partName);
  }
  return {
    partName: ref.partName,
    index,
    width: info.width,
    height: info.height,
    background: bg.fill,
    backgroundPart,
    elements: out,
  };
}
