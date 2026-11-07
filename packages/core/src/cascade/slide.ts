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
  intAttr,
  isEmptiedTable,
  NS,
  path,
  presentationInfo,
  REL,
  relTarget,
  relTargetByType,
  shapeTreeOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../store/index.ts';
import { contentElements } from './alternate-content.ts';
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
  type ResolvedFill,
} from './paint-props.ts';
import { layerKey, partContent, relsKey, sameKey } from './resolve-keys.ts';
import { resolveFrameTable } from './table.ts';
import {
  diagramParts,
  MAX_DIAGRAM_GROUP_DEPTH,
  MAX_DIAGRAM_SHAPES,
  MAX_SLIDE_DIAGRAM_SHAPES,
} from './diagram.ts';
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

/** Placeholder identity on a shape. */
interface PlaceholderKey {
  readonly type: string;
  readonly idx?: string;
}

interface PlaceholderEntry extends PlaceholderKey {
  readonly el: Element;
}

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

function placeholderOf(el: Element): PlaceholderKey | undefined {
  const nv = elements(el).find((c) => c.local.startsWith('nv') && c.ns === NS.p);
  const ph = path(nv, 'p:nvPr', 'p:ph');
  if (!ph) return undefined;
  const idx = attr(ph, 'idx');
  return { type: attr(ph, 'type') ?? 'obj', ...(idx !== undefined ? { idx } : {}) };
}

/** The shape tree of a part and the part root, where `mc:Choice` prefixes may be declared. */
function treeOf(pkg: PptxPackage, part: string | undefined): { tree?: Element; root?: Element } {
  if (!part) return {};
  const tree = shapeTreeOf(pkg, part);
  const root = xmlPartOf(pkg, part)?.root;
  return { ...(tree ? { tree } : {}), ...(root ? { root } : {}) };
}

/** A part's content elements, split into placeholders and the rest. */
interface PartElements {
  readonly tree: Element;
  readonly placeholders: readonly PlaceholderEntry[];
  /** The content elements that are not placeholders, in document order. */
  readonly others: readonly Element[];
}

/**
 * The content of a part by its root, which an edit of the part replaces: a layout or a master
 * splits into placeholders and content once, not on each resolution of each slide.
 */
const partElements = new WeakMap<Element, PartElements>();

function partElementsOf(pkg: PptxPackage, part: string | undefined): PartElements | undefined {
  const { tree, root } = treeOf(pkg, part);
  if (!tree) return undefined;
  const hit = root ? partElements.get(root) : undefined;
  if (hit && hit.tree === tree) return hit;
  const placeholders: PlaceholderEntry[] = [];
  const others: Element[] = [];
  for (const el of contentElements(tree, root)) {
    const key = placeholderOf(el);
    if (key) placeholders.push({ ...key, el });
    else others.push(el);
  }
  const out = { tree, placeholders, others };
  if (root) partElements.set(root, out);
  return out;
}

function collectPlaceholders(
  pkg: PptxPackage,
  part: string | undefined
): readonly PlaceholderEntry[] {
  return partElementsOf(pkg, part)?.placeholders ?? [];
}

/** The master placeholder type a layout/slide placeholder type falls back to. */
function masterTypeFor(type: string): string {
  if (type === 'ctrTitle') return 'title';
  if (type === 'subTitle' || type === 'obj' || type === 'tbl' || type === 'chart' || type === 'pic')
    return 'body';
  return type;
}

/** The first entry of each idx, type, and master type of a placeholder list. */
interface PlaceholderIndex {
  readonly byIdx: ReadonlyMap<string, Element>;
  readonly byType: ReadonlyMap<string, Element>;
  readonly byMasterType: ReadonlyMap<string, Element>;
}

/**
 * Lookups by key, built once per list: a slide can hold thousands of placeholders that match
 * nothing, and a scan per placeholder made that quadratic.
 */
const placeholderIndexes = new WeakMap<readonly PlaceholderEntry[], PlaceholderIndex>();

function placeholderIndex(list: readonly PlaceholderEntry[]): PlaceholderIndex {
  let index = placeholderIndexes.get(list);
  if (!index) {
    const byIdx = new Map<string, Element>();
    const byType = new Map<string, Element>();
    const byMasterType = new Map<string, Element>();
    for (const p of list) {
      if (p.idx !== undefined && !byIdx.has(p.idx)) byIdx.set(p.idx, p.el);
      if (!byType.has(p.type)) byType.set(p.type, p.el);
      const master = masterTypeFor(p.type);
      if (!byMasterType.has(master)) byMasterType.set(master, p.el);
    }
    index = { byIdx, byType, byMasterType };
    placeholderIndexes.set(list, index);
  }
  return index;
}

function findLayoutPlaceholder(
  key: PlaceholderKey,
  list: readonly PlaceholderEntry[]
): Element | undefined {
  const index = placeholderIndex(list);
  if (key.idx !== undefined && key.idx !== '0') {
    const byIdx = index.byIdx.get(key.idx);
    if (byIdx) return byIdx;
  }
  return (
    index.byType.get(key.type) ??
    (key.type !== 'obj' ? index.byMasterType.get(masterTypeFor(key.type)) : undefined)
  );
}

function findMasterPlaceholder(
  type: string,
  list: readonly PlaceholderEntry[]
): Element | undefined {
  const want = masterTypeFor(type);
  const index = placeholderIndex(list);
  return index.byType.get(want) ?? index.byMasterType.get(want);
}

/** The largest offset or size the cascade reads from `a:xfrm`, in EMU (four slide widths). @internal */
export const XFRM_BOUND = 51_206_400 * 4;

function readXfrm(xfrm: Element | undefined): Xfrm | undefined {
  if (!xfrm) return undefined;
  const off = child(xfrm, 'a:off');
  const ext = child(xfrm, 'a:ext');
  if (!off && !ext) return undefined;
  const bound = (v: number | undefined): number =>
    Math.max(-XFRM_BOUND, Math.min(v ?? 0, XFRM_BOUND));
  return {
    x: bound(intAttr(off, 'x')),
    y: bound(intAttr(off, 'y')),
    cx: Math.max(0, bound(intAttr(ext, 'cx'))),
    cy: Math.max(0, bound(intAttr(ext, 'cy'))),
    rot: ((intAttr(xfrm, 'rot') ?? 0) / 60_000) % 360,
    flipH: boolAttr(xfrm, 'flipH') ?? false,
    flipV: boolAttr(xfrm, 'flipV') ?? false,
  };
}

const ZERO_XFRM: Xfrm = { x: 0, y: 0, cx: 0, cy: 0, rot: 0, flipH: false, flipV: false };

function readGeometry(spPr: Element | undefined): GeometrySpec | undefined {
  const prst = child(spPr, 'a:prstGeom');
  if (prst) {
    const adjust = new Map<string, number>();
    for (const gd of elements(child(prst, 'a:avLst'))) {
      const name = attr(gd, 'name');
      const fmla = attr(gd, 'fmla') ?? '';
      const m = /^val\s+(-?\d+)$/.exec(fmla.trim());
      if (name && m) adjust.set(name, Number(m[1]));
    }
    return { kind: 'preset', preset: attr(prst, 'prst') ?? 'rect', adjust };
  }
  const cust = child(spPr, 'a:custGeom');
  if (cust) return { kind: 'custom', element: cust };
  return undefined;
}

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
  scope: SlideScope,
  ph: PlaceholderKey | undefined,
  inherited: readonly Element[],
  fontRef: FontRef | undefined
): ResolvedTextBody | undefined {
  const txBody = ownChild(el, 'txBody');
  if (!txBody) return undefined;
  const bodyChain: Element[] = [];
  const lstChain: Element[] = [];
  for (const source of inherited) {
    const tb = ownChild(source, 'txBody');
    const bp = child(tb, 'a:bodyPr');
    if (bp) bodyChain.push(bp);
    const ls = child(tb, 'a:lstStyle');
    if (ls) lstChain.push(ls);
  }
  const ownBody = child(txBody, 'a:bodyPr');
  if (ownBody) bodyChain.push(ownBody);
  const ownList = child(txBody, 'a:lstStyle');
  if (ownList) lstChain.push(ownList);
  const masterStyle = txStyleFor(ph?.type, scope.txStyles);
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
  if (ph && layer === 'slide') {
    const layoutPh = findLayoutPlaceholder(ph, scope.layoutPlaceholders);
    const masterPh = findMasterPlaceholder(
      layoutPh ? (placeholderOf(layoutPh)?.type ?? ph.type) : ph.type,
      scope.masterPlaceholders
    );
    if (masterPh) inherited.push(masterPh);
    if (layoutPh) inherited.push(layoutPh);
  } else if (ph && layer === 'layout') {
    const masterPh = findMasterPlaceholder(ph.type, scope.masterPlaceholders);
    if (masterPh) inherited.push(masterPh);
  }
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
  const resolvedEffects = effects
    ? resolveEffects(effects, ctx)
    : effIdx > 0
      ? resolveEffects(scope.theme.effectStyles[effIdx - 1], styleRefContext(effectRef, ctx))
      : undefined;
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
      ? resolveText(el, scope, ph, inherited, styleFontRef(style, scope))
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
 * GIF, or BMP image of a sane size. An external target (`TargetMode="External"`) is never
 * fetched, so it fails, as do a missing part and any other content (an EMF or WMF poster).
 */
function drawableImage(pkg: PptxPackage, sourcePart: string, rId: string): boolean {
  const bytes = binaryPartOf(pkg, relTarget(pkg, sourcePart, rId));
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
      return resolveShapeLike(el, 'shape', layer, drawingPart, scope);
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
  };
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
  const budget: DiagramBudget = { shapes: 0, overLimit: false };
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
      };
    }
    case 'graphicFrame': {
      const graphicData = path(el, 'a:graphic', 'a:graphicData');
      const uri = attr(graphicData, 'uri') ?? '';
      const kind = frameKind(uri);
      // A table that concurrent deletes left with no column or row is not drawn: the save
      // leaves it out (`isEmptiedTable`).
      if (kind === 'table' && isEmptiedTable(el)) return undefined;
      const table = kind === 'table' ? resolveFrameTable(el, scope) : undefined;
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

/**
 * The placeholders a slide element inherits from, nearest first: its layout placeholder, then
 * the master placeholder, matched as `resolveSlide` matches them. Empty for an element that is
 * not a placeholder. An edit that starts an own value from the inherited one reads these.
 */
export function inheritedPlaceholders(pkg: PptxPackage, slidePart: string, el: Element): Element[] {
  const ph = placeholderOf(el);
  const ref = presentationInfo(pkg).slides.find((s) => s.partName === slidePart);
  if (!ph || !ref) return [];
  const layoutPh = findLayoutPlaceholder(ph, collectPlaceholders(pkg, ref.layoutPart));
  const masterPh = findMasterPlaceholder(
    layoutPh ? (placeholderOf(layoutPh)?.type ?? ph.type) : ph.type,
    collectPlaceholders(pkg, ref.masterPart)
  );
  return [layoutPh, masterPh].filter((e): e is Element => e !== undefined);
}

/**
 * The layout placeholder that a placeholder on slide `index` inherits from (the matching rule of
 * `findLayoutPlaceholder`), or undefined. The editor reads its prompt text from it. @internal
 */
export function layoutPlaceholderOf(
  pkg: PptxPackage,
  index: number,
  el: Element
): Element | undefined {
  const ref = presentationInfo(pkg).slides[index];
  const key = placeholderOf(el);
  if (!ref || !key) return undefined;
  return findLayoutPlaceholder(key, collectPlaceholders(pkg, ref.layoutPart));
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
  const ref = info.slides[index];
  if (!ref) return undefined;
  const theme = parseTheme(xmlPartOf(pkg, ref.themePart));
  const master = xmlPartOf(pkg, ref.masterPart)?.root;
  const layout = xmlPartOf(pkg, ref.layoutPart)?.root;
  const slide = xmlPartOf(pkg, ref.partName)?.root;
  let colorMap = readColorMap(child(master, 'p:clrMap')) ?? DEFAULT_COLOR_MAP;
  for (const over of [layout, slide]) {
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
    slideNumber: index + 1,
    ...(child(presentation, 'p:defaultTextStyle')
      ? { defaultTextStyle: child(presentation, 'p:defaultTextStyle') }
      : {}),
    ...(child(master, 'p:txStyles') ? { txStyles: child(master, 'p:txStyles') } : {}),
    ...(kinsoku ? { kinsoku } : {}),
    layoutPlaceholders: collectPlaceholders(pkg, ref.layoutPart),
    masterPlaceholders: collectPlaceholders(pkg, ref.masterPart),
    diagramShapes: { count: 0 },
    textContext: JSON.stringify([[...colorMap], index + 1, kinsoku ?? null]),
  };
  const bg = resolveBackground([slide, layout, master], scope);
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
  const addLayer = (layer: Layer, partName: string | undefined): void => {
    const { tree, root } = treeOf(pkg, partName);
    if (!tree || !partName) return;
    const key = layerKey(pkg, layer, partName, root, shared);
    // Layout and master placeholders are templates, not content: they never paint on a slide.
    const content =
      layer === 'slide'
        ? contentElements(tree, root)
        : (partElementsOf(pkg, partName)?.others ?? []);
    for (const el of content) {
      const hit = resolvedByElement.get(el);
      if (hit && sameKey(hit.key, key)) {
        out.push(hit.value);
        continue;
      }
      const diagrams = scope.diagramShapes.count;
      const r = resolveElement(el, layer, partName, scope, 0);
      if (!r) continue;
      out.push(r);
      // A diagram's shapes count across the slide: an element with diagrams always resolves.
      if (scope.diagramShapes.count === diagrams) resolvedByElement.set(el, { key, value: r });
    }
  };
  if (showMaster(slide) && showMaster(layout)) addLayer('master', ref.masterPart);
  if (showMaster(slide)) addLayer('layout', ref.layoutPart);
  addLayer('slide', ref.partName);
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
