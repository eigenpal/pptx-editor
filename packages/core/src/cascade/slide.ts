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
  relTarget,
  shapeTreeOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../store/index.ts';
import { contentElements } from './alternate-content.ts';
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
import { resolveFrameTable } from './table.ts';
import { DEFAULT_COLOR_MAP, parseTheme, readColorMap, type Theme } from './theme.ts';
import { fontRefEntry, levelEntries, resolveParagraph, type FontRef } from './text-style.ts';
import type {
  Autofit,
  BodyProps,
  GeometrySpec,
  Layer,
  ResolvedElement,
  ResolvedPicture,
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
  readonly layoutPlaceholders: readonly PlaceholderEntry[];
  readonly masterPlaceholders: readonly PlaceholderEntry[];
}

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

function collectPlaceholders(pkg: PptxPackage, part: string | undefined): PlaceholderEntry[] {
  const { tree, root } = treeOf(pkg, part);
  const out: PlaceholderEntry[] = [];
  for (const el of contentElements(tree, root)) {
    const key = placeholderOf(el);
    if (key) out.push({ ...key, el });
  }
  return out;
}

/** The master placeholder type a layout/slide placeholder type falls back to. */
function masterTypeFor(type: string): string {
  if (type === 'ctrTitle') return 'title';
  if (type === 'subTitle' || type === 'obj' || type === 'tbl' || type === 'chart' || type === 'pic')
    return 'body';
  return type;
}

function findLayoutPlaceholder(
  key: PlaceholderKey,
  list: readonly PlaceholderEntry[]
): Element | undefined {
  if (key.idx !== undefined && key.idx !== '0') {
    const byIdx = list.find((p) => p.idx === key.idx);
    if (byIdx) return byIdx.el;
  }
  return (
    list.find((p) => p.type === key.type)?.el ??
    list.find((p) => masterTypeFor(p.type) === masterTypeFor(key.type) && key.type !== 'obj')?.el
  );
}

function findMasterPlaceholder(
  type: string,
  list: readonly PlaceholderEntry[]
): Element | undefined {
  const want = masterTypeFor(type);
  return (
    list.find((p) => p.type === want)?.el ?? list.find((p) => masterTypeFor(p.type) === want)?.el
  );
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

function nonVisualProps(el: Element): Element | undefined {
  const nv = elements(el).find((c) => c.ns === NS.p && c.local.startsWith('nv'));
  return child(nv, 'p:cNvPr');
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
  const txBody = child(el, 'p:txBody');
  if (!txBody) return undefined;
  const bodyChain: Element[] = [];
  const lstChain: Element[] = [];
  for (const source of inherited) {
    const tb = child(source, 'p:txBody');
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
  const paragraphs = elements(txBody)
    .filter((p) => p.kind === 'paragraph')
    .map((p) =>
      resolveParagraph({
        paragraph: p,
        chainFor,
        ctx: scope.ctx,
        theme: scope.theme,
        ...(fontRef ? { fontRef } : {}),
        slideNumber: scope.slideNumber,
      })
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
  const spPrOf = (e: Element): Element | undefined => child(e, 'p:spPr');
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
    style ??= child(chain[i], 'p:style');
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
  const base = {
    nodeId: el.id,
    partName,
    layer,
    name: attr(cNvPr, 'name') ?? '',
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
  return { ...base, kind, fill, ...(text ? { text } : {}), ...(textBox ? { textBox } : {}) };
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
      return {
        kind: 'graphicFrame',
        nodeId: el.id,
        partName,
        layer,
        name: attr(cNvPr, 'name') ?? '',
        xfrm,
        hidden: boolAttr(cNvPr, 'hidden') ?? false,
        uri,
        frameKind: kind,
        ...(table ? { table } : {}),
        ...(poster ? { poster } : {}),
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
  const scope: SlideScope = {
    pkg,
    theme,
    ctx: { scheme: theme.colors, colorMap },
    slideNumber: index + 1,
    ...(child(presentation, 'p:defaultTextStyle')
      ? { defaultTextStyle: child(presentation, 'p:defaultTextStyle') }
      : {}),
    ...(child(master, 'p:txStyles') ? { txStyles: child(master, 'p:txStyles') } : {}),
    layoutPlaceholders: collectPlaceholders(pkg, ref.layoutPart),
    masterPlaceholders: collectPlaceholders(pkg, ref.masterPart),
  };
  const bg = resolveBackground([slide, layout, master], scope);
  const backgroundPart = [ref.partName, ref.layoutPart, ref.masterPart][bg.index] ?? ref.partName;
  const showMaster = (el: Element | undefined): boolean => attr(el, 'showMasterSp') !== '0';
  const out: ResolvedElement[] = [];
  const addLayer = (layer: Layer, partName: string | undefined): void => {
    const { tree, root } = treeOf(pkg, partName);
    if (!tree || !partName) return;
    for (const el of contentElements(tree, root)) {
      // Layout and master placeholders are templates, not content: they never paint on a slide.
      if (layer !== 'slide' && placeholderOf(el)) continue;
      const r = resolveElement(el, layer, partName, scope, 0);
      if (r) out.push(r);
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
