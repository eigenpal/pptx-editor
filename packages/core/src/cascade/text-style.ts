// The text property cascade: one chain per paragraph level, lowest precedence first.
//
//   presentation defaultTextStyle → master txStyles (title / body / other)
//   → shape style fontRef → master placeholder lstStyle → layout placeholder lstStyle
//   → shape lstStyle → paragraph pPr → run rPr
//
// Each chain entry is a pPr-like element (`a:defPPr`, `a:lvlNpPr`, `a:pPr`). For every property
// the LAST entry that states it wins. Run defaults come from each entry's `a:defRPr`.

import {
  attr,
  boolAttr,
  child,
  children,
  elements,
  intAttr,
  NS,
  textOf,
  type Element,
} from '../store/index.ts';
import { colorChild, resolveColor, schemeColor, type ColorContext, type Rgba } from './color.ts';
import {
  fillElementOf,
  resolveFillElement,
  resolveLine,
  type ResolvedFill,
} from './paint-props.ts';
import { resolveThemeFont, type Theme } from './theme.ts';
import type { ResolvedBullet, ResolvedParagraph, ResolvedRun, RunProps, Spacing } from './types.ts';

/** The theme font and color a shape's `p:style/a:fontRef` contributes. */
export interface FontRef {
  readonly font?: 'major' | 'minor';
  readonly color?: Rgba;
}

/** A DrawingML element built by the cascade, not read from a part (ids are never addressed). */
export function syntheticElement(
  local: string,
  attrs: Readonly<Record<string, string>> = {},
  kids: readonly Element[] = []
): Element {
  return {
    type: 'element',
    id: 'synthetic',
    kind: 'generic',
    ns: NS.a,
    local,
    nsDecls: [],
    attrs: Object.entries(attrs).map(([name, value]) => ({ ns: '', local: name, value })),
    children: kids,
  };
}

/** An `a:srgbClr` element for a resolved color, keeping its alpha. */
export function colorElement(color: Rgba): Element {
  const hex = [color.r, color.g, color.b]
    .map((v) =>
      Math.round(Math.max(0, Math.min(255, v)))
        .toString(16)
        .padStart(2, '0')
    )
    .join('')
    .toUpperCase();
  const alpha =
    color.a < 1 ? [syntheticElement('alpha', { val: String(Math.round(color.a * 100_000)) })] : [];
  return syntheticElement('srgbClr', { val: hex }, alpha);
}

/**
 * A style's font reference as a chain entry (`a:defPPr/a:defRPr`): its theme font and color
 * then take precedence over everything below it in the chain.
 */
export function fontRefEntry(ref: FontRef | undefined): Element | undefined {
  if (!ref || (!ref.font && !ref.color)) return undefined;
  const kids: Element[] = [];
  if (ref.color) kids.push(syntheticElement('solidFill', {}, [colorElement(ref.color)]));
  if (ref.font) {
    const m = ref.font === 'major' ? 'mj' : 'mn';
    kids.push(syntheticElement('latin', { typeface: `+${m}-lt` }));
    kids.push(syntheticElement('ea', { typeface: `+${m}-ea` }));
    kids.push(syntheticElement('cs', { typeface: `+${m}-cs` }));
  }
  return syntheticElement('defPPr', {}, [syntheticElement('defRPr', {}, kids)]);
}

/** The pPr-like entries for one level of a list style (`defPPr` first, then `lvlNpPr`). */
export function levelEntries(lstStyle: Element | undefined, level: number): Element[] {
  if (!lstStyle) return [];
  const out: Element[] = [];
  const def = child(lstStyle, 'a:defPPr');
  if (def) out.push(def);
  const lvl = child(lstStyle, `a:lvl${Math.min(9, Math.max(1, level + 1))}pPr`);
  if (lvl) out.push(lvl);
  return out;
}

function lastAttr(chain: readonly Element[], name: string): string | undefined {
  for (let i = chain.length - 1; i >= 0; i -= 1) {
    const v = attr(chain[i], name);
    if (v !== undefined) return v;
  }
  return undefined;
}

function lastChild(chain: readonly Element[], names: readonly string[]): Element | undefined {
  for (let i = chain.length - 1; i >= 0; i -= 1)
    for (const c of elements(chain[i])) if (c.ns === NS.a && names.includes(c.local)) return c;
  return undefined;
}

function spacing(el: Element | undefined, fallback: Spacing): Spacing {
  const pctEl = child(el, 'a:spcPct');
  // ST_TextSpacingPercentOrPercentString: 0 to 13200000 (132000 %). A file value outside it
  // would make line positions too large to compute.
  if (pctEl)
    return { pct: Math.max(0, Math.min(intAttr(pctEl, 'val') ?? 100_000, 13_200_000)) / 100_000 };
  const pts = child(el, 'a:spcPts');
  if (pts) return { pts: Math.max(0, Math.min(intAttr(pts, 'val') ?? 0, 158_400)) };
  return fallback;
}

const clampEmu = (v: number | undefined, d: number): number =>
  v === undefined ? d : Math.max(-51_206_400, Math.min(v, 51_206_400));

/**
 * The theme font the Latin family comes from: `+mj-lt` or `+mn-lt` (ECMA-376 Part 1,
 * 20.1.4.1.17 `latin`, theme font references), the shape style's `a:fontRef` (20.1.4.2.13), or
 * the theme's minor font when no level names a family. Undefined for a literal family.
 */
function latinThemeOf(
  face: string | undefined,
  own: string | undefined,
  fontRef: FontRef | undefined,
  refCollection: Theme['majorFont'] | undefined,
  theme: Theme
): 'major' | 'minor' | undefined {
  if (own !== undefined)
    return face === '+mj-lt' ? 'major' : face === '+mn-lt' ? 'minor' : undefined;
  if (refCollection?.latin) return fontRef?.font === 'major' ? 'major' : 'minor';
  return theme.minorFont.latin ? 'minor' : undefined;
}

/** Resolve run properties over a chain of rPr-like elements (lowest first). */
export function resolveRunProps(
  chain: readonly Element[],
  ctx: ColorContext,
  theme: Theme,
  fontRef: FontRef | undefined
): RunProps {
  const pick = (name: string): string | undefined => lastAttr(chain, name);
  // The typeface the nearest level names, as written: a family or a theme font reference.
  const faceOf = (slot: 'latin' | 'ea' | 'cs' | 'sym'): string | undefined => {
    for (let i = chain.length - 1; i >= 0; i -= 1) {
      const face = attr(child(chain[i], `a:${slot}`), 'typeface');
      if (face) return face;
    }
    return undefined;
  };
  const fontOf = (slot: 'latin' | 'ea' | 'cs' | 'sym'): string | undefined => {
    const face = faceOf(slot);
    return face ? resolveThemeFont(face, theme) : undefined;
  };
  const refCollection =
    fontRef?.font === 'major'
      ? theme.majorFont
      : fontRef?.font === 'minor'
        ? theme.minorFont
        : undefined;
  const latinFace = faceOf('latin');
  const ownLatin = latinFace ? resolveThemeFont(latinFace, theme) : undefined;
  const latin = ownLatin ?? refCollection?.latin ?? theme.minorFont.latin ?? 'Calibri';
  const latinTheme = latinThemeOf(latinFace, ownLatin, fontRef, refCollection, theme);
  const ea = fontOf('ea') ?? refCollection?.ea ?? theme.minorFont.ea ?? latin;
  const cs = fontOf('cs') ?? refCollection?.cs ?? theme.minorFont.cs ?? latin;
  const sym = fontOf('sym');
  let fill: ResolvedFill | undefined;
  for (let i = chain.length - 1; i >= 0 && !fill; i -= 1) {
    const f = fillElementOf(chain[i]);
    if (f) fill = resolveFillElement(f, ctx);
  }
  if (!fill && fontRef?.color) fill = { type: 'solid', color: fontRef.color };
  const highlightEl = lastChild(chain, ['highlight']);
  const highlight = highlightEl ? resolveColor(colorChild(highlightEl), ctx) : undefined;
  const lnEl = lastChild(chain, ['ln']);
  const outline = lnEl ? resolveLine(lnEl, undefined, ctx, ctx) : undefined;
  // ECMA-376 Part 1, DrawingML `EG_TextUnderlineFill`: `a:uFill` gives the underline its own
  // fill, and `a:uFillTx` makes it follow the text fill. The nearest level that sets one wins.
  const uFillEl = lastChild(chain, ['uFill', 'uFillTx']);
  const uFill = uFillEl?.local === 'uFill' ? fillElementOf(uFillEl) : undefined;
  const underlineFill = uFill ? resolveFillElement(uFill, ctx) : undefined;
  const capRaw = pick('cap');
  const size = Number(pick('sz') ?? '1800');
  const spc = Number(pick('spc') ?? '0');
  const baseline = Number(pick('baseline') ?? '0');
  const lang = pick('lang');
  const kernRaw = pick('kern');
  const kern = kernRaw !== undefined && /^\d{1,6}$/.test(kernRaw) ? Number(kernRaw) : undefined;
  const hlink = lastChild(chain, ['hlinkClick']);
  const hyperlink = attr(hlink, 'r:id');
  // Link text takes the theme's hyperlink color, unless the link asks for the text's own
  // color (`ahyp:hlinkClr val="tx"`, an `a:extLst` extension).
  const ownColor = elements(child(hlink, 'a:extLst')).some((ext) =>
    elements(ext).some((c) => c.local === 'hlinkClr' && attr(c, 'val') === 'tx')
  );
  if (hyperlink && !ownColor) {
    const color = schemeColor('hlink', ctx);
    if (color) fill = { type: 'solid', color };
  }
  return {
    size: Number.isFinite(size) ? Math.max(100, Math.min(size, 400_000)) : 1800,
    bold: (pick('b') ?? '0') === '1' || pick('b') === 'true',
    italic: (pick('i') ?? '0') === '1' || pick('i') === 'true',
    // Links are underlined unless the run says otherwise.
    underline: pick('u') ?? (hyperlink ? 'sng' : 'none'),
    strike: pick('strike') ?? 'noStrike',
    caps: capRaw === 'all' || capRaw === 'small' ? capRaw : 'none',
    spacing: Number.isFinite(spc) ? Math.max(-400_000, Math.min(spc, 400_000)) : 0,
    baseline: Number.isFinite(baseline) ? baseline / 1000 : 0,
    ...(kern !== undefined && kern > 0 ? { kern } : {}),
    latin,
    ...(latinTheme ? { latinTheme } : {}),
    ea,
    cs,
    ...(sym ? { sym } : {}),
    fill: fill ?? { type: 'solid', color: { r: 0, g: 0, b: 0, a: 1 } },
    ...(underlineFill ? { underlineFill } : {}),
    ...(highlight ? { highlight } : {}),
    ...(outline && outline.fill.type !== 'none' ? { outline } : {}),
    ...(lang ? { lang } : {}),
    ...(hyperlink ? { hyperlink } : {}),
  };
}

function resolveBullet(
  chain: readonly Element[],
  ctx: ColorContext,
  theme: Theme
): ResolvedBullet | null {
  const kind = lastChild(chain, ['buNone', 'buChar', 'buAutoNum', 'buBlip']);
  if (!kind || kind.local === 'buNone' || kind.local === 'buBlip') return null;
  const fontEl = lastChild(chain, ['buFont', 'buFontTx']);
  const sizeEl = lastChild(chain, ['buSzPct', 'buSzPts', 'buSzTx']);
  const colorEl = lastChild(chain, ['buClr', 'buClrTx']);
  const face = fontEl?.local === 'buFont' ? attr(fontEl, 'typeface') : undefined;
  const font = face ? resolveThemeFont(face, theme) : undefined;
  const size =
    sizeEl?.local === 'buSzPct'
      ? { pct: percentAttr(attr(sizeEl, 'val')) ?? 1 }
      : sizeEl?.local === 'buSzPts'
        ? { pts: intAttr(sizeEl, 'val') ?? 1800 }
        : { pct: 1 };
  const color = colorEl?.local === 'buClr' ? resolveColor(colorChild(colorEl), ctx) : undefined;
  const common = { ...(font ? { font } : {}), size, ...(color ? { color } : {}) };
  if (kind.local === 'buChar') return { type: 'char', char: attr(kind, 'char') ?? '•', ...common };
  return {
    type: 'autonum',
    scheme: attr(kind, 'type') ?? 'arabicPeriod',
    startAt: Math.max(1, Math.min(intAttr(kind, 'startAt') ?? 1, 32_767)),
    ...common,
  };
}

/**
 * A percentage attribute in either form: thousandths of a percent (`150000`) or the schema's
 * `ST_TextBulletSizePercent` pattern (`150%`). Both occur in the wild.
 */
function percentAttr(raw: string | undefined): number | undefined {
  if (!raw) return undefined;
  const m = /^\s*(-?\d+(?:\.\d+)?)\s*(%?)\s*$/.exec(raw);
  if (!m) return undefined;
  const v = Number(m[1]);
  return m[2] ? v / 100 : v / 100_000;
}

const ALIGNS = new Set(['l', 'ctr', 'r', 'just', 'dist', 'justLow', 'thaiDist']);

/** Inputs for resolving one paragraph. */
export interface ParagraphInputs {
  readonly paragraph: Element;
  /** Builds the level chain (lowest first, excluding the paragraph's own pPr). */
  readonly chainFor: (level: number) => readonly Element[];
  readonly ctx: ColorContext;
  readonly theme: Theme;
  readonly fontRef?: FontRef;
  /** Text for `a:fld type="slidenum"`. */
  readonly slideNumber: number;
}

/** Resolve one `a:p`. */
export function resolveParagraph(input: ParagraphInputs): ResolvedParagraph {
  const { paragraph, ctx, theme, fontRef } = input;
  const pPr = child(paragraph, 'a:pPr');
  const level = Math.max(0, Math.min(intAttr(pPr, 'lvl') ?? 0, 8));
  const chain = [...input.chainFor(level), ...(pPr ? [pPr] : [])];
  const runBase: Element[] = [];
  for (const entry of chain) {
    const def = child(entry, 'a:defRPr');
    if (def) runBase.push(def);
  }
  const runs: ResolvedRun[] = [];
  for (const item of elements(paragraph)) {
    if (item.kind === 'run' || item.kind === 'field') {
      const rPr = child(item, 'a:rPr');
      const t = child(item, 'a:t');
      const fieldType = item.kind === 'field' ? attr(item, 'type') : undefined;
      let text = textOf(t);
      if (fieldType === 'slidenum') text = String(input.slideNumber);
      runs.push({
        kind: item.kind === 'run' ? 'text' : 'field',
        nodeId: item.id,
        ...(t ? { textNodeId: t.id } : {}),
        text,
        props: resolveRunProps(rPr ? [...runBase, rPr] : runBase, ctx, theme, fontRef),
        ...(fieldType ? { fieldType } : {}),
      });
    } else if (item.kind === 'break') {
      const rPr = child(item, 'a:rPr');
      runs.push({
        kind: 'break',
        nodeId: item.id,
        text: '\n',
        props: resolveRunProps(rPr ? [...runBase, rPr] : runBase, ctx, theme, fontRef),
      });
    }
  }
  const end = child(paragraph, 'a:endParaRPr');
  const align = lastAttr(chain, 'algn') ?? 'l';
  const tabs: { pos: number; align: string }[] = [];
  const tabList = lastChild(chain, ['tabLst']);
  for (const tab of children(tabList, 'a:tab'))
    tabs.push({ pos: clampEmu(intAttr(tab, 'pos'), 0), align: attr(tab, 'algn') ?? 'l' });
  return {
    nodeId: paragraph.id,
    level,
    align: (ALIGNS.has(align) ? align : 'l') as ResolvedParagraph['align'],
    marL: clampEmu(numberAttr(chain, 'marL'), 0),
    marR: clampEmu(numberAttr(chain, 'marR'), 0),
    indent: clampEmu(numberAttr(chain, 'indent'), 0),
    lineSpacing: spacing(lastChild(chain, ['lnSpc']), { pct: 1 }),
    spaceBefore: spacing(lastChild(chain, ['spcBef']), { pts: 0 }),
    spaceAfter: spacing(lastChild(chain, ['spcAft']), { pts: 0 }),
    bullet: resolveBullet(chain, ctx, theme),
    defaultTab: clampEmu(numberAttr(chain, 'defTabSz'), 914_400) || 914_400,
    tabs,
    rtl: lastAttr(chain, 'rtl') === '1',
    eaLineBreak: lastAttr(chain, 'eaLnBrk') !== '0',
    latinLineBreak: lastAttr(chain, 'latinLnBrk') === '1',
    runs,
    endProps: resolveRunProps(end ? [...runBase, end] : runBase, ctx, theme, fontRef),
  };
}

function numberAttr(chain: readonly Element[], name: string): number | undefined {
  for (let i = chain.length - 1; i >= 0; i -= 1) {
    const v = intAttr(chain[i], name);
    if (v !== undefined) return v;
  }
  return undefined;
}

/** `xsd:boolean` read for bodyPr flags. */
export function flag(el: Element | undefined, name: string): boolean | undefined {
  return boolAttr(el, name);
}
