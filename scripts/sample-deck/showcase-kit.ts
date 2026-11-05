// Small DrawingML helpers for the showcase deck (scripts/sample-deck/showcase.ts). Positions
// and sizes are in inches, line widths and font sizes in points, percentages from 0 to 100.
// Colors are a theme slot (`accent1`, `tx1`, `bg2`, …) or six hex digits.

import { escapeText, shape, type FixtureSlide, type ThemeColors } from '../fixtures/builder.ts';

export const IN = 914_400;
const PT = 12_700;
export const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
export const R_LINK =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

/** Inches to EMU. */
export const emu = (inches: number): number => Math.round(inches * IN);

/** The showcase theme: EigenPal ink and paper, a deep orange, EigenPal green. */
export const SHOWCASE_COLORS: ThemeColors = {
  dk1: '1B1A19',
  lt1: 'FFFFFF',
  dk2: '75716B',
  lt2: 'F4F3F0',
  accent1: 'C43E1C',
  accent2: '2EB52E',
  accent3: 'E3A21A',
  accent4: '2F6DB5',
  accent5: '7B5EA7',
  accent6: '1E8C8C',
  hlink: '2F6DB5',
  folHlink: '7B5EA7',
};

/** A color element, with an optional alpha in percent. */
export function color(c: string, alpha?: number): string {
  const kids = alpha === undefined ? '' : `<a:alpha val="${Math.round(alpha * 1000)}"/>`;
  return /^[0-9A-F]{6}$/.test(c)
    ? `<a:srgbClr val="${c}">${kids}</a:srgbClr>`
    : `<a:schemeClr val="${c}">${kids}</a:schemeClr>`;
}

export const fill = (c: string, alpha?: number): string =>
  `<a:solidFill>${color(c, alpha)}</a:solidFill>`;
export const NO_FILL = '<a:noFill/>';
export const NO_LINE = '<a:ln><a:noFill/></a:ln>';

/** An `a:ln`: width in points, a color, then children after the fill (dash, join, ends). */
export const line = (w: number, c: string, after = '', attrs = ''): string =>
  `<a:ln w="${Math.round(w * PT)}"${attrs}>${fill(c)}${after}</a:ln>`;

/** A gradient stop list from `[position %, color]` pairs. */
export const stops = (...s: [number, string][]): string =>
  `<a:gsLst>${s.map(([pos, c]) => `<a:gs pos="${pos * 1000}">${color(c)}</a:gs>`).join('')}</a:gsLst>`;

/** A hard outer shadow (no blur): distance in points, direction in degrees. */
export const hardShadow = (dist: number, dir: number, c: string, alpha?: number): string =>
  `<a:effectLst><a:outerShdw dist="${Math.round(dist * PT)}" dir="${dir * 60000}" algn="tl" rotWithShape="0">${color(c, alpha)}</a:outerShdw></a:effectLst>`;

export interface RunOptions {
  readonly sz?: number;
  readonly b?: boolean;
  readonly i?: boolean;
  readonly u?: 'sng' | 'dbl';
  readonly strike?: 'sngStrike' | 'dblStrike';
  readonly caps?: boolean;
  /** Character spacing in points. */
  readonly spc?: number;
  readonly color?: string;
  readonly highlight?: string;
  readonly font?: string;
  /** Relationship id of a hyperlink. */
  readonly link?: string;
}

/** One text run. */
export function r(text: string, o: RunOptions = {}): string {
  const attrs = [
    'lang="en-US"',
    o.sz ? `sz="${o.sz * 100}"` : '',
    o.b ? 'b="1"' : '',
    o.i ? 'i="1"' : '',
    o.u ? `u="${o.u}"` : '',
    o.strike ? `strike="${o.strike}"` : '',
    o.caps ? 'cap="all"' : '',
    o.spc === undefined ? '' : `spc="${Math.round(o.spc * 100)}"`,
    'dirty="0"',
  ]
    .filter(Boolean)
    .join(' ');
  const kids =
    (o.color ? fill(o.color) : '') +
    (o.highlight ? `<a:highlight>${color(o.highlight)}</a:highlight>` : '') +
    (o.font ? `<a:latin typeface="${o.font}"/>` : '') +
    (o.link ? `<a:hlinkClick r:id="${o.link}"/>` : '');
  return `<a:r><a:rPr ${attrs}>${kids}</a:rPr><a:t>${escapeText(text)}</a:t></a:r>`;
}

export interface ParagraphOptions {
  readonly algn?: 'l' | 'ctr' | 'r' | 'just';
  readonly lvl?: number;
  /** Line spacing in percent of single. */
  readonly lnSpc?: number;
  /** Space before and after, in points. */
  readonly before?: number;
  readonly after?: number;
  /** The size of the paragraph mark, in points. */
  readonly end?: number;
  /** `a:pPr` children after the spacing (bullets). */
  readonly bullet?: string;
  readonly marL?: number;
  readonly indent?: number;
}

/** One paragraph. */
export function p(runs: string, o: ParagraphOptions = {}): string {
  const attrs =
    (o.marL === undefined ? '' : ` marL="${emu(o.marL)}"`) +
    (o.lvl ? ` lvl="${o.lvl}"` : '') +
    (o.indent === undefined ? '' : ` indent="${emu(o.indent)}"`) +
    (o.algn ? ` algn="${o.algn}"` : '');
  const kids =
    (o.lnSpc ? `<a:lnSpc><a:spcPct val="${o.lnSpc * 1000}"/></a:lnSpc>` : '') +
    (o.before ? `<a:spcBef><a:spcPts val="${o.before * 100}"/></a:spcBef>` : '') +
    (o.after ? `<a:spcAft><a:spcPts val="${o.after * 100}"/></a:spcAft>` : '') +
    (o.bullet ?? '');
  const pPr = attrs || kids ? `<a:pPr${attrs}>${kids}</a:pPr>` : '';
  const end = o.end ? `<a:endParaRPr lang="en-US" sz="${o.end * 100}" dirty="0"/>` : '';
  return `<a:p>${pPr}${runs}${end}</a:p>`;
}

/** One list level for an `a:lstStyle`: margin and hanging indent in inches, size in points. */
export const listLevel = (
  n: number,
  marL: number,
  sz: number,
  bullet: string,
  before = 0
): string =>
  `<a:lvl${n}pPr marL="${emu(marL)}" indent="-342900">` +
  (before ? `<a:spcBef><a:spcPts val="${before * 100}"/></a:spcBef>` : '') +
  `${bullet}<a:defRPr sz="${sz * 100}"/></a:lvl${n}pPr>`;

/** Shape ids, unique in one slide. */
let nextId = 2;
export const resetIds = (): void => {
  nextId = 2;
};
export const newId = (): number => nextId++;

export interface BoxOptions {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly name?: string;
  readonly id?: number;
  readonly prst?: string;
  readonly avLst?: string;
  readonly rot?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  readonly fill?: string;
  readonly line?: string;
  readonly effects?: string;
  readonly text?: string;
  readonly lstStyle?: string;
  readonly anchor?: 't' | 'ctr' | 'b';
  /** Insets in inches: left, top, right, bottom. */
  readonly ins?: readonly [number, number, number, number];
  readonly fit?: 'none' | 'shape' | `shrink ${number}`;
  readonly wrap?: boolean;
  readonly vert?: 'vert' | 'vert270';
  /** Column count and spacing in inches. */
  readonly cols?: readonly [number, number];
  readonly ph?: string;
}

/** A shape (with text when `text` is set). */
export function box(o: BoxOptions): string {
  const [l, t, rt, b] = o.ins ?? [0.1, 0.05, 0.1, 0.05];
  const fit =
    o.fit === 'shape'
      ? '<a:spAutoFit/>'
      : o.fit?.startsWith('shrink')
        ? `<a:normAutofit fontScale="${Number(o.fit.slice(7)) * 1000}"/>`
        : o.ph
          ? '<a:normAutofit/>'
          : '<a:noAutofit/>';
  const bodyPr =
    `<a:bodyPr wrap="${o.wrap === false ? 'none' : 'square'}" lIns="${emu(l)}" tIns="${emu(t)}" rIns="${emu(rt)}" bIns="${emu(b)}"` +
    `${o.anchor ? ` anchor="${o.anchor}"` : ''}${o.vert ? ` vert="${o.vert}"` : ''}` +
    `${o.cols ? ` numCol="${o.cols[0]}" spcCol="${emu(o.cols[1])}"` : ''} rtlCol="0">${fit}</a:bodyPr>`;
  const id = o.id ?? newId();
  return shape({
    id,
    name: o.name ?? `${o.text === undefined ? 'Shape' : 'Text'} ${id}`,
    x: emu(o.x),
    y: emu(o.y),
    cx: emu(o.w),
    cy: emu(o.h),
    prst: o.prst,
    avLst: o.avLst,
    rot: o.rot ? Math.round((((o.rot % 360) + 360) % 360) * 60000) : undefined,
    flipH: o.flipH,
    flipV: o.flipV,
    fill: o.fill ?? (o.text === undefined ? undefined : NO_FILL),
    line: o.line,
    effects: o.effects,
    ...(o.text === undefined ? {} : { text: o.text, bodyPr }),
    lstStyle: o.lstStyle,
    ph: o.ph,
  });
}

/** A round-cornered rectangle; `radius` is the corner adjustment in percent. */
export const rounded = (radius: number): Pick<BoxOptions, 'prst' | 'avLst'> => ({
  prst: 'roundRect',
  avLst: `<a:gd name="adj" fmla="val ${radius * 1000}"/>`,
});

export interface PictureOptions {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly name: string;
  readonly descr: string;
  readonly rId?: string;
  /** Crop in percent: left, top, right, bottom. */
  readonly crop?: readonly [number, number, number, number];
  readonly blip?: string;
  readonly prst?: string;
  readonly line?: string;
  readonly effects?: string;
}

/** A picture. */
export function pic(o: PictureOptions): string {
  const crop = o.crop
    ? `<a:srcRect l="${o.crop[0] * 1000}" t="${o.crop[1] * 1000}" r="${o.crop[2] * 1000}" b="${o.crop[3] * 1000}"/>`
    : '';
  return (
    `<p:pic><p:nvPicPr><p:cNvPr id="${newId()}" name="${escapeText(o.name)}" descr="${escapeText(o.descr)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${o.rId ?? 'rIdImg'}">${o.blip ?? ''}</a:blip>${crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${emu(o.x)}" y="${emu(o.y)}"/><a:ext cx="${emu(o.w)}" cy="${emu(o.h)}"/></a:xfrm>` +
    `<a:prstGeom prst="${o.prst ?? 'rect'}"><a:avLst/></a:prstGeom>${o.line ?? ''}${o.effects ?? ''}</p:spPr></p:pic>`
  );
}

export interface ConnectorOptions {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
  readonly line: string;
  readonly prst?: string;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  readonly name?: string;
  /** Glue: `[shape id, connection site]`. */
  readonly start?: readonly [number, number];
  readonly end?: readonly [number, number];
}

/** A connector (`p:cxnSp`). */
export function cxn(o: ConnectorOptions): string {
  const id = newId();
  const glue =
    (o.start ? `<a:stCxn id="${o.start[0]}" idx="${o.start[1]}"/>` : '') +
    (o.end ? `<a:endCxn id="${o.end[0]}" idx="${o.end[1]}"/>` : '');
  const flips = `${o.flipH ? ' flipH="1"' : ''}${o.flipV ? ' flipV="1"' : ''}`;
  return (
    `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="${o.name ?? `Connector ${id}`}"/><p:cNvCxnSpPr>${glue}</p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>` +
    `<p:spPr><a:xfrm${flips}><a:off x="${emu(o.x)}" y="${emu(o.y)}"/><a:ext cx="${emu(o.w)}" cy="${emu(o.h)}"/></a:xfrm>` +
    `<a:prstGeom prst="${o.prst ?? 'straightConnector1'}"><a:avLst/></a:prstGeom>${o.line}</p:spPr></p:cxnSp>`
  );
}

/** A group whose child coordinates are slide coordinates (chOff = off, chExt = ext). */
export function group(
  o: { x: number; y: number; w: number; h: number; name: string; rot?: number },
  children: (id: number) => string
): string {
  const id = newId();
  const off = `<a:off x="${emu(o.x)}" y="${emu(o.y)}"/><a:ext cx="${emu(o.w)}" cy="${emu(o.h)}"/>`;
  return (
    `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="${o.name}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm${o.rot ? ` rot="${Math.round((((o.rot % 360) + 360) % 360) * 60000)}"` : ''}>${off}` +
    `<a:chOff x="${emu(o.x)}" y="${emu(o.y)}"/><a:chExt cx="${emu(o.w)}" cy="${emu(o.h)}"/></a:xfrm></p:grpSpPr>` +
    `${children(id)}</p:grpSp>`
  );
}

/** The kicker, the title placeholder, and an optional one-line lede of a content slide. */
export function header(kicker: string, title: string, lede?: string): string {
  return (
    box({
      name: 'Kicker',
      x: 0.7,
      y: 0.42,
      w: 9,
      h: 0.3,
      ins: [0, 0, 0, 0],
      text: p(r(kicker, { sz: 12, b: true, caps: true, spc: 2, color: 'accent1' })),
    }) +
    box({
      name: 'Title',
      ph: '<p:ph type="title"/>',
      x: 0.7,
      y: 0.7,
      w: 11.93,
      h: 0.66,
      ins: [0, 0.05, 0, 0.05],
      anchor: 'b',
      text: p(r(title, { sz: 32, b: true })),
    }) +
    (lede
      ? box({
          name: 'Lede',
          x: 0.7,
          y: 1.36,
          w: 11.93,
          h: 0.4,
          ins: [0, 0.05, 0, 0.05],
          text: p(r(lede, { sz: 16, color: 'tx2' })),
        })
      : '')
  );
}

/** A small caps label above a block of content. */
export const label = (x: number, y: number, w: number, text: string, c = 'accent1'): string =>
  box({
    name: `Label ${text}`,
    x,
    y,
    w,
    h: 0.3,
    ins: [0, 0, 0, 0],
    anchor: 'b',
    text: p(r(text, { sz: 12, b: true, caps: true, spc: 1.5, color: c })),
  });

export const IMAGE: [string, string, string] = ['rIdImg', R_IMAGE, '../media/image1.png'];
/**
 * A dark slide: its color map swaps the light and dark theme colors (ECMA-376 Part 1, 19.3.1.7
 * clrMapOvr), so the background is `bg1` and text is `tx1` there. A text box that a user adds
 * on the slide gets the light `tx1`, not the background color, and stays readable.
 */
export const DARK_SLIDE = {
  background:
    '<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>',
  clrMapOvr:
    '<a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>',
} as const;
export const LIGHT_TEXT = 'D9D5CE';
export const FAINT_TEXT = 'A8A39B';

/** A slide built with fresh shape ids. */
export function slide(build: () => Omit<FixtureSlide, 'notes'>, notes: string[]): FixtureSlide {
  resetIds();
  return { ...build(), notes: notes.join('\n') };
}
