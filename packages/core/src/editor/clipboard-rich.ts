// Formatted text on the clipboard, between this editor and other applications. Another
// application's text arrives as HTML: `richFromHtml` reads it into paragraphs of formatted runs,
// with a bounded tokenizer of its own (no DOM, nothing loaded, nothing run), and
// `paragraphSpecs` turns them into DrawingML paragraphs. Copied text leaves as HTML that
// presentation and document applications read (`richHtml`).
//
// What is kept: bold, italic, underline, strikethrough, superscript and subscript, the size, the
// color (a theme color when the HTML names one), the font (a theme font when the HTML names
// one), the alignment, the list level, and the bullet or number. Links, pictures, and tables
// in the HTML are not kept; a table's cells become text.

import { NS, type NodeSpec } from '../store/index.ts';
import type { ClipboardNode } from '../contracts/index.ts';

/** One run of pasted text: its text, or a line break. Sizes in points. @internal */
export interface RichRun {
  readonly text: string;
  readonly br?: true;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strike?: boolean;
  /** Superscript (positive) or subscript (negative), thousandths of a percent of the size. */
  readonly baseline?: number;
  readonly size?: number;
  /** `RRGGBB`. */
  readonly color?: string;
  /** A theme color (`ST_SchemeColorVal`), in place of `color`. */
  readonly scheme?: string;
  readonly font?: string;
  /** A theme font, in place of `font`. */
  readonly fontTheme?: 'major' | 'minor';
}

/** A paragraph's bullet: none, a character, or a numbering scheme. @internal */
export type RichBullet =
  | { readonly kind: 'none' }
  | { readonly kind: 'char'; readonly char: string; readonly font?: string }
  | { readonly kind: 'number'; readonly scheme: string; readonly startAt: number };

/** One pasted paragraph. @internal */
export interface RichParagraph {
  readonly runs: readonly RichRun[];
  /** The list level, 0 to 8. */
  readonly level?: number;
  readonly align?: 'l' | 'ctr' | 'r' | 'just';
  readonly bullet?: RichBullet;
  /** The left margin (`marL`) and the first line's indent (`indent`), EMU. */
  readonly marL?: number;
  readonly indent?: number;
}

/** The longest HTML read; longer HTML is cut there. */
export const MAX_HTML_CHARS = 4 * 1024 * 1024;
const MAX_PARAGRAPHS = 5_000;
const MAX_RUNS = 50_000;
const MAX_DEPTH = 256;

// ---- Reading HTML ----

interface Style {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean;
  readonly strike?: boolean;
  readonly baseline?: number;
  readonly size?: number;
  readonly color?: string;
  readonly scheme?: string;
  readonly font?: string;
  readonly fontTheme?: 'major' | 'minor';
}

const NAMED: Readonly<Record<string, string>> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
  nbsp: '\u00A0',
  ndash: '\u2013',
  mdash: '\u2014',
  hellip: '\u2026',
  lsquo: '\u2018',
  rsquo: '\u2019',
  ldquo: '\u201C',
  rdquo: '\u201D',
  bull: '\u2022',
  middot: '\u00B7',
  copy: '\u00A9',
  reg: '\u00AE',
  trade: '\u2122',
  euro: '\u20AC',
  pound: '\u00A3',
  yen: '\u00A5',
  deg: '\u00B0',
  times: '\u00D7',
  divide: '\u00F7',
  laquo: '\u00AB',
  raquo: '\u00BB',
};

/** Decode character references. Unknown names stay as they are. */
export function decodeEntities(text: string): string {
  return text.replace(/&(#x[0-9a-f]{1,6}|#[0-9]{1,7}|[a-z]{2,8});/gi, (all, ref: string) => {
    if (ref[0] !== '#') return NAMED[ref.toLowerCase()] ?? all;
    const code =
      ref[1] === 'x' || ref[1] === 'X' ? parseInt(ref.slice(2), 16) : Number(ref.slice(1));
    // Control characters, lone surrogates, and out-of-range values are not text.
    if (!(code >= 0x20 && code <= 0x10ffff) || (code >= 0xd800 && code <= 0xdfff)) return '';
    return String.fromCodePoint(code);
  });
}

const BLOCKS = new Set([
  'p',
  'div',
  'li',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'blockquote',
  'pre',
  'tr',
  'dt',
  'dd',
  'section',
  'article',
  'header',
  'footer',
  'address',
  'figcaption',
  'table',
  'ul',
  'ol',
]);
/** Elements whose content is never text on a slide. */
const SKIPPED = new Set([
  'script',
  'style',
  'head',
  'title',
  'template',
  'noscript',
  'svg',
  'math',
  'object',
  'iframe',
  'select',
  'textarea',
  'button',
  'video',
  'audio',
  'canvas',
]);
const VOID = new Set(['br', 'img', 'hr', 'meta', 'link', 'input', 'col', 'wbr', 'source', 'area']);

const COLORS: Readonly<Record<string, string>> = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  blue: '0000FF',
  yellow: 'FFFF00',
  gray: '808080',
  grey: '808080',
  silver: 'C0C0C0',
  maroon: '800000',
  navy: '000080',
  purple: '800080',
  teal: '008080',
  olive: '808000',
  orange: 'FFA500',
  windowtext: '000000',
};

/** Theme color names in HTML (`mso-…-themecolor`) and their `ST_SchemeColorVal`. */
const THEME_COLORS: Readonly<Record<string, string>> = {
  text1: 'tx1',
  text2: 'tx2',
  background1: 'bg1',
  background2: 'bg2',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
  hyperlink: 'hlink',
  followedhyperlink: 'folHlink',
};

/**
 * Numbering schemes by the number that HTML from presentation applications gives them
 * (`mso-special-format:numbullet<n>`), in `ST_TextAutonumberScheme` names.
 */
const NUMBER_SCHEMES = [
  'alphaLcPeriod',
  'alphaUcPeriod',
  'arabicParenR',
  'arabicPeriod',
  'romanLcParenBoth',
  'romanLcParenR',
  'romanLcPeriod',
  'romanUcPeriod',
  'alphaLcParenBoth',
  'alphaLcParenR',
  'alphaUcParenBoth',
  'alphaUcParenR',
  'arabicParenBoth',
  'arabicPlain',
  'romanUcParenBoth',
  'romanUcParenR',
];

const GENERIC_FONTS = new Set([
  'serif',
  'sans-serif',
  'monospace',
  'cursive',
  'fantasy',
  'system-ui',
  'ui-sans-serif',
  'ui-serif',
  'ui-monospace',
  'inherit',
  'initial',
]);

/** `#rgb`, `#rrggbb`, `rgb(r, g, b)`, or a basic color name as `RRGGBB`. */
function cssColor(value: string): string | undefined {
  const v = value.trim().toLowerCase();
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/.exec(v)?.[1];
  if (hex) return (hex.length === 3 ? [...hex].map((c) => c + c).join('') : hex).toUpperCase();
  const rgb =
    /^rgba?\(\s*(\d{1,3})[\s,]+(\d{1,3})[\s,]+(\d{1,3})\s*(?:[,/]\s*([\d.]+%?)\s*)?\)$/.exec(v);
  if (rgb) {
    // A see-through color is not a text color a slide can take.
    if (rgb[4] !== undefined && parseFloat(rgb[4]) === 0) return undefined;
    return [rgb[1], rgb[2], rgb[3]]
      .map((n) => Math.min(255, Number(n)).toString(16).padStart(2, '0'))
      .join('')
      .toUpperCase();
  }
  return COLORS[v];
}

/** A CSS font size in points; `em` and `%` scale the inherited size. */
function cssSize(value: string, inherited: number | undefined): number | undefined {
  const m = /^([\d.]+)(pt|px|em|rem|%)$/.exec(value.trim().toLowerCase());
  if (!m) return undefined;
  const n = parseFloat(m[1]!);
  if (!Number.isFinite(n)) return undefined;
  const pt =
    m[2] === 'pt'
      ? n
      : m[2] === 'px'
        ? n * 0.75
        : m[2] === '%'
          ? inherited !== undefined
            ? (inherited * n) / 100
            : undefined
          : inherited !== undefined
            ? inherited * n
            : m[2] === 'rem'
              ? n * 12
              : undefined;
  return pt !== undefined && pt >= 1 && pt <= 4000 ? Math.round(pt * 100) / 100 : undefined;
}

/** A CSS length in EMU (`in`, `pt`, `px`, `cm`, `mm`); undefined for another unit. */
function cssLength(value: string | undefined): number | undefined {
  const m = /^(-?(?:\d+\.?\d*|\.\d+))(in|pt|px|cm|mm)$/.exec(value?.trim().toLowerCase() ?? '');
  if (!m) return undefined;
  const per = { in: 914_400, pt: 12_700, px: 9_525, cm: 360_000, mm: 36_000 }[m[2] as 'in'];
  const emu = Math.round(parseFloat(m[1]!) * per);
  // The range of `ST_TextMargin` and `ST_TextIndent`.
  return Math.abs(emu) <= 51_206_400 ? emu : undefined;
}

/** The first family of a CSS font list that names a font, or a theme font. */
function cssFont(value: string): { font?: string; fontTheme?: 'major' | 'minor' } {
  for (const raw of value.split(',')) {
    const name = raw
      .trim()
      .replace(/^["']|["']$/g, '')
      .trim();
    if (!name) continue;
    if (/^\+mj-/i.test(name)) return { fontTheme: 'major' };
    if (/^\+mn-/i.test(name)) return { fontTheme: 'minor' };
    if (GENERIC_FONTS.has(name.toLowerCase()) || name.startsWith('-')) continue;
    if (name.length > 64 || /[<>&"\\]/.test(name)) continue;
    return { font: name };
  }
  return {};
}

/** The declarations of a `style` attribute, split at semicolons outside quotes. */
export function cssDeclarations(style: string): Map<string, string> {
  const out = new Map<string, string>();
  let quote = '';
  let start = 0;
  const add = (end: number): void => {
    const decl = style.slice(start, end);
    const colon = decl.indexOf(':');
    if (colon > 0) out.set(decl.slice(0, colon).trim().toLowerCase(), decl.slice(colon + 1).trim());
  };
  for (let i = 0; i < style.length; i += 1) {
    const c = style[i];
    if (quote) {
      if (c === '\\') i += 1;
      else if (c === quote) quote = '';
    } else if (c === '"' || c === "'") quote = c;
    else if (c === ';') {
      add(i);
      start = i + 1;
    }
  }
  add(style.length);
  return out;
}

/** The style an element gives its text. */
function styleOf(tag: string, attrs: Map<string, string>, parent: Style): Style {
  let s: Style = { ...parent };
  if (tag === 'b' || tag === 'strong' || /^h[1-6]$/.test(tag)) s = { ...s, bold: true };
  if (tag === 'i' || tag === 'em' || tag === 'cite') s = { ...s, italic: true };
  if (tag === 'u' || tag === 'ins') s = { ...s, underline: true };
  if (tag === 's' || tag === 'strike' || tag === 'del') s = { ...s, strike: true };
  if (tag === 'sup') s = { ...s, baseline: 30000 };
  if (tag === 'sub') s = { ...s, baseline: -25000 };
  if (tag === 'font') {
    const color = cssColor(attrs.get('color') ?? '');
    if (color) s = { ...s, color, scheme: undefined };
    const face = attrs.get('face');
    if (face) s = { ...s, ...cssFont(face) };
  }
  const css = cssDeclarations(attrs.get('style') ?? '');
  const weight = css.get('font-weight');
  if (weight) s = { ...s, bold: weight === 'bold' || weight === 'bolder' || Number(weight) >= 600 };
  const fontStyle = css.get('font-style');
  if (fontStyle) s = { ...s, italic: fontStyle === 'italic' || fontStyle === 'oblique' };
  const decoration = css.get('text-decoration-line') ?? css.get('text-decoration');
  if (decoration) {
    const d = decoration.toLowerCase();
    if (d.includes('none')) s = { ...s, underline: false, strike: false };
    if (d.includes('underline')) s = { ...s, underline: true };
    if (d.includes('line-through')) s = { ...s, strike: true };
  }
  const align = css.get('vertical-align');
  if (align === 'super') s = { ...s, baseline: 30000 };
  else if (align === 'sub') s = { ...s, baseline: -25000 };
  else if (align === 'baseline') s = { ...s, baseline: undefined };
  const size = css.get('font-size');
  if (size) {
    const pt = cssSize(size, parent.size);
    if (pt !== undefined) s = { ...s, size: pt };
  }
  const family = css.get('font-family');
  if (family) {
    const f = cssFont(family);
    if (f.font || f.fontTheme) s = { ...s, font: f.font, fontTheme: f.fontTheme };
  }
  const theme = css.get('mso-ascii-theme-font')?.toLowerCase();
  if (theme === 'minor-latin') s = { ...s, font: undefined, fontTheme: 'minor' };
  else if (theme === 'major-latin') s = { ...s, font: undefined, fontTheme: 'major' };
  const color = css.get('color');
  if (color) {
    const c = cssColor(color);
    if (c) s = { ...s, color: c, scheme: undefined };
  }
  const themeColor = css.get('mso-style-textfill-fill-themecolor')?.toLowerCase();
  if (themeColor && THEME_COLORS[themeColor]) s = { ...s, scheme: THEME_COLORS[themeColor] };
  return s;
}

/** A bullet marked by `mso-special-format` (`bullet`, or `numbullet<n>,<start>`). */
function specialBullet(css: Map<string, string>, text: string): RichBullet | undefined {
  const raw = css.get('mso-special-format');
  if (!raw) return undefined;
  const v = raw.replace(/["']/g, '').replace(/\\/g, '');
  if (v === 'bullet') {
    const char = [...text.trim()][0];
    if (!char) return undefined;
    const font = cssFont(css.get('font-family') ?? '').font;
    return { kind: 'char', char, ...(font ? { font } : {}) };
  }
  const m = /^numbullet(\d{1,2})(?:,(\d{1,5}))?$/.exec(v);
  const scheme = m ? NUMBER_SCHEMES[Number(m[1])] : undefined;
  if (!scheme) return undefined;
  return { kind: 'number', scheme, startAt: Math.max(1, Math.min(32767, Number(m![2] ?? 1))) };
}

/** Attributes of a start tag, names lowercased. */
function attributesOf(source: string): Map<string, string> {
  const out = new Map<string, string>();
  const re = /([^\s"'<>/=]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/g;
  for (let m = re.exec(source); m; m = re.exec(source))
    out.set(m[1]!.toLowerCase(), decodeEntities(m[2] ?? m[3] ?? m[4] ?? ''));
  return out;
}

/** The part of clipboard HTML between its fragment markers, or all of it. */
function fragmentOf(html: string): string {
  const start = html.indexOf('<!--StartFragment-->');
  const end = html.lastIndexOf('<!--EndFragment-->');
  if (start >= 0 && end > start) return html.slice(start + 20, end);
  const body = html.search(/<body[\s>]/i);
  const open = body < 0 ? -1 : html.indexOf('>', body);
  return open < 0 ? html : html.slice(open + 1);
}

/** How far after a bullet span's start tag its character is looked for. */
const BULLET_LOOKAHEAD = 256;

/** A token of HTML: text, or a start or end tag with its attribute source. */
type HtmlToken =
  | { readonly kind: 'text'; readonly text: string }
  | {
      readonly kind: 'tag';
      readonly close: boolean;
      readonly name: string;
      readonly attrs: string;
      /** The index after the tag's `>`. */
      readonly end: number;
    };

const isLetter = (c: number): boolean => (c >= 65 && c <= 90) || (c >= 97 && c <= 122);
const isNameChar = (c: number): boolean =>
  isLetter(c) || (c >= 48 && c <= 57) || c === 58 || c === 45;

/**
 * The tokens of HTML, in one pass: comments are skipped, and a `<` that does not start a tag is
 * text. A tag or comment that the HTML does not close ends the HTML, as in a browser, so the
 * scan stays linear for any input.
 */
function* htmlTokens(source: string): Generator<HtmlToken> {
  let i = 0;
  let textStart = 0;
  const n = source.length;
  while (i < n) {
    const lt = source.indexOf('<', i);
    if (lt < 0) {
      yield { kind: 'text', text: source.slice(textStart) };
      return;
    }
    if (source.startsWith('<!--', lt)) {
      if (lt > textStart) yield { kind: 'text', text: source.slice(textStart, lt) };
      const end = source.indexOf('-->', lt + 4);
      if (end < 0) return;
      i = end + 3;
      textStart = i;
      continue;
    }
    const close = source.charCodeAt(lt + 1) === 47; // '/'
    const nameStart = lt + (close ? 2 : 1);
    if (!isLetter(source.charCodeAt(nameStart))) {
      // A literal '<' stays with adjacent text, rather than creating a run per character.
      i = lt + 1;
      continue;
    }
    if (lt > textStart) yield { kind: 'text', text: source.slice(textStart, lt) };
    let j = nameStart + 1;
    while (j < n && isNameChar(source.charCodeAt(j))) j += 1;
    const name = source.slice(nameStart, j);
    const attrsStart = j;
    let quote = 0;
    for (; j < n; j += 1) {
      const c = source.charCodeAt(j);
      if (quote) {
        if (c === quote) quote = 0;
      } else if (c === 34 || c === 39) quote = c;
      else if (c === 62) break; // '>'
    }
    if (j >= n) return;
    yield { kind: 'tag', close, name, attrs: source.slice(attrsStart, j), end: j + 1 };
    i = j + 1;
    textStart = i;
  }
  if (textStart < n) yield { kind: 'text', text: source.slice(textStart) };
}

interface ListFrame {
  readonly ordered: boolean;
  /** The first number (`ol/@start`): every item names it, and the numbers follow on. */
  readonly start: number;
}

interface Open {
  readonly tag: string;
  readonly style: Style;
  readonly block: boolean;
}

/** Control characters other than tab, line feed, and carriage return. */
const CONTROL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g;

/**
 * Read clipboard HTML into paragraphs of formatted runs. Null when it holds no text. The HTML is
 * untrusted: it is cut at `MAX_HTML_CHARS`, and the paragraph and run counts are bounded.
 */
export function richFromHtml(html: string): RichParagraph[] | null {
  const cut = html.length > MAX_HTML_CHARS ? html.slice(0, MAX_HTML_CHARS) : html;
  // Control characters are not text: some applications end their HTML with a NUL.
  const source = fragmentOf(cut.replace(CONTROL, ''));
  const paragraphs: (RichParagraph & { fromList?: boolean })[] = [];
  let runs: RichRun[] = [];
  let para: {
    level?: number;
    align?: RichParagraph['align'];
    bullet?: RichBullet;
    marL?: number;
    indent?: number;
    fromList?: boolean;
  } = {};
  let runCount = 0;
  /** The paragraph holds a no-break space: `<p>&nbsp;</p>` is an empty paragraph, kept. */
  let spaced = false;
  let pre = 0;
  const stack: Open[] = [];
  const lists: ListFrame[] = [];
  const style = (): Style => stack[stack.length - 1]?.style ?? {};
  const flush = (force = false): void => {
    if (paragraphs.length >= MAX_PARAGRAPHS) return;
    // Spaces at the paragraph's edges are not shown in HTML.
    const trimmed = trimEdges(runs);
    // A break at a block's end shows no line: `<p><br></p>` is an empty paragraph.
    if (trimmed[trimmed.length - 1]?.br) {
      trimmed.pop();
      force = true;
    }
    if (trimmed.length || force || spaced || para.bullet)
      paragraphs.push({ runs: trimmed, ...para });
    spaced = false;
    runs = [];
    para = {};
  };
  const addText = (raw: string): void => {
    if (runCount >= MAX_RUNS) return;
    let text = decodeEntities(raw);
    text = pre ? text.replace(/\r\n?/g, '\n') : text.replace(/[\t\n\r\f ]+/g, ' ');
    if (text.includes('\u00A0')) spaced = true;
    text = text.replace(/\u00A0/g, ' ');
    if (!text) return;
    // Collapsed spaces: one space between runs.
    const last = runs[runs.length - 1];
    if (!pre && text.startsWith(' ') && (!last || last.br || last.text.endsWith(' ')))
      text = text.slice(1);
    if (!text) return;
    const parts = pre ? text.split('\n') : [text];
    parts.forEach((piece, i) => {
      if (i > 0) flush(true);
      if (piece) {
        runs.push({ text: piece, ...runStyle(style()) });
        runCount += 1;
      }
    });
  };
  let skipping: string | null = null;
  let skipDepth = 0;
  for (const token of htmlTokens(source)) {
    if (token.kind === 'text') {
      if (!skipping) addText(token.text);
      continue;
    }
    const { close, end: tagEnd, attrs: rawAttrs } = token;
    const tag = token.name.toLowerCase().replace(/^\w+:/, '');
    if (skipping) {
      if (tag === skipping) skipDepth += close ? -1 : 1;
      if (skipDepth === 0) skipping = null;
      continue;
    }
    if (close) {
      const at = stack.map((o) => o.tag).lastIndexOf(tag);
      if (at < 0) continue;
      const closed = stack.splice(at);
      if (tag === 'ul' || tag === 'ol') lists.pop();
      if (tag === 'pre') pre = Math.max(0, pre - 1);
      if (closed.some((o) => o.block)) flush();
      continue;
    }
    if (SKIPPED.has(tag)) {
      if (!rawAttrs.trimEnd().endsWith('/')) {
        skipping = tag;
        skipDepth = 1;
      }
      continue;
    }
    const attrs = attributesOf(rawAttrs);
    if (tag === 'br') {
      if (runCount < MAX_RUNS) runs.push({ text: '', br: true, ...runStyle(style()) });
      runCount += 1;
      continue;
    }
    if (tag === 'td' || tag === 'th') {
      if (runs.length) addText(' ');
      continue;
    }
    if (VOID.has(tag)) continue;
    const css = cssDeclarations(attrs.get('style') ?? '');
    // A bullet span holds the bullet's own character: it marks the paragraph, it is not text.
    if (tag === 'span' && css.has('mso-special-format')) {
      // The bullet's character is near: a bounded look, so many bullet spans stay linear.
      const near = source.slice(tagEnd, tagEnd + BULLET_LOOKAHEAD);
      const end = near.indexOf('</span>');
      // The character is the text before `</span>`, after any tag inside the span.
      const head = end < 0 ? '' : near.slice(0, end);
      const inner = decodeEntities(head.slice(head.lastIndexOf('>') + 1));
      const bullet = specialBullet(css, inner);
      if (bullet) para = { ...para, bullet };
      skipping = 'span';
      skipDepth = 1;
      continue;
    }
    const block = BLOCKS.has(tag);
    if (block) {
      if (runs.length) flush();
      if (tag === 'ul' || tag === 'ol')
        lists.push({
          ordered: tag === 'ol',
          start: Math.max(1, Math.min(32767, Math.floor(Number(attrs.get('start') ?? 1)) || 1)),
        });
      if (tag === 'pre') pre += 1;
      const level = /(?:^|\s)O([1-8])(?:\s|$)/.exec(attrs.get('class') ?? '')?.[1];
      if (level) para = { ...para, level: Number(level) };
      if (tag === 'li') {
        const list = lists[lists.length - 1];
        const depth = Math.max(0, Math.min(8, lists.length - 1));
        if (list) {
          para = {
            ...para,
            fromList: true,
            level: para.level ?? depth,
            bullet: list.ordered
              ? { kind: 'number', scheme: 'arabicPeriod', startAt: list.start }
              : { kind: 'char', char: '\u2022' },
          };
        }
      }
      const align = css.get('text-align')?.toLowerCase();
      if (align === 'center') para = { ...para, align: 'ctr' };
      else if (align === 'right') para = { ...para, align: 'r' };
      else if (align === 'justify') para = { ...para, align: 'just' };
      else if (align === 'left') para = { ...para, align: 'l' };
      const marL = cssLength(css.get('margin-left'));
      const indent = cssLength(css.get('text-indent'));
      if (marL !== undefined && marL >= 0) para = { ...para, marL };
      if (indent !== undefined) para = { ...para, indent };
    }
    if (stack.length >= MAX_DEPTH) continue;
    stack.push({ tag, style: styleOf(tag, attrs, style()), block });
  }
  flush();
  // HTML that marks its bullets (`mso-special-format`) marks every one: a paragraph without a
  // mark has no bullet, also where the slide's list style gives one.
  const marked = paragraphs.some((p) => p.bullet && p.bullet.kind !== 'none' && !p.fromList);
  const out = paragraphs.map(({ fromList, ...p }): RichParagraph =>
    marked && !p.bullet && !fromList ? { ...p, bullet: { kind: 'none' } } : p
  );
  return out.some((p) => p.runs.length || p.bullet) ? out : null;
}

/** A style's run fields. */
function runStyle(s: Style): Omit<RichRun, 'text' | 'br'> {
  const out: { -readonly [K in keyof RichRun]?: RichRun[K] } = {};
  if (s.bold) out.bold = true;
  if (s.italic) out.italic = true;
  if (s.underline) out.underline = true;
  if (s.strike) out.strike = true;
  if (s.baseline) out.baseline = s.baseline;
  if (s.size !== undefined) out.size = s.size;
  if (s.scheme) out.scheme = s.scheme;
  else if (s.color) out.color = s.color;
  if (s.fontTheme) out.fontTheme = s.fontTheme;
  else if (s.font) out.font = s.font;
  return out;
}

/** Runs without the spaces at the paragraph's start and end, and without empty runs. */
function trimEdges(runs: readonly RichRun[]): RichRun[] {
  const out = runs.map((r) => ({ ...r }));
  const first = out.findIndex((r) => !r.br);
  if (first >= 0) out[first] = { ...out[first]!, text: out[first]!.text.replace(/^ +/, '') };
  for (let i = out.length - 1; i >= 0; i -= 1) {
    if (out[i]!.br) continue;
    out[i] = { ...out[i]!, text: out[i]!.text.replace(/ +$/, '') };
    break;
  }
  return out.filter((r) => r.br || r.text);
}

/** Paragraphs of plain text: a line feed starts a paragraph, U+000B is a line break. */
export function richFromText(text: string): RichParagraph[] {
  return text
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .slice(0, MAX_PARAGRAPHS)
    .map((line) => ({
      runs: line
        .split('\u000B')
        .flatMap((piece, i): RichRun[] => [
          ...(i > 0 ? [{ text: '', br: true as const }] : []),
          ...(piece ? [{ text: piece.replace(/\t/g, ' ') }] : []),
        ]),
    }));
}

// ---- DrawingML ----

const a = (
  local: string,
  attrs: Record<string, string> = {},
  children: NodeSpec['children'] = []
): NodeSpec => ({
  ns: NS.a,
  local,
  prefix: 'a',
  attrs: Object.entries(attrs).map(([k, value]) => ({ ns: '', local: k, value })),
  nsDecls: [],
  children,
});

/** A run's `a:rPr` (or `a:endParaRPr`, with `local`). */
function runProps(r: Omit<RichRun, 'text'>, local = 'rPr'): NodeSpec {
  const attrs: Record<string, string> = { lang: 'en-US' };
  if (r.size !== undefined)
    attrs.sz = String(Math.round(Math.max(1, Math.min(4000, r.size)) * 100));
  if (r.bold) attrs.b = '1';
  if (r.italic) attrs.i = '1';
  if (r.underline) attrs.u = 'sng';
  if (r.strike) attrs.strike = 'sngStrike';
  if (r.baseline) attrs.baseline = String(r.baseline);
  const kids: NodeSpec[] = [];
  if (r.scheme) kids.push(a('solidFill', {}, [a('schemeClr', { val: r.scheme })]));
  else if (r.color) kids.push(a('solidFill', {}, [a('srgbClr', { val: r.color })]));
  if (r.fontTheme)
    kids.push(a('latin', { typeface: r.fontTheme === 'major' ? '+mj-lt' : '+mn-lt' }));
  else if (r.font) kids.push(a('latin', { typeface: r.font }));
  return a(local, attrs, kids);
}

/** DrawingML paragraphs (`a:p`) for pasted paragraphs. @internal */
export function paragraphSpecs(paragraphs: readonly RichParagraph[]): NodeSpec[] {
  return paragraphs.map((p) => {
    const pAttrs: Record<string, string> = {};
    if (p.level) pAttrs.lvl = String(Math.max(0, Math.min(8, p.level)));
    if (p.marL !== undefined) pAttrs.marL = String(p.marL);
    if (p.indent !== undefined) pAttrs.indent = String(p.indent);
    if (p.align) pAttrs.algn = p.align;
    const bullet: NodeSpec[] =
      p.bullet?.kind === 'none'
        ? [a('buNone')]
        : p.bullet?.kind === 'char'
          ? [
              ...(p.bullet.font ? [a('buFont', { typeface: p.bullet.font })] : []),
              a('buChar', { char: p.bullet.char }),
            ]
          : p.bullet?.kind === 'number'
            ? [
                a('buAutoNum', {
                  type: p.bullet.scheme,
                  ...(p.bullet.startAt > 1 ? { startAt: String(p.bullet.startAt) } : {}),
                }),
              ]
            : [];
    const kids: NodeSpec[] = [];
    if (Object.keys(pAttrs).length || bullet.length) kids.push(a('pPr', pAttrs, bullet));
    for (const r of p.runs)
      kids.push(
        r.br
          ? a('br', {}, [runProps(r)])
          : a('r', {}, [runProps(r), a('t', {}, [{ text: r.text }])])
      );
    const last = [...p.runs].reverse().find((r) => !r.br) ?? p.runs[p.runs.length - 1];
    kids.push(runProps(last ?? {}, 'endParaRPr'));
    return a('p', {}, kids);
  });
}

// ---- Writing HTML ----

const escapeHtml = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/**
 * A `font-family` declaration for a font name, or '' for a name with characters a font name
 * does not need (the name is from the deck: untrusted).
 */
const fontFamilyCss = (name: string): string =>
  /^[\p{L}\p{N} _.-]{1,64}$/u.test(name) ? "font-family:'" + name + "'" : '';

/** HTML for paragraphs of formatted text, as presentation and document applications read it. @internal */
export function richHtml(paragraphs: readonly RichParagraph[]): string {
  const body = paragraphs
    .map((p) => {
      const blockCss = [
        `margin:0`,
        `text-align:${{ l: 'left', ctr: 'center', r: 'right', just: 'justify' }[p.align ?? 'l']}`,
      ];
      if (p.level) blockCss.push(`margin-left:${p.level * 0.5}in`);
      const cls = p.level ? ` class=O${p.level}` : '';
      const bullet =
        p.bullet?.kind === 'char'
          ? `<span style="mso-special-format:bullet${p.bullet.font && fontFamilyCss(p.bullet.font) ? ';' + fontFamilyCss(p.bullet.font) : ''}">${escapeHtml(p.bullet.char)}</span> `
          : p.bullet?.kind === 'number'
            ? `<span style="mso-special-format:numbullet${Math.max(0, NUMBER_SCHEMES.indexOf(p.bullet.scheme))}\\,${p.bullet.startAt}">${p.bullet.startAt}.</span> `
            : '';
      const runs = p.runs
        .map((r) => {
          if (r.br) return '<br>';
          const css: string[] = [];
          if (r.size !== undefined) css.push(`font-size:${r.size}pt`);
          if (r.font && fontFamilyCss(r.font)) css.push(fontFamilyCss(r.font));
          if (r.color) css.push(`color:#${r.color}`);
          if (r.bold) css.push('font-weight:bold');
          if (r.italic) css.push('font-style:italic');
          const deco = [r.underline ? 'underline' : '', r.strike ? 'line-through' : ''].filter(
            Boolean
          );
          if (deco.length) css.push(`text-decoration:${deco.join(' ')}`);
          if (r.baseline) css.push(`vertical-align:${r.baseline > 0 ? 'super' : 'sub'}`);
          return `<span style="${escapeHtml(css.join(';'))}">${escapeHtml(r.text)}</span>`;
        })
        .join('');
      const tag = p.bullet && p.bullet.kind !== 'none' ? 'div' : 'p';
      // A block's last break shows no line in HTML: a paragraph that ends in a break, or is
      // empty, takes one more.
      const end = !runs || p.runs[p.runs.length - 1]?.br ? '<br>' : '';
      return `<${tag}${cls} style="${blockCss.join(';')}">${bullet}${runs}${end}</${tag}>`;
    })
    .join('');
  return `<meta charset="utf-8"><!--StartFragment-->${body}<!--EndFragment-->`;
}

/**
 * DrawingML paragraphs for the `pasteText` command, from clipboard HTML or plain text: HTML
 * keeps its formatting as a paste from another application does. Use it to paste text from your
 * own code. HTML is read up to 4 MB and 5,000 paragraphs; nothing in it runs or loads. Returns
 * an empty array when there is no text. @public
 */
export function textClipboardParagraphs(source: {
  readonly html?: string;
  readonly text?: string;
}): ClipboardNode[] {
  const rich = source.html ? richFromHtml(source.html) : null;
  if (rich) return paragraphSpecs(rich);
  return source.text ? paragraphSpecs(richFromText(source.text)) : [];
}
