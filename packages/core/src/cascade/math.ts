// Office Math in a paragraph (ECMA-376 Part 1, 22.1): an `a14:m` that holds `m:oMathPara` or
// `m:oMath`. This module reads the math into a tree of the constructs the layout draws, with the
// run properties of each math run resolved through the paragraph's chain. A zone with any other
// element is not supported: `isSupportedMath` says no, so `alternateBranch` draws the
// `mc:Fallback` of the shape instead, and the math stays in the file as it is.

import { child, elements, NS, textOf, type Element } from '../store/index.ts';
import type { RunProps } from './types.ts';

/** The Office Math namespace (ECMA-376 Part 1, 22.1). */
export const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

/** The deepest nesting of math constructs that the reader follows. */
const MAX_MATH_DEPTH = 32;
/** The most math elements in one zone. */
const MAX_MATH_NODES = 4_000;
/** The most characters of one math run that draw. */
const MAX_MATH_TEXT = 1_000;

/** A math style (`m:sty`): plain, bold, italic, or bold italic. */
export type MathStyle = 'p' | 'b' | 'i' | 'bi';

/** One math run (`m:r`). */
export interface MathRun {
  readonly kind: 'run';
  readonly text: string;
  readonly props: RunProps;
  /** `m:sty`, when the run states one. */
  readonly sty?: MathStyle;
  /** `m:nor`: normal text, which keeps its characters and the run font's style. */
  readonly normal: boolean;
}

/** A row of math: the content of an argument (`m:e`, `m:num`, …). */
export type MathRow = readonly MathNode[];

/** A math construct. `props` sizes and colors the parts the construct draws itself. */
export type MathNode =
  | MathRun
  | {
      readonly kind: 'frac';
      readonly type: 'bar' | 'skw' | 'lin' | 'noBar';
      readonly num: MathRow;
      readonly den: MathRow;
      readonly props: RunProps;
    }
  | {
      readonly kind: 'script';
      /** `sSup`, `sSub`, `sSubSup`, or `sPre` (scripts before the base). */
      readonly type: 'sup' | 'sub' | 'subSup' | 'pre';
      readonly base: MathRow;
      readonly sub?: MathRow;
      readonly sup?: MathRow;
      readonly props: RunProps;
    }
  | {
      readonly kind: 'rad';
      readonly degree?: MathRow;
      readonly base: MathRow;
      readonly props: RunProps;
    }
  | {
      readonly kind: 'nary';
      readonly chr: string;
      /** `m:limLoc`; absent: beside an integral, under and over any other operator. */
      readonly limLoc?: 'undOvr' | 'subSup';
      readonly sub?: MathRow;
      readonly sup?: MathRow;
      readonly base: MathRow;
      readonly props: RunProps;
    }
  | {
      readonly kind: 'delim';
      readonly begin: string;
      readonly end: string;
      readonly separator: string;
      readonly grow: boolean;
      readonly items: readonly MathRow[];
      readonly props: RunProps;
    }
  | {
      readonly kind: 'matrix';
      readonly rows: readonly (readonly MathRow[])[];
      readonly props: RunProps;
    }
  | {
      readonly kind: 'acc';
      readonly chr: string;
      readonly base: MathRow;
      readonly props: RunProps;
    }
  | {
      readonly kind: 'bar';
      readonly pos: 'top' | 'bot';
      readonly base: MathRow;
      readonly props: RunProps;
    }
  | { readonly kind: 'func'; readonly name: MathRow; readonly base: MathRow }
  | {
      readonly kind: 'limit';
      readonly pos: 'low' | 'upp';
      readonly base: MathRow;
      readonly limit: MathRow;
      readonly props: RunProps;
    }
  | {
      readonly kind: 'groupChr';
      readonly chr: string;
      readonly pos: 'top' | 'bot';
      readonly base: MathRow;
      readonly props: RunProps;
    }
  | { readonly kind: 'box'; readonly base: MathRow }
  | {
      readonly kind: 'borderBox';
      readonly hide: {
        readonly t: boolean;
        readonly b: boolean;
        readonly l: boolean;
        readonly r: boolean;
      };
      readonly base: MathRow;
      readonly props: RunProps;
    }
  | { readonly kind: 'eqArr'; readonly rows: readonly MathRow[] }
  | {
      readonly kind: 'phant';
      readonly show: boolean;
      readonly zeroWidth: boolean;
      readonly zeroAscent: boolean;
      readonly zeroDescent: boolean;
      readonly base: MathRow;
    };

/** One math zone (`a14:m`): display math (`m:oMathPara`) or inline math (`m:oMath`). */
export interface MathZone {
  /** True for `m:oMathPara`: display math, on its own line. */
  readonly display: boolean;
  /** `m:oMathParaPr/m:jc`. */
  readonly justification: 'left' | 'right' | 'center' | 'centerGroup';
  /** Each `m:oMath`; display math draws each on its own line. */
  readonly lines: readonly MathRow[];
  /** The properties of the zone: the first run's, which size its line. */
  readonly props: RunProps;
}

const isMath = (el: Element, local?: string): boolean =>
  el.ns === MATH_NS && (local === undefined || el.local === local);

/** The `m:val` of a property element, or undefined. */
const mval = (el: Element | undefined): string | undefined =>
  el
    ? (el.attrs.find((a) => a.local === 'val' && (a.ns === MATH_NS || a.ns === ''))?.value ??
      undefined)
    : undefined;

/** A math on/off property: present means on, unless `m:val` says off. @internal */
export function onOff(pr: Element | undefined, local: string): boolean | undefined {
  const el = pr && elements(pr).find((c) => isMath(c, local));
  if (!el) return undefined;
  const v = mval(el);
  return v === undefined || v === '1' || v === 'on' || v === 'true';
}

const mchild = (el: Element | undefined, local: string): Element | undefined =>
  el ? elements(el).find((c) => isMath(c, local)) : undefined;

/** The elements of the property part (`m:…Pr`) of each construct, which the layout reads. */
const PROPERTY_CHILDREN: Readonly<Record<string, ReadonlySet<string>>> = {
  oMathParaPr: new Set(['jc']),
  rPr: new Set(['lit', 'nor', 'scr', 'sty', 'brk', 'aln']),
  fPr: new Set(['type', 'ctrlPr']),
  sSupPr: new Set(['ctrlPr']),
  sSubPr: new Set(['ctrlPr']),
  sSubSupPr: new Set(['alnScr', 'ctrlPr']),
  sPrePr: new Set(['ctrlPr']),
  radPr: new Set(['degHide', 'ctrlPr']),
  naryPr: new Set(['chr', 'limLoc', 'grow', 'subHide', 'supHide', 'ctrlPr']),
  dPr: new Set(['begChr', 'sepChr', 'endChr', 'grow', 'shp', 'ctrlPr']),
  mPr: new Set(['baseJc', 'plcHide', 'rSpRule', 'cGpRule', 'rSp', 'cSp', 'cGp', 'mcs', 'ctrlPr']),
  accPr: new Set(['chr', 'ctrlPr']),
  barPr: new Set(['pos', 'ctrlPr']),
  funcPr: new Set(['ctrlPr']),
  limLowPr: new Set(['ctrlPr']),
  limUppPr: new Set(['ctrlPr']),
  groupChrPr: new Set(['chr', 'pos', 'vertJc', 'ctrlPr']),
  boxPr: new Set(['opEmu', 'noBreak', 'diff', 'brk', 'aln', 'ctrlPr']),
  borderBoxPr: new Set([
    'hideTop',
    'hideBot',
    'hideLeft',
    'hideRight',
    'strikeH',
    'strikeV',
    'strikeBLTR',
    'strikeTLBR',
    'ctrlPr',
  ]),
  eqArrPr: new Set(['baseJc', 'maxDist', 'objDist', 'rSpRule', 'rSp', 'ctrlPr']),
  phantPr: new Set(['show', 'zeroWid', 'zeroAsc', 'zeroDesc', 'transp', 'ctrlPr']),
};

/** The arguments of each construct, in schema order (ECMA-376 Part 1, 22.1.2). */
const ARGUMENTS: Readonly<Record<string, readonly string[]>> = {
  f: ['num', 'den'],
  sSup: ['e', 'sup'],
  sSub: ['e', 'sub'],
  sSubSup: ['e', 'sub', 'sup'],
  sPre: ['sub', 'sup', 'e'],
  rad: ['deg', 'e'],
  nary: ['sub', 'sup', 'e'],
  acc: ['e'],
  bar: ['e'],
  func: ['fName', 'e'],
  limLow: ['e', 'lim'],
  limUpp: ['e', 'lim'],
  groupChr: ['e'],
  box: ['e'],
  borderBox: ['e'],
  phant: ['e'],
};

/** Whether `pr` holds only the property elements the layout reads (and `m:ctrlPr`). */
function knownProperties(pr: Element): boolean {
  const allowed = PROPERTY_CHILDREN[pr.local];
  if (!allowed) return false;
  return elements(pr).every((c) => {
    if (!isMath(c) || !allowed.has(c.local)) return false;
    // `m:mcs` lists column properties: `m:mc` with `m:mcPr` of `m:count` and `m:mcJc`.
    if (c.local === 'mcs')
      return elements(c).every(
        (mc) =>
          isMath(mc, 'mc') &&
          elements(mc).every(
            (p) =>
              isMath(p, 'mcPr') && elements(p).every((x) => isMath(x, 'count') || isMath(x, 'mcJc'))
          )
      );
    return true;
  });
}

interface Budget {
  nodes: number;
}

/** Whether every element of a row is a construct the layout draws. */
function rowSupported(row: Element, depth: number, budget: Budget): boolean {
  if (depth > MAX_MATH_DEPTH) return false;
  for (const el of elements(row)) {
    if ((budget.nodes += 1) > MAX_MATH_NODES) return false;
    // `m:argPr` and `m:ctrlPr` of an argument hold no content.
    if (isMath(el, 'argPr') || isMath(el, 'ctrlPr')) continue;
    if (!isMath(el)) return false;
    if (!constructSupported(el, depth, budget)) return false;
  }
  return true;
}

function constructSupported(el: Element, depth: number, budget: Budget): boolean {
  if (el.local === 'r') {
    return elements(el).every(
      (c) =>
        (isMath(c, 'rPr') && knownProperties(c)) ||
        isMath(c, 't') ||
        // The DrawingML run properties of the math run.
        (c.ns === NS.a && c.local === 'rPr')
    );
  }
  if (el.local === 'd' || el.local === 'eqArr' || el.local === 'm') {
    const pr = `${el.local}Pr`;
    const arg = el.local === 'm' ? 'mr' : 'e';
    return elements(el).every((c) => {
      if (isMath(c, pr)) return knownProperties(c);
      if (!isMath(c, arg)) return false;
      if (el.local !== 'm') return rowSupported(c, depth + 1, budget);
      return elements(c).every(
        (cell) => isMath(cell, 'e') && rowSupported(cell, depth + 1, budget)
      );
    });
  }
  const args = ARGUMENTS[el.local];
  if (!args) return false;
  let at = 0;
  for (const c of elements(el)) {
    if (isMath(c, `${el.local}Pr`)) {
      if (!knownProperties(c)) return false;
      continue;
    }
    // Arguments come once each, in schema order.
    while (at < args.length && !isMath(c, args[at]!)) at += 1;
    if (at >= args.length) return false;
    if (!rowSupported(c, depth + 1, budget)) return false;
    at += 1;
  }
  return true;
}

/** The `m:oMath` children of a zone, or undefined when the zone has another shape. */
function zoneLines(zone: Element): { lines: Element[]; para?: Element } | undefined {
  const items = elements(zone);
  if (items.length !== 1) return undefined;
  const top = items[0]!;
  if (isMath(top, 'oMath')) return { lines: [top] };
  if (!isMath(top, 'oMathPara')) return undefined;
  const lines: Element[] = [];
  for (const c of elements(top)) {
    if (isMath(c, 'oMathParaPr')) {
      if (!knownProperties(c)) return undefined;
      continue;
    }
    if (!isMath(c, 'oMath')) return undefined;
    lines.push(c);
  }
  return lines.length > 0 ? { lines, para: top } : undefined;
}

const supportCache = new WeakMap<Element, boolean>();

/**
 * Whether the layout draws an `a14:m` zone: an `m:oMathPara` or `m:oMath` with only the
 * constructs that this module reads, within the depth and size limits. Cached per element:
 * trees are immutable.
 */
export function isSupportedMath(zone: Element): boolean {
  const hit = supportCache.get(zone);
  if (hit !== undefined) return hit;
  const found = zoneLines(zone);
  const budget: Budget = { nodes: 0 };
  const ok = found !== undefined && found.lines.every((line) => rowSupported(line, 0, budget));
  supportCache.set(zone, ok);
  return ok;
}

/** Whether `el` is an `a14:m` math zone in a paragraph. */
export const isMathZone = (el: Element): boolean => el.ns === NS.a14 && el.local === 'm';

/** One character of a property (`m:chr`, `m:begChr`, …): the first code point, or `fallback`. */
function charProp(pr: Element | undefined, local: string, fallback: string): string {
  const el = mchild(pr, local);
  if (!el) return fallback;
  const v = mval(el);
  // A present property with no value is an empty character: nothing draws.
  if (v === undefined || v === '') return '';
  return String.fromCodePoint(v.codePointAt(0) ?? 32);
}

/**
 * Read a supported zone. `propsOf` resolves a math run's `a:rPr` (or none) through the
 * paragraph's chain. Returns undefined for a zone that `isSupportedMath` refuses.
 */
export function readMathZone(
  zone: Element,
  propsOf: (rPr: Element | undefined) => RunProps
): MathZone | undefined {
  if (!isSupportedMath(zone)) return undefined;
  const found = zoneLines(zone)!;
  const base = propsOf(undefined);
  let first: RunProps | undefined;

  /** The properties of a construct's own marks: its `m:ctrlPr/a:rPr`, else the zone's. */
  const ctrlProps = (pr: Element | undefined): RunProps => {
    const rPr = child(mchild(pr, 'ctrlPr'), 'a:rPr');
    return rPr ? propsOf(rPr) : (first ?? base);
  };

  const row = (el: Element | undefined): MathRow => {
    if (!el) return [];
    const out: MathNode[] = [];
    for (const c of elements(el)) {
      const node = construct(c);
      if (node) out.push(node);
    }
    return out;
  };

  const arg = (el: Element, local: string): MathRow | undefined => {
    const a = mchild(el, local);
    return a ? row(a) : undefined;
  };

  const construct = (el: Element): MathNode | undefined => {
    if (!isMath(el)) return undefined;
    const pr = mchild(el, `${el.local}Pr`);
    switch (el.local) {
      case 'r': {
        const rPr = mchild(el, 'rPr');
        const props = propsOf(child(el, 'a:rPr'));
        first ??= props;
        const sty = mval(mchild(rPr, 'sty'));
        const text = elements(el)
          .filter((c) => isMath(c, 't'))
          .map((t) => textOf(t))
          .join('')
          .slice(0, MAX_MATH_TEXT);
        return {
          kind: 'run',
          text,
          props,
          ...(sty === 'p' || sty === 'b' || sty === 'i' || sty === 'bi' ? { sty } : {}),
          normal: onOff(rPr, 'nor') === true,
        };
      }
      case 'f': {
        const t = mval(mchild(pr, 'type'));
        return {
          kind: 'frac',
          type: t === 'skw' || t === 'lin' || t === 'noBar' ? t : 'bar',
          num: arg(el, 'num') ?? [],
          den: arg(el, 'den') ?? [],
          props: ctrlProps(pr),
        };
      }
      case 'sSup':
      case 'sSub':
      case 'sSubSup':
      case 'sPre': {
        const type =
          el.local === 'sSup'
            ? 'sup'
            : el.local === 'sSub'
              ? 'sub'
              : el.local === 'sPre'
                ? 'pre'
                : 'subSup';
        const sub = arg(el, 'sub');
        const sup = arg(el, 'sup');
        return {
          kind: 'script',
          type,
          base: arg(el, 'e') ?? [],
          ...(sub ? { sub } : {}),
          ...(sup ? { sup } : {}),
          props: ctrlProps(pr),
        };
      }
      case 'rad': {
        const degree = arg(el, 'deg');
        const hidden = onOff(pr, 'degHide') === true;
        return {
          kind: 'rad',
          ...(degree && degree.length > 0 && !hidden ? { degree } : {}),
          base: arg(el, 'e') ?? [],
          props: ctrlProps(pr),
        };
      }
      case 'nary': {
        const sub = onOff(pr, 'subHide') ? undefined : arg(el, 'sub');
        const sup = onOff(pr, 'supHide') ? undefined : arg(el, 'sup');
        const limLoc = mval(mchild(pr, 'limLoc'));
        return {
          kind: 'nary',
          chr: charProp(pr, 'chr', '∫'),
          ...(limLoc === 'undOvr' || limLoc === 'subSup' ? { limLoc } : {}),
          ...(sub && sub.length > 0 ? { sub } : {}),
          ...(sup && sup.length > 0 ? { sup } : {}),
          base: arg(el, 'e') ?? [],
          props: ctrlProps(pr),
        };
      }
      case 'd':
        return {
          kind: 'delim',
          begin: charProp(pr, 'begChr', '('),
          end: charProp(pr, 'endChr', ')'),
          separator: charProp(pr, 'sepChr', '|'),
          grow: onOff(pr, 'grow') !== false,
          items: elements(el)
            .filter((c) => isMath(c, 'e'))
            .map((c) => row(c)),
          props: ctrlProps(pr),
        };
      case 'm':
        return {
          kind: 'matrix',
          rows: elements(el)
            .filter((c) => isMath(c, 'mr'))
            .map((r) =>
              elements(r)
                .filter((c) => isMath(c, 'e'))
                .map((c) => row(c))
            ),
          props: ctrlProps(pr),
        };
      case 'acc':
        return {
          kind: 'acc',
          chr: charProp(pr, 'chr', '̂'),
          base: arg(el, 'e') ?? [],
          props: ctrlProps(pr),
        };
      case 'bar':
        return {
          kind: 'bar',
          // ECMA-376 Part 1, 22.1.2.80 pos: the bar is at the bottom unless stated.
          pos: mval(mchild(pr, 'pos')) === 'top' ? 'top' : 'bot',
          base: arg(el, 'e') ?? [],
          props: ctrlProps(pr),
        };
      case 'func':
        return { kind: 'func', name: arg(el, 'fName') ?? [], base: arg(el, 'e') ?? [] };
      case 'limLow':
      case 'limUpp':
        return {
          kind: 'limit',
          pos: el.local === 'limLow' ? 'low' : 'upp',
          base: arg(el, 'e') ?? [],
          limit: arg(el, 'lim') ?? [],
          props: ctrlProps(pr),
        };
      case 'groupChr':
        return {
          kind: 'groupChr',
          chr: charProp(pr, 'chr', '⏟'),
          pos: mval(mchild(pr, 'pos')) === 'top' ? 'top' : 'bot',
          base: arg(el, 'e') ?? [],
          props: ctrlProps(pr),
        };
      case 'box':
        return { kind: 'box', base: arg(el, 'e') ?? [] };
      case 'borderBox':
        return {
          kind: 'borderBox',
          hide: {
            t: onOff(pr, 'hideTop') === true,
            b: onOff(pr, 'hideBot') === true,
            l: onOff(pr, 'hideLeft') === true,
            r: onOff(pr, 'hideRight') === true,
          },
          base: arg(el, 'e') ?? [],
          props: ctrlProps(pr),
        };
      case 'eqArr':
        return {
          kind: 'eqArr',
          rows: elements(el)
            .filter((c) => isMath(c, 'e'))
            .map((c) => row(c)),
        };
      case 'phant':
        return {
          kind: 'phant',
          show: onOff(pr, 'show') !== false,
          zeroWidth: onOff(pr, 'zeroWid') === true,
          zeroAscent: onOff(pr, 'zeroAsc') === true,
          zeroDescent: onOff(pr, 'zeroDesc') === true,
          base: arg(el, 'e') ?? [],
        };
      default:
        return undefined;
    }
  };

  const lines = found.lines.map((line) => row(line));
  const jc = mval(mchild(mchild(found.para, 'oMathParaPr'), 'jc'));
  return {
    display: found.para !== undefined,
    justification: jc === 'left' || jc === 'right' || jc === 'center' ? jc : 'centerGroup',
    lines,
    props: first ?? base,
  };
}

/** The plain text of a zone's runs, for find, the selection pane, and copies as text. */
export function mathPlainText(zone: Element): string {
  let out = '';
  const walk = (el: Element, depth: number): void => {
    if (depth > MAX_MATH_DEPTH || out.length > MAX_MATH_TEXT * 4) return;
    for (const c of elements(el)) {
      if (isMath(c, 't')) out += textOf(c);
      else walk(c, depth + 1);
    }
  };
  walk(zone, 0);
  return out;
}
