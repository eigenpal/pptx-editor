// The linear form of an equation in a file (`a14:m`), for Edit Equation. Only the constructs
// that the linear syntax of `equation-linear.ts` writes back the same way have a linear form; an
// equation with any other construct, run property, or character has none, and stays as it is.

import { elements, textOf, type Element } from '../../store/index.ts';
import { MATH_NS, onOff } from '../../cascade/math.ts';

const isM = (el: Element | undefined, local?: string): el is Element =>
  !!el && el.ns === MATH_NS && (local === undefined || el.local === local);
const kid = (el: Element | undefined, local: string): Element | undefined =>
  el ? elements(el).find((c) => isM(c, local)) : undefined;
const val = (el: Element | undefined): string | undefined =>
  el?.attrs.find((a) => a.local === 'val')?.value;

const SYMBOL_NAMES: Readonly<Record<string, string>> = {
  '∑': 'sum',
  '∏': 'prod',
  '∫': 'int',
  '∮': 'oint',
};
const ACCENT_NAMES: Readonly<Record<string, string>> = {
  '̂': 'hat',
  '̇': 'dot',
  '̈': 'ddot',
  '⃗': 'vec',
  '̃': 'tilde',
};
const FUNCTION_NAMES = new Set([
  'sin',
  'cos',
  'tan',
  'sec',
  'csc',
  'cot',
  'sinh',
  'cosh',
  'tanh',
  'log',
  'ln',
  'exp',
  'lim',
  'max',
  'min',
  'det',
  'gcd',
  'arg',
]);

/** Characters that mean syntax in the linear form. */
const SYNTAX = /[\\/^_{}()[\]&@\s]/u;

class NoLinearForm extends Error {}

/** The linear form of an equation zone, or undefined when it has none. */
export function linearOfMath(zone: Element): string | undefined {
  const top = elements(zone);
  if (top.length !== 1) return undefined;
  let lines: Element[];
  if (isM(top[0], 'oMath')) lines = [top[0]];
  else if (isM(top[0], 'oMathPara')) lines = elements(top[0]).filter((c) => !isM(c, 'oMathParaPr'));
  else return undefined;
  if (lines.length !== 1 || !isM(lines[0], 'oMath')) return undefined;
  try {
    return row(lines[0]).trim();
  } catch (e) {
    if (e instanceof NoLinearForm) return undefined;
    throw e;
  }
}

const fail = (): never => {
  throw new NoLinearForm();
};

/** The content elements of an argument, without its `m:argPr` and `m:ctrlPr`. */
const content = (el: Element | undefined): Element[] =>
  el ? elements(el).filter((c) => !isM(c, 'argPr') && !isM(c, 'ctrlPr')) : [];

function row(el: Element | undefined, depth = 0): string {
  if (depth > 24) fail();
  return content(el)
    .map((c) => node(c, depth + 1))
    .join('');
}

/** An argument as an operand: parentheses unless it is one run of letters or digits. */
function wrap(el: Element | undefined, depth: number): string {
  const kids = content(el);
  const text = row(el, depth);
  if (kids.length === 1 && isM(kids[0], 'r') && /^([0-9.]+|\p{L})$/u.test(text)) return text;
  return `(${text})`;
}

function runText(r: Element): string {
  const rPr = kid(r, 'rPr');
  // Normal text and literal runs have no linear form.
  if (
    rPr &&
    elements(rPr).some((c) => isM(c, 'nor') || isM(c, 'lit') || isM(c, 'scr') || isM(c, 'brk'))
  )
    fail();
  const sty = val(kid(rPr, 'sty'));
  const text = elements(r)
    .filter((c) => isM(c, 't'))
    .map((t) => textOf(t))
    .join('');
  if (!text || SYNTAX.test(text)) fail();
  // A plain run of letters that is not a function name would read back as italic letters.
  if (sty === 'p' && /\p{L}/u.test(text) && !FUNCTION_NAMES.has(text)) fail();
  if (sty && sty !== 'p' && sty !== 'i') fail();
  // Letters spell a function name only inside `m:fName`.
  if (sty !== 'p' && FUNCTION_NAMES.has(text)) fail();
  return text.replace(/−/g, '-');
}

function node(el: Element, depth: number): string {
  if (!isM(el)) fail();
  switch (el.local) {
    case 'r':
      return runText(el);
    case 'f': {
      if (
        val(kid(kid(el, 'fPr'), 'type')) !== undefined &&
        val(kid(kid(el, 'fPr'), 'type')) !== 'bar'
      )
        fail();
      return `${wrap(kid(el, 'num'), depth)}/${wrap(kid(el, 'den'), depth)}`;
    }
    case 'sSup':
    case 'sSub':
    case 'sSubSup': {
      const base = content(kid(el, 'e'));
      const b =
        base.length === 1 && (isM(base[0], 'r') || isM(base[0], 'd'))
          ? row(kid(el, 'e'), depth)
          : `{${row(kid(el, 'e'), depth)}}`;
      if (isM(base[0], 'r') && [...b].length !== 1 && !/^[0-9.]+$/u.test(b)) fail();
      const sub = kid(el, 'sub') ? `_${wrap(kid(el, 'sub'), depth)}` : '';
      const sup = kid(el, 'sup') ? `^${wrap(kid(el, 'sup'), depth)}` : '';
      return `${b}${sub}${sup}`;
    }
    case 'rad': {
      // The layout's reading of the flag. A hidden degree with content has no linear form:
      // writing it back would show it, and leaving it out would lose it.
      const hidden = onOff(kid(el, 'radPr'), 'degHide') === true;
      const deg = content(kid(el, 'deg'));
      if (hidden && deg.length > 0) fail();
      const degree = deg.length > 0 ? `[${row(kid(el, 'deg'), depth)}]` : '';
      return `\\sqrt${degree}(${row(kid(el, 'e'), depth)})`;
    }
    case 'd': {
      const pr = kid(el, 'dPr');
      const beg = kid(pr, 'begChr') ? val(kid(pr, 'begChr')) : '(';
      const end = kid(pr, 'endChr') ? val(kid(pr, 'endChr')) : ')';
      const items = elements(el).filter((c) => isM(c, 'e'));
      if (items.length !== 1) fail();
      if (!((beg === '(' && end === ')') || (beg === '[' && end === ']'))) fail();
      return `${beg}${row(items[0], depth)}${end}`;
    }
    case 'nary': {
      const pr = kid(el, 'naryPr');
      const chr = kid(pr, 'chr') ? val(kid(pr, 'chr')) : '∫';
      const name = chr ? SYMBOL_NAMES[chr] : undefined;
      if (!name) fail();
      // As for a radical: a hidden limit with content has no linear form.
      const hidden = (local: string): boolean => {
        const off = onOff(pr, local) === true;
        if (off && content(kid(el, local === 'subHide' ? 'sub' : 'sup')).length) fail();
        return off;
      };
      const sub =
        !hidden('subHide') && content(kid(el, 'sub')).length
          ? `_${wrap(kid(el, 'sub'), depth)}`
          : '';
      const sup =
        !hidden('supHide') && content(kid(el, 'sup')).length
          ? `^${wrap(kid(el, 'sup'), depth)}`
          : '';
      return `\\${name}${sub}${sup} ${wrap(kid(el, 'e'), depth)}`;
    }
    case 'func': {
      const name = content(kid(el, 'fName'));
      if (name.length !== 1) fail();
      let head: string;
      if (isM(name[0], 'r')) head = funcName(name[0]!);
      else if (isM(name[0], 'limLow')) {
        const base = content(kid(name[0], 'e'));
        if (base.length !== 1 || !isM(base[0], 'r')) fail();
        head = `${funcName(base[0]!)}_${wrap(kid(name[0], 'lim'), depth)}`;
      } else fail();
      return `${head!} ${wrap(kid(el, 'e'), depth)} `;
    }
    case 'acc': {
      const chr = kid(kid(el, 'accPr'), 'chr') ? val(kid(kid(el, 'accPr'), 'chr')) : '̂';
      const name = chr ? ACCENT_NAMES[chr] : undefined;
      if (!name) fail();
      return `\\${name}(${row(kid(el, 'e'), depth)})`;
    }
    case 'bar':
      if (val(kid(kid(el, 'barPr'), 'pos')) !== 'top') fail();
      return `\\bar(${row(kid(el, 'e'), depth)})`;
    case 'm': {
      const rows = elements(el)
        .filter((c) => isM(c, 'mr'))
        .map((r) =>
          elements(r)
            .filter((c) => isM(c, 'e'))
            .map((c) => row(c, depth))
            .join('&')
        );
      return `\\matrix(${rows.join('@')})`;
    }
    default:
      return fail();
  }
}

function funcName(r: Element): string {
  const text = elements(r)
    .filter((c) => isM(c, 't'))
    .map((t) => textOf(t))
    .join('');
  if (!FUNCTION_NAMES.has(text)) fail();
  return text;
}
