// The linear form of an equation: a short text syntax that Insert Equation and Edit Equation
// read and write, and its Office Math markup (ECMA-376 Part 1, 22.1). DOM-free.
//
// The syntax covers a subset of the common linear math forms:
//
// - `a/b` is a fraction. Its operands are the atoms next to the slash; parentheses around an
//   operand group it and do not draw: `(a+b)/(2c)`.
// - `x^2`, `x_i`, and `x_i^2` are scripts. Parentheses around a script group it: `e^(i\pi)`.
// - `\sqrt(x)` and `\sqrt[3](x)` are radicals.
// - `\sum`, `\prod`, `\int`, and `\oint` take limits with `_` and `^`, then one term.
// - `sin`, `cos`, `tan`, `log`, `ln`, `exp`, `lim`, `max`, and `min` are function names;
//   `lim_(n\to\infty)` puts the limit under the name.
// - `\hat(x)`, `\dot(x)`, `\vec(x)`, `\tilde(x)`, and `\bar(x)` are accents and bars.
// - `\matrix(a&b@c&d)` is a matrix: `&` separates cells, `@` rows.
// - `\alpha` and the other Greek letter names, and `\pm`, `\times`, `\le`, `\ge`, `\ne`,
//   `\to`, `\infty`, and the other symbol names below, are characters.
// - Parentheses and brackets that are not operands draw as delimiters that grow. Braces group
//   and do not draw. Spaces separate names and do not draw.

import { CommandRefusal } from './op-batch.ts';

/** The longest linear form that the parser reads. */
export const MAX_LINEAR_LENGTH = 1_000;
/** The deepest nesting of groups and constructs that the parser reads. */
const MAX_LINEAR_DEPTH = 24;

/** Character names: Greek letters, then symbols. */
const NAMES: Readonly<Record<string, string>> = {
  alpha: 'α',
  beta: 'β',
  gamma: 'γ',
  delta: 'δ',
  epsilon: 'ε',
  varepsilon: 'ε',
  zeta: 'ζ',
  eta: 'η',
  theta: 'θ',
  iota: 'ι',
  kappa: 'κ',
  lambda: 'λ',
  mu: 'μ',
  nu: 'ν',
  xi: 'ξ',
  pi: 'π',
  rho: 'ρ',
  sigma: 'σ',
  tau: 'τ',
  upsilon: 'υ',
  phi: 'φ',
  varphi: 'φ',
  chi: 'χ',
  psi: 'ψ',
  omega: 'ω',
  Gamma: 'Γ',
  Delta: 'Δ',
  Theta: 'Θ',
  Lambda: 'Λ',
  Xi: 'Ξ',
  Pi: 'Π',
  Sigma: 'Σ',
  Phi: 'Φ',
  Psi: 'Ψ',
  Omega: 'Ω',
  pm: '±',
  mp: '∓',
  times: '×',
  div: '÷',
  cdot: '⋅',
  le: '≤',
  leq: '≤',
  ge: '≥',
  geq: '≥',
  ne: '≠',
  neq: '≠',
  approx: '≈',
  equiv: '≡',
  sim: '∼',
  propto: '∝',
  to: '→',
  rightarrow: '→',
  leftarrow: '←',
  Rightarrow: '⇒',
  Leftrightarrow: '⇔',
  infty: '∞',
  partial: '∂',
  nabla: '∇',
  in: '∈',
  notin: '∉',
  subset: '⊂',
  supset: '⊃',
  cup: '∪',
  cap: '∩',
  forall: '∀',
  exists: '∃',
  degree: '°',
  cdots: '⋯',
  ldots: '…',
};

const NARY: Readonly<Record<string, string>> = { sum: '∑', prod: '∏', int: '∫', oint: '∮' };
const ACCENTS: Readonly<Record<string, string>> = {
  hat: '̂',
  dot: '̇',
  ddot: '̈',
  vec: '⃗',
  tilde: '̃',
};
const FUNCTIONS = new Set([
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
/** Function names whose `_` limit goes under the name. */
const LIMIT_FUNCTIONS = new Set(['lim', 'max', 'min']);

/** A parsed linear form. */
export type LinearNode =
  | { readonly t: 'run'; readonly text: string; readonly kind: 'letter' | 'number' | 'op' | 'name' }
  | { readonly t: 'frac'; readonly num: LinearRow; readonly den: LinearRow }
  | {
      readonly t: 'script';
      readonly base: LinearRow;
      readonly sub?: LinearRow;
      readonly sup?: LinearRow;
    }
  | { readonly t: 'rad'; readonly degree?: LinearRow; readonly body: LinearRow }
  | { readonly t: 'delim'; readonly open: string; readonly close: string; readonly body: LinearRow }
  | {
      readonly t: 'nary';
      readonly chr: string;
      readonly sub?: LinearRow;
      readonly sup?: LinearRow;
      readonly body: LinearRow;
    }
  | { readonly t: 'func'; readonly name: LinearRow; readonly body: LinearRow }
  | { readonly t: 'limLow'; readonly base: LinearRow; readonly limit: LinearRow }
  | { readonly t: 'acc'; readonly chr: string; readonly body: LinearRow }
  | { readonly t: 'bar'; readonly body: LinearRow }
  | { readonly t: 'matrix'; readonly rows: readonly (readonly LinearRow[])[] };

export type LinearRow = readonly LinearNode[];

type Token =
  | { readonly k: 'letter' | 'number' | 'op'; readonly v: string }
  | { readonly k: 'cmd'; readonly v: string }
  | { readonly k: 'space' };

function tokenize(src: string): Token[] {
  const out: Token[] = [];
  const chars = [...src];
  for (let i = 0; i < chars.length;) {
    const ch = chars[i]!;
    if (/\s/u.test(ch)) {
      if (out[out.length - 1]?.k !== 'space') out.push({ k: 'space' });
      i += 1;
    } else if (ch === '\\') {
      let j = i + 1;
      while (j < chars.length && /[A-Za-z]/.test(chars[j]!)) j += 1;
      const name = chars.slice(i + 1, j).join('');
      if (!name) throw refuse(`a backslash with no name`);
      out.push({ k: 'cmd', v: name });
      i = j;
    } else if (/[0-9]/.test(ch)) {
      let j = i + 1;
      while (j < chars.length && /[0-9.]/.test(chars[j]!)) j += 1;
      out.push({ k: 'number', v: chars.slice(i, j).join('') });
      i = j;
    } else if (/\p{L}/u.test(ch)) {
      out.push({ k: 'letter', v: ch });
      i += 1;
    } else {
      out.push({ k: 'op', v: ch === '-' ? '−' : ch === '*' ? '∗' : ch });
      i += 1;
    }
  }
  return out;
}

const refuse = (what: string): CommandRefusal =>
  new CommandRefusal('errors.unsupported', `equation: ${what}`);

/** The node list of a group: its content when it is a group, else the node itself. */
function operand(node: LinearNode | LinearRow): LinearRow {
  if (Array.isArray(node)) return node as LinearRow;
  const n = node as LinearNode;
  if (n.t === 'delim' && n.open === '(' && n.close === ')') return n.body;
  return [n];
}

class Parser {
  private at = 0;
  private depth = 0;
  constructor(private readonly tokens: readonly Token[]) {}

  private peek(skipSpace = true): Token | undefined {
    let i = this.at;
    while (skipSpace && this.tokens[i]?.k === 'space') i += 1;
    return this.tokens[i];
  }

  private next(): Token | undefined {
    while (this.tokens[this.at]?.k === 'space') this.at += 1;
    return this.tokens[this.at++];
  }

  private isOp(v: string, skipSpace = true): boolean {
    const t = this.peek(skipSpace);
    return t?.k === 'op' && t.v === v;
  }

  private expect(v: string): void {
    const t = this.next();
    if (t?.k !== 'op' || t.v !== v) throw refuse(`expected ${v}`);
  }

  private enter(): void {
    if ((this.depth += 1) > MAX_LINEAR_DEPTH) throw refuse('too deep');
  }

  /** A row until a closing character, `&`, `@`, or the end. */
  row(stops: ReadonlySet<string>): LinearNode[] {
    const out: LinearNode[] = [];
    for (;;) {
      const t = this.peek();
      if (!t || (t.k === 'op' && stops.has(t.v))) return out;
      let term = this.term();
      // `/` binds the terms next to it into a fraction.
      while (this.isOp('/')) {
        this.next();
        const den = this.term();
        term = { t: 'frac', num: operand(term), den: operand(den) };
      }
      out.push(term);
    }
  }

  /** An atom with its scripts. */
  term(): LinearNode {
    const atom = this.atom();
    return this.scripts(atom);
  }

  private scripts(base: LinearNode, scriptsOnly = false): LinearNode {
    let sub: LinearRow | undefined;
    let sup: LinearRow | undefined;
    for (;;) {
      // Scripts attach only with no space before them.
      if (this.isOp('_', false) && !sub) {
        this.next();
        sub = operand(this.atom());
      } else if (this.isOp('^', false) && !sup) {
        this.next();
        sup = operand(this.atom());
      } else break;
    }
    if (!sub && !sup) return base;
    if (scriptsOnly)
      return { t: 'script', base: [base], ...(sub ? { sub } : {}), ...(sup ? { sup } : {}) };
    const b = base.t === 'delim' && base.open === '{' ? base.body : [base];
    return { t: 'script', base: b, ...(sub ? { sub } : {}), ...(sup ? { sup } : {}) };
  }

  private group(open: string, close: string): LinearNode {
    this.enter();
    const body = this.row(new Set([close]));
    this.expect(close);
    this.depth -= 1;
    return { t: 'delim', open, close, body };
  }

  atom(): LinearNode {
    const t = this.next();
    if (!t) throw refuse('a missing operand');
    if (t.k === 'number') return { t: 'run', text: t.v, kind: 'number' };
    if (t.k === 'letter') return this.letters(t.v);
    if (t.k === 'op') {
      if (t.v === '(') return this.group('(', ')');
      if (t.v === '[') return this.group('[', ']');
      if (t.v === '{') return this.group('{', '}');
      if (')]}&@^_/'.includes(t.v)) throw refuse(`an unexpected ${t.v}`);
      return { t: 'run', text: t.v, kind: 'op' };
    }
    if (t.k === 'cmd') return this.command(t.v);
    throw refuse('an unexpected space');
  }

  /** A letter, or a function name spelled with letters that follow it. */
  private letters(first: string): LinearNode {
    let word = first;
    let j = this.at;
    while (this.tokens[j]?.k === 'letter') {
      word += (this.tokens[j] as { v: string }).v;
      j += 1;
    }
    if (FUNCTIONS.has(word)) {
      this.at = j;
      return this.func(word);
    }
    const isGreekOrLatin = /\p{L}/u.test(first);
    return { t: 'run', text: first, kind: isGreekOrLatin ? 'letter' : 'op' };
  }

  private func(word: string): LinearNode {
    this.enter();
    let name: LinearNode = { t: 'run', text: word, kind: 'name' };
    if (LIMIT_FUNCTIONS.has(word) && this.isOp('_', false)) {
      this.next();
      name = { t: 'limLow', base: [name], limit: operand(this.atom()) };
    }
    name = this.scripts(name, true);
    const body = this.peek() && !this.isStop() ? operand(this.term()) : [];
    this.depth -= 1;
    return { t: 'func', name: [name], body };
  }

  private isStop(): boolean {
    const t = this.peek();
    return t?.k === 'op' && ')]}&@'.includes(t.v);
  }

  private command(name: string): LinearNode {
    if (Object.hasOwn(NAMES, name)) {
      const ch = NAMES[name]!;
      return { t: 'run', text: ch, kind: /\p{L}/u.test(ch) ? 'letter' : 'op' };
    }
    this.enter();
    try {
      if (name === 'sqrt') {
        let degree: LinearRow | undefined;
        if (this.isOp('[', false)) {
          this.next();
          degree = this.row(new Set([']']));
          this.expect(']');
        }
        return { t: 'rad', ...(degree ? { degree } : {}), body: operand(this.atom()) };
      }
      if (name === 'frac') {
        const num = operand(this.atom());
        return { t: 'frac', num, den: operand(this.atom()) };
      }
      if (Object.hasOwn(NARY, name)) {
        let sub: LinearRow | undefined;
        let sup: LinearRow | undefined;
        for (;;) {
          if (this.isOp('_', false) && !sub) {
            this.next();
            sub = operand(this.atom());
          } else if (this.isOp('^', false) && !sup) {
            this.next();
            sup = operand(this.atom());
          } else break;
        }
        const body = this.peek() && !this.isStop() ? operand(this.term()) : [];
        return {
          t: 'nary',
          chr: NARY[name]!,
          ...(sub ? { sub } : {}),
          ...(sup ? { sup } : {}),
          body,
        };
      }
      if (Object.hasOwn(ACCENTS, name))
        return { t: 'acc', chr: ACCENTS[name]!, body: operand(this.atom()) };
      if (name === 'bar' || name === 'overline') return { t: 'bar', body: operand(this.atom()) };
      if (name === 'matrix') {
        this.expect('(');
        const rows: LinearRow[][] = [];
        let row: LinearRow[] = [];
        for (;;) {
          row.push(this.row(new Set([')', '&', '@'])));
          const t = this.next();
          if (t?.k !== 'op') throw refuse('an open matrix');
          if (t.v === '&') continue;
          rows.push(row);
          row = [];
          if (t.v === ')') break;
        }
        return { t: 'matrix', rows };
      }
      throw refuse(`the name \\${name}`);
    } finally {
      this.depth -= 1;
    }
  }

  done(): boolean {
    return this.peek() === undefined;
  }
}

/** Parse a linear form. Throws a `CommandRefusal` for text that the syntax does not cover. */
export function parseLinear(src: string): LinearRow {
  if (src.length > MAX_LINEAR_LENGTH) throw refuse('too long');
  const parser = new Parser(tokenize(src));
  const row = parser.row(new Set());
  if (!parser.done()) throw refuse('an unexpected closing character');
  if (row.length === 0) throw refuse('no math');
  return row;
}

const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

const esc = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/**
 * The `m:oMath` content of a parsed linear form. Each run carries an empty `a:rPr` with `lang`
 * (and `i` for letters), which `equation-commands.ts` fills with the run format.
 */
export function ommlOf(row: LinearRow): string {
  const r = (n: LinearRow): string => n.map(node).join('');
  const arg = (tag: string, n: LinearRow): string =>
    n.length ? `<m:${tag}>${r(n)}</m:${tag}>` : `<m:${tag}/>`;
  const node = (n: LinearNode): string => {
    switch (n.t) {
      case 'run': {
        const plain = n.kind === 'name' ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : '';
        const italic = n.kind === 'letter' ? ' i="1"' : '';
        return `<m:r>${plain}<a:rPr lang="en-US"${italic}/><m:t>${esc(n.text)}</m:t></m:r>`;
      }
      case 'frac':
        return `<m:f>${arg('num', n.num)}${arg('den', n.den)}</m:f>`;
      case 'script':
        if (n.sub && n.sup)
          return `<m:sSubSup>${arg('e', n.base)}${arg('sub', n.sub)}${arg('sup', n.sup)}</m:sSubSup>`;
        if (n.sup) return `<m:sSup>${arg('e', n.base)}${arg('sup', n.sup)}</m:sSup>`;
        return `<m:sSub>${arg('e', n.base)}${arg('sub', n.sub ?? [])}</m:sSub>`;
      case 'rad':
        return n.degree
          ? `<m:rad>${arg('deg', n.degree)}${arg('e', n.body)}</m:rad>`
          : `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>${arg('e', n.body)}</m:rad>`;
      case 'delim': {
        if (n.open === '{') return r(n.body);
        const pr =
          n.open === '('
            ? ''
            : `<m:dPr><m:begChr m:val="${esc(n.open)}"/><m:endChr m:val="${esc(n.close)}"/></m:dPr>`;
        return `<m:d>${pr}${arg('e', n.body)}</m:d>`;
      }
      case 'nary': {
        const loc = n.chr === '∫' || n.chr === '∮' ? 'subSup' : 'undOvr';
        const hide =
          (n.sub ? '' : '<m:subHide m:val="1"/>') + (n.sup ? '' : '<m:supHide m:val="1"/>');
        return (
          `<m:nary><m:naryPr><m:chr m:val="${n.chr}"/><m:limLoc m:val="${loc}"/>${hide}</m:naryPr>` +
          `${arg('sub', n.sub ?? [])}${arg('sup', n.sup ?? [])}${arg('e', n.body)}</m:nary>`
        );
      }
      case 'func':
        return `<m:func>${arg('fName', n.name)}${arg('e', n.body)}</m:func>`;
      case 'limLow':
        return `<m:limLow>${arg('e', n.base)}${arg('lim', n.limit)}</m:limLow>`;
      case 'acc':
        return `<m:acc><m:accPr><m:chr m:val="${n.chr}"/></m:accPr>${arg('e', n.body)}</m:acc>`;
      case 'bar':
        return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>${arg('e', n.body)}</m:bar>`;
      case 'matrix':
        return `<m:m>${n.rows.map((row) => `<m:mr>${row.map((c) => arg('e', c)).join('')}</m:mr>`).join('')}</m:m>`;
    }
  };
  return r(row);
}

/** The Office Math markup of a linear form, as an `m:oMathPara` with its namespace. */
export function oMathParaOf(linear: string): string {
  return (
    `<m:oMathPara xmlns:m="${MATH_NS}"><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr>` +
    `<m:oMath>${ommlOf(parseLinear(linear))}</m:oMath></m:oMathPara>`
  );
}

/** The plain text of a linear form, for the fallback copy that a reader of no math shows. */
export function plainTextOf(row: LinearRow): string {
  const r = (n: LinearRow): string => n.map(node).join('');
  const wrap = (n: LinearRow): string => (n.length === 1 && n[0]!.t === 'run' ? r(n) : `(${r(n)})`);
  const node = (n: LinearNode): string => {
    switch (n.t) {
      case 'run':
        return n.text;
      case 'frac':
        return `${wrap(n.num)}/${wrap(n.den)}`;
      case 'script':
        return `${r(n.base)}${n.sub ? `_${wrap(n.sub)}` : ''}${n.sup ? `^${wrap(n.sup)}` : ''}`;
      case 'rad':
        return `${n.degree ? `${r(n.degree)}` : ''}√${wrap(n.body)}`;
      case 'delim':
        return n.open === '{' ? r(n.body) : `${n.open}${r(n.body)}${n.close}`;
      case 'nary':
        return `${n.chr}${n.sub ? `_${wrap(n.sub)}` : ''}${n.sup ? `^${wrap(n.sup)}` : ''} ${r(n.body)}`;
      case 'func':
        return `${r(n.name)} ${r(n.body)}`;
      case 'limLow':
        return `${r(n.base)}_${wrap(n.limit)}`;
      case 'acc':
      case 'bar':
        return r(n.body);
      case 'matrix':
        return `(${n.rows.map((row) => row.map(r).join(' ')).join('; ')})`;
    }
  };
  return r(row);
}
