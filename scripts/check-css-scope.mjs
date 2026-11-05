#!/usr/bin/env node
// Keeps the shipped chrome stylesheet inside the editor, so it never styles the host page:
//
// 1. Every top-level selector starts with a `pptx-` class, optionally after a tag name
//    (`.pptx-menu`, `fieldset.pptx-dialog__section`). Pseudo-classes after that class only
//    narrow it. A rule for `body`, `:root`, `*`, `:not(.pptx-x)`, `:is(body, .pptx-x)`,
//    `body:has(.pptx-x)`, or a host class (`.dark .pptx-editor`) fails.
// 2. A sibling combinator (`+`, `~`) leads only to a `pptx-` class, at the top level and in
//    nested rules, so `.pptx-editor ~ div` and `& + p` fail.
// 3. Every named at-rule (`@keyframes`, `@property`, `@counter-style`, `@font-palette-values`,
//    `@position-try`, `@layer`) has a `pptx-` or `--pptx-` name, so it cannot replace or
//    reorder a host definition. `@scope` roots start with a `pptx-` class.
// 4. No `@import` and no `@font-face`: the stylesheet pulls nothing else in, and the engine
//    registers deck fonts through the FontFace API.
//
// Rules nested in a scoped rule (`&:hover`, `@starting-style`) are descendants of it, so only
// rule 2 applies to them, and a nested selector with `&` must start with it. Grouping rules (`@media`, `@supports`, `@container`,
// `@starting-style`) are checked by their contents.

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const STYLESHEET = 'packages/core/src/styles/editor.css';
const GROUPING = new Set(['media', 'supports', 'container', 'starting-style']);
const NAMED = new Set([
  'keyframes',
  'property',
  'counter-style',
  'font-palette-values',
  'position-try',
]);
const FORBIDDEN = new Set(['import', 'font-face']);
/** A compound that starts inside the editor: an optional tag, then a `pptx-` class. */
const SCOPED_START = /^(?:[a-zA-Z][\w-]*|\*)?\.pptx-[\w-]/;
const PPTX_NAME = /^(?:--)?pptx-[\w-]*$/;

/** Remove comments, keeping line breaks so the reported lines stay right. */
function stripComments(css) {
  // Strings first, so `content: "/*"` does not open a comment that hides the rules after it.
  return css.replace(/("(?:[^"\\\n]|\\.)*"|'(?:[^'\\\n]|\\.)*')|\/\*[\s\S]*?\*\//g, (m, string) =>
    string ? m : m.replace(/[^\n]/g, ' ')
  );
}

/** The index after the block that opens at `open` (a `{`), skipping strings. */
function blockEnd(css, open) {
  let depth = 0;
  for (let i = open; i < css.length; i++) {
    const c = css[i];
    if (c === '\\') i++;
    else if (c === '"' || c === "'") {
      let j = i + 1;
      while (j < css.length && css[j] !== c) j += css[j] === '\\' ? 2 : 1;
      i = j;
    } else if (c === '{') depth++;
    else if (c === '}' && --depth === 0) return i + 1;
  }
  return css.length;
}

/** Split at top-level occurrences of `at` (not inside `(...)`, `[...]`, or a string). */
function splitTop(text, at) {
  const parts = [];
  let depth = 0;
  let start = 0;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === '\\') i++;
    else if (c === '"' || c === "'") {
      const close = text.indexOf(c, i + 1);
      i = close < 0 ? text.length : close;
    } else if (c === '(' || c === '[') depth++;
    else if (c === ')' || c === ']') depth--;
    else if (depth === 0 && at.test(c)) {
      parts.push({ text: text.slice(start, i), sep: c });
      start = i + 1;
    }
  }
  parts.push({ text: text.slice(start), sep: '' });
  return parts;
}

/** The compounds of a selector with the combinator before each (`''` for the first). */
function compounds(selector) {
  const out = [];
  let combinator = '';
  for (const part of splitTop(selector, /[\s>+~]/)) {
    const text = part.text.trim();
    if (text) {
      out.push({ combinator, text });
      combinator = ' ';
    }
    // `>`, `+`, or `~` replaces the descendant combinator that whitespace implies.
    if (part.sep.trim()) combinator = part.sep;
  }
  // A selector that ends in a combinator (`.pptx-a +`) leads to nothing scoped.
  if (combinator === '+' || combinator === '~') out.push({ combinator, text: '' });
  return out;
}

/** Whether every sibling combinator in a selector leads to a `pptx-` class. */
function siblingsScoped(selector) {
  // Only a sibling of the first compound can leave the editor: once a descendant or child step
  // goes under it, every sibling after that is inside it too (`.pptx-menu li + li`).
  for (const c of compounds(selector)) {
    if (c.combinator === ' ' || c.combinator === '>') return true;
    if ((c.combinator === '+' || c.combinator === '~') && !SCOPED_START.test(c.text)) return false;
  }
  return true;
}

/** Whether a top-level selector starts inside the editor and stays in it. */
export function selectorScoped(selector) {
  const first = compounds(selector)[0];
  return (
    !!first && first.combinator === '' && SCOPED_START.test(first.text) && siblingsScoped(selector)
  );
}

/** Whether a nested selector (relative to a scoped parent) stays in the editor. */
export function nestedSelectorScoped(selector) {
  // A leading combinator (`+ p`) is relative to the parent: read it as `& + p`.
  const relative = /^[>+~]/.test(selector.trim()) ? `& ${selector.trim()}` : selector;
  // A selector with `&` is not prefixed by its parent: the `&` must start it, so the rule
  // stays at or under the parent (`body:has(&)` and `:is(body, &)` style the host).
  if (relative.includes('&') && !compounds(relative)[0]?.text.startsWith('&')) return false;
  return siblingsScoped(relative);
}

/** The selectors of a prelude, split at top-level commas. */
function selectorList(prelude) {
  return splitTop(prelude, /,/)
    .map((p) => p.text.trim())
    .filter(Boolean);
}

/** The index of the next `char` at or after `from` that is not inside a string, or -1. */
function nextOutsideStrings(css, char, from) {
  for (let i = from; i < css.length; i++) {
    const c = css[i];
    if (c === '\\') i++;
    else if (c === '"' || c === "'") {
      let j = i + 1;
      while (j < css.length && css[j] !== c) j += css[j] === '\\' ? 2 : 1;
      i = j;
    } else if (c === char) return i;
  }
  return -1;
}

/** Every scope failure in a stylesheet, as `line: message`. */
export function cssScopeErrors(source) {
  const css = stripComments(source);
  const errors = [];
  const lineOf = (i) => css.slice(0, i).split('\n').length;
  // `nested`: inside a scoped style rule or a `@scope` block, where selectors are descendants.
  const walk = (from, to, nested) => {
    let i = from;
    while (i < to) {
      const open = nextOutsideStrings(css, '{', i);
      const semi = nextOutsideStrings(css, ';', i);
      const close = nextOutsideStrings(css, '}', i);
      const blockOpen = open >= 0 && open < to ? open : -1;
      // A declaration or a statement at-rule ends before any block.
      if (semi >= 0 && semi < to && (blockOpen < 0 || semi < blockOpen)) {
        const text = css.slice(i, semi).trim();
        const at = i + css.slice(i, semi).search(/\S/);
        const name = /^@([\w-]+)/.exec(text)?.[1];
        if (name === 'layer') {
          for (const layer of splitTop(text.slice(6), /,/).map((p) => p.text.trim()))
            if (!PPTX_NAME.test(layer))
              errors.push(`${lineOf(at)}: @layer ${layer} needs a pptx- name`);
        } else if (name && name !== 'charset') {
          errors.push(`${lineOf(at)}: @${name} is not allowed here`);
        } else if (!name && !nested && text) {
          errors.push(`${lineOf(at)}: a declaration outside a rule`);
        }
        i = semi + 1;
        continue;
      }
      if (blockOpen < 0 || (close >= 0 && close < blockOpen)) return;
      const prelude = css.slice(i, blockOpen).trim();
      const at = i + css.slice(i, blockOpen).search(/\S/);
      const end = blockEnd(css, blockOpen);
      const name = /^@([\w-]+)/.exec(prelude)?.[1];
      if (name) {
        const ident = prelude.slice(name.length + 1).trim();
        if (FORBIDDEN.has(name)) errors.push(`${lineOf(at)}: @${name} is not allowed`);
        else if (NAMED.has(name)) {
          if (!PPTX_NAME.test(ident))
            errors.push(`${lineOf(at)}: @${name} ${ident} needs a pptx- name`);
        } else if (name === 'layer') {
          if (ident && !PPTX_NAME.test(ident))
            errors.push(`${lineOf(at)}: @layer ${ident} needs a pptx- name`);
          walk(blockOpen + 1, end - 1, nested);
        } else if (name === 'scope') {
          const scopeRoot = /^\(([^)]*)\)/.exec(ident)?.[1] ?? '';
          const roots = selectorList(scopeRoot);
          if (!nested && (!roots.length || !roots.every(selectorScoped)))
            errors.push(`${lineOf(at)}: @scope ${ident} needs a pptx- root`);
          walk(blockOpen + 1, end - 1, true);
        } else if (GROUPING.has(name)) walk(blockOpen + 1, end - 1, nested);
        else errors.push(`${lineOf(at)}: @${name} is not checked`);
      } else {
        const check = nested ? nestedSelectorScoped : selectorScoped;
        for (const selector of selectorList(prelude))
          if (!check(selector))
            errors.push(`${lineOf(at)}: "${selector}" is not scoped to a pptx- class`);
        // Nested rules in the body are descendants of this rule.
        walk(blockOpen + 1, end - 1, true);
      }
      i = end;
    }
  };
  walk(0, css.length, false);
  return errors;
}

function main() {
  const errors = cssScopeErrors(readFileSync(join(root, STYLESHEET), 'utf8'));
  if (errors.length) {
    console.error(`✘ ${STYLESHEET} styles outside the editor:\n`);
    for (const error of errors) console.error(`  - ${STYLESHEET}:${error}`);
    process.exit(1);
  }
  console.log('✓ the chrome stylesheet is scoped to pptx- classes.');
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) main();
