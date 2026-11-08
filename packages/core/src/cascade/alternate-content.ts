// The branch of `mc:AlternateContent` that the engine reads (ECMA-376 Part 3, 10.2.1
// AlternateContent, 10.2.2 Choice, 10.2.3 Fallback). The cascade, find, the selection pane, the
// layout commands, and the edit rule (`automation/commands/alternate-copies.ts`) all call `alternateBranch`,
// so they see the same shapes.

import { attr, child, elements, NS, type Element } from '../store/index.ts';
import { isMathZone, isSupportedMath } from './math.ts';

/**
 * The namespaces a `mc:Choice` may require for the engine to select it. A namespace is here only
 * when the engine reads its markup in a shape tree or can skip it with no change to the picture.
 * `p14` holds extensions inside `p:extLst` (media, creation ids) and `p14:contentPart` (ink); the
 * second is not a shape kind, so `isDrawable` refuses a branch that holds it. VML (`v`), `a14`
 * picture effects, and chart extensions are not here: their `mc:Fallback` draws instead. `a14`
 * counts only for a branch whose `a14` markup is Office Math that the layout draws
 * (`mathOnly`).
 */
const UNDERSTOOD = new Set<string>([NS.p, NS.a, NS.r, NS.p14]);

/** The PresentationML shape kinds that the cascade resolves (`p:spTree` children). */
const DRAWABLE = new Set<string>(['sp', 'pic', 'grpSp', 'cxnSp', 'graphicFrame']);

/** Whether `el` is `mc:AlternateContent`. @internal */
export const isAlternateContent = (el: Element): boolean =>
  el.ns === NS.mc && el.local === 'AlternateContent';

/** The deepest nesting the scope walk follows: the reader's XML depth cap. */
const MAX_DEPTH = 256;

/**
 * For each `mc:AlternateContent` under a part root, the elements above it that declare
 * namespaces, nearest first. Built once per root: a part's tree is immutable, so an edit makes a
 * new root and a new entry.
 */
const scopesByRoot = new WeakMap<Element, Map<Element, readonly Element[]>>();

function declaringAncestors(ac: Element, root: Element | undefined): readonly Element[] {
  if (!root) return [];
  let map = scopesByRoot.get(root);
  if (!map) {
    const found = new Map<Element, readonly Element[]>();
    const walk = (el: Element, above: readonly Element[], depth: number): void => {
      const scope = el.nsDecls.length > 0 ? [el].concat(above) : above;
      for (const c of elements(el)) {
        if (isAlternateContent(c)) found.set(c, scope);
        if (depth < MAX_DEPTH) walk(c, scope, depth + 1);
      }
    };
    walk(root, [], 0);
    scopesByRoot.set(root, found);
    map = found;
  }
  // An `ac` that is not under `root` resolves on the root only.
  return map.get(ac) ?? [root];
}

/**
 * Each element's own prefix declarations, first one wins, built once per element: a `Requires`
 * list checks each prefix on every scope, and a part can declare thousands of prefixes.
 */
const declsByElement = new WeakMap<Element, ReadonlyMap<string, string>>();

function declsOf(el: Element): ReadonlyMap<string, string> {
  let map = declsByElement.get(el);
  if (!map) {
    const built = new Map<string, string>();
    for (const d of el.nsDecls) if (!built.has(d.prefix)) built.set(d.prefix, d.uri);
    declsByElement.set(el, built);
    map = built;
  }
  return map;
}

/** The namespace bound to `prefix` on the first of `scopes` that declares it. */
function namespaceOf(prefix: string, scopes: readonly Element[]): string | undefined {
  for (const el of scopes) {
    if (el.nsDecls.length === 0) continue;
    const uri = declsOf(el).get(prefix);
    if (uri !== undefined) return uri;
  }
  return undefined;
}

const mathOnlyCache = new WeakMap<Element, boolean>();

/**
 * Whether the `a14` markup of a branch is equations that the layout draws: one or more `a14:m`
 * zones that `isSupportedMath` accepts, and no other `a14` element outside an extension list
 * (`a:extLst`, `p:extLst`), whose content a reader may skip.
 */
function mathOnly(branch: Element): boolean {
  const hit = mathOnlyCache.get(branch);
  if (hit !== undefined) return hit;
  let zones = 0;
  let other = false;
  const walk = (el: Element, depth: number): void => {
    for (const c of elements(el)) {
      if (other) return;
      if (c.local === 'extLst' && (c.ns === NS.a || c.ns === NS.p)) continue;
      if (c.ns === NS.a14) {
        if (isMathZone(c) && isSupportedMath(c)) zones += 1;
        else other = true;
        continue;
      }
      if (depth < MAX_DEPTH) walk(c, depth + 1);
      else other = true;
    }
  };
  walk(branch, 0);
  const ok = zones > 0 && !other;
  mathOnlyCache.set(branch, ok);
  return ok;
}

/**
 * Whether every element of a branch is a shape kind the cascade draws, or another
 * `mc:AlternateContent`, which resolves in turn.
 */
function isDrawable(branch: Element): boolean {
  return elements(branch).every(
    (el) => (el.ns === NS.p && DRAWABLE.has(el.local)) || isAlternateContent(el)
  );
}

/**
 * The branch of `ac` that the engine reads: the first `mc:Choice` whose `Requires` prefixes all
 * name an understood namespace and whose content the cascade draws (an `a14` choice: equations
 * that the layout draws), else `mc:Fallback`
 * (ECMA-376 Part 3, 10.2.1). A prefix resolves in the scope of the `mc:Choice`: on it, on `ac`,
 * then on each element above `ac` up to `root` (the part root). A prefix that none of them
 * declares is not understood. Without `root`, only the `mc:Choice` and `ac` count.
 *
 * With no `mc:Fallback` and no Choice selected, ECMA-376 ignores the whole element. The engine
 * reads the first `mc:Choice` instead, so its shapes stay on the slide and in the selection pane
 * and an edit can reach them. Reference PDF: fixture `preserve-unsupported` draws the
 * `Requires="p14"` choice. @internal
 */
export function alternateBranch(ac: Element, root?: Element): Element | undefined {
  let first: Element | undefined;
  for (const branch of elements(ac)) {
    if (branch.ns !== NS.mc || branch.local !== 'Choice') continue;
    first ??= branch;
    const prefixes = (attr(branch, 'Requires') ?? '').split(/\s+/).filter(Boolean);
    if (prefixes.length === 0) continue;
    const scopes = [branch, ac].concat(declaringAncestors(ac, root));
    const understood = prefixes.every((prefix) => {
      const ns = namespaceOf(prefix, scopes);
      return ns !== undefined && (UNDERSTOOD.has(ns) || (ns === NS.a14 && mathOnly(branch)));
    });
    if (understood && isDrawable(branch)) return branch;
  }
  return child(ac, { ns: NS.mc, local: 'Fallback' }) ?? first;
}

/**
 * The children of a shape tree or group, with each `mc:AlternateContent` replaced by the content
 * of {@link alternateBranch}. @internal
 */
export function contentElements(tree: Element | undefined, root?: Element): Element[] {
  const out: Element[] = [];
  const add = (el: Element, depth: number): void => {
    if (!isAlternateContent(el)) return void out.push(el);
    // A chosen branch can hold another `mc:AlternateContent`: it resolves the same way.
    if (depth >= MAX_DEPTH) return;
    for (const inner of elements(alternateBranch(el, root))) add(inner, depth + 1);
  };
  for (const el of elements(tree)) add(el, 0);
  return out;
}
