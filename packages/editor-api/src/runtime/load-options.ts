/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Copied from DOCX Editor's `packages/editor-api/src/runtime/load-options.ts` (commit 5993bd223945).
// What `load(...)` accepts, and what it refuses.
//
// Four call shapes, as Office.js takes them: `load('text')`, `load(['text', 'style'])`,
// `load({ select: 'text', top: 5 })`, and the property tree `load({ font: { bold: true } })`.
// Each can name properties by path (`shapes/items/textFrame/textRange/text`). They all resolve
// to one tree of requests, one node for each object on a path, so an object's load planner
// reads a single shape and hands each navigation property its own node.
//
// REFUSAL IS THE POINT. A misspelled option key is the difference between "load selected
// properties" and "load nothing", and silently ignoring it produces a `PropertyNotLoaded` later
// at a call that looks correct. So an unknown key, a non-integer `top`, a property name that is
// not an identifier, or a value of the wrong type is `InvalidArgument` HERE, naming the option.
// A name that is not a property of the object it reaches is refused by that object's planner.
//
// THE TREE IS BOUNDED. At most 8 segments in a path or levels in a property tree, and at most
// 256 names in one call. A larger request is refused as it is read, so attacker-influenced data
// in a consumer's own pipeline cannot drive deep recursion or a large allocation here.
//
// Selected names are never used as object keys anywhere in this runtime — loaded values live in
// a `Map` — so `__proto__` as a property name is inert by construction rather than by filtering.
// It is still refused, because it is not a property any object of this API has.

import { failArgument, nearest } from './errors.ts';

/**
 * The object form of `load(...)`: which properties, and how much of a collection.
 *
 * An unknown key, a non-integer `top`, a property name that is not an identifier, or a value of
 * the wrong type is refused as `InvalidArgument` here, naming the option — a misspelled key is
 * the difference between "load selected properties" and "load nothing", which would otherwise
 * surface as a `PropertyNotLoaded` much later at a call that looks correct.
 *
 * @public
 */
export interface LoadQueryOptions {
  /** Which properties to load, as names or paths such as `items/fill/type`. */
  readonly select?: string | readonly string[];
  /**
   * Navigation properties to load with their default set, as names or paths such as
   * `items/fill`. A path that `select` also names deeper loads only what `select` names.
   */
  readonly expand?: string | readonly string[];
  /** For a collection: at most this many items. */
  readonly top?: number;
  /** For a collection: skip this many items first. */
  readonly skip?: number;
}

/**
 * The property-tree form of `load(...)`, as the Office.js `*LoadOptions` interfaces give it.
 *
 * `true` loads a property. An object loads a navigation property with the properties it names;
 * an empty object loads its default set. `$all: true` loads every scalar property of that
 * object. On a collection, the keys name properties of each item, and `$top` and `$skip` limit
 * the items. `items` is accepted as a key too, so `{ shapes: { items: { name: true } } }` and
 * `{ shapes: { name: true } }` load the same thing.
 *
 * @public
 */
export interface LoadPropertyOptions {
  /** Load every scalar property of this object, or of each item of a collection. */
  readonly $all?: boolean;
  /** For a collection: at most this many items. */
  readonly $top?: number;
  /** For a collection: skip this many items first. */
  readonly $skip?: number;
  /** A property: `true` loads it, an object loads a navigation property with what it names. */
  readonly [property: string]: boolean | number | LoadPropertyOptions | undefined;
}

/**
 * Everything `load(...)` accepts: one property name or path, several, a
 * {@link LoadQueryOptions} object, or a {@link LoadPropertyOptions} tree.
 *
 * @example
 * ```ts
 * shape.load('name');
 * shape.load(['name', 'left', 'fill/foregroundColor']);
 * slide.load('shapes/items/textFrame/textRange/text');
 * slide.load({ shapes: { items: { name: true, fill: { type: true } } } });
 * shapes.load({ select: 'items/name', top: 5 });
 * shapes.load('$all');
 * ```
 *
 * A path names a property through navigation properties, for example `fill/type` on a shape.
 * On a collection, `items/<path>` loads `items` and that path on each item in the same
 * `context.sync()`. A path without the `items/` prefix means the same thing on a collection, so
 * `shapes.load('name')` equals `shapes.load('items/name')`. `$all` loads every scalar property
 * of the object it ends on, and `items/$all` of each item. `load()` and `load('items')` on a
 * collection load every scalar property of each item, in one more round trip.
 *
 * A name that is not a property of the object it reaches is refused as `InvalidArgument`, with
 * the object and the name as the error's `target`.
 *
 * @public
 */
export type LoadOption = string | readonly string[] | LoadQueryOptions | LoadPropertyOptions;

export interface ResolvedLoadOptions {
  /** Selected scalar or navigation property names of this object, without `$all`. */
  readonly select: readonly string[];
  /** True when `$all` was named: load every scalar property as well as `select`. */
  readonly all: boolean;
  /** Navigation properties named with a deeper path, each with the request for that object. */
  readonly nested: ReadonlyMap<string, ResolvedLoadOptions>;
  /**
   * The request for each item of a collection, from `items/<path>` paths or an `items` key.
   * An object that is not a collection refuses it.
   */
  readonly items?: ResolvedLoadOptions;
  /** The property names that `items` selects. Empty when none was named. */
  readonly itemSelect: readonly string[];
  readonly top?: number;
  readonly skip?: number;
  /**
   * Set on the request for each item of a collection and on everything below it: a text read of
   * an item that cannot hold text (a picture, a group) fills nothing instead of failing the sync.
   */
  readonly orNull?: boolean;
}

/** True when a request names nothing: the object loads its default set. */
export function isDefaultLoad(request: ResolvedLoadOptions): boolean {
  return request.select.length === 0 && request.nested.size === 0 && request.items === undefined;
}

/** True when a request loads every scalar property: `$all`, or nothing named. */
export function loadsAllScalars(request: ResolvedLoadOptions): boolean {
  return request.all || isDefaultLoad(request);
}

/** An empty request: load the default set. */
export const DEFAULT_LOAD: ResolvedLoadOptions = Object.freeze({
  select: [],
  all: false,
  nested: new Map(),
  itemSelect: [],
});

const QUERY_KEYS = new Set(['select', 'expand', 'top', 'skip']);
const PROPERTY_NAME = /^[A-Za-z][A-Za-z0-9]*$/;
const ALL = '$all';
const ITEMS = 'items';
/** Segments in one path, and levels in one property tree. */
const MAX_DEPTH = 8;
/** Names in one call, over every path and level. */
const MAX_NAMES = 256;
const MAX_SEGMENT_LENGTH = 64;

/** A request while it is being built. */
interface Draft {
  readonly select: string[];
  all: boolean;
  readonly nested: Map<string, Draft>;
  top?: number;
  skip?: number;
}

function draft(): Draft {
  return { select: [], all: false, nested: new Map() };
}

/** Counts names across one call, and refuses past the cap. */
class Budget {
  #left = MAX_NAMES;
  constructor(readonly target: string) {}
  spend(): void {
    this.#left -= 1;
    if (this.#left < 0)
      failArgument(this.target, `at most ${MAX_NAMES} property names in one load`);
  }
}

function child(node: Draft, name: string): Draft {
  let next = node.nested.get(name);
  if (!next) {
    next = draft();
    node.nested.set(name, next);
  }
  return next;
}

/**
 * The segments of one path. The refusal names the loaded object alone, because the target must
 * never echo an arbitrary value.
 */
function segmentsOf(entry: unknown, target: string): readonly string[] {
  if (typeof entry !== 'string') failArgument(target, 'a property path as a string');
  const segments = entry.split('/', MAX_DEPTH + 1);
  const valid =
    segments.length <= MAX_DEPTH &&
    segments.every(
      (segment, i) =>
        segment.length <= MAX_SEGMENT_LENGTH &&
        (PROPERTY_NAME.test(segment) || (segment === ALL && i === segments.length - 1))
    );
  if (!valid) failArgument(target, "property names separated by '/', such as 'items/name'");
  return segments;
}

/** Add one path. `leaf` false makes every segment a navigation node (an `expand` path). */
function addPath(root: Draft, segments: readonly string[], leaf: boolean, budget: Budget): void {
  let node = root;
  const last = leaf ? segments.length - 1 : segments.length;
  for (let i = 0; i < last; i += 1) {
    budget.spend();
    node = child(node, segments[i]!);
  }
  if (!leaf) return;
  budget.spend();
  const name = segments[last]!;
  if (name === ALL) node.all = true;
  else if (!node.select.includes(name)) node.select.push(name);
}

function listOf(value: unknown, target: string): readonly unknown[] {
  if (typeof value === 'string') return value.split(',', MAX_NAMES + 1).map((e) => e.trim());
  if (Array.isArray(value)) {
    if (value.length > MAX_NAMES) failArgument(target, `at most ${MAX_NAMES} paths`);
    return value.map((entry: unknown) => (typeof entry === 'string' ? entry.trim() : entry));
  }
  return failArgument(target, 'a path string, an array of path strings, or an options object');
}

function addPaths(root: Draft, value: unknown, leaf: boolean, target: string): void {
  const budget = new Budget(target);
  const entries = listOf(value, target);
  // An empty entry is what a trailing comma or a stray whitespace string produces. Dropping it
  // silently would make `load(' ')` mean `load()` — "load the default set" — which is the
  // opposite of what was asked for. `segmentsOf` refuses it.
  if (entries.length === 0) failArgument(target, 'at least one property name');
  for (const entry of entries) addPath(root, segmentsOf(entry, target), leaf, budget);
}

function count(value: unknown, target: string): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 0) {
    failArgument(target, 'a whole number of 0 or more');
  }
  return value;
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false;
  const proto = Object.getPrototypeOf(value) as unknown;
  return proto === Object.prototype || proto === null;
}

/** One level of a property tree. */
function addTree(
  node: Draft,
  tree: Record<string, unknown>,
  target: string,
  depth: number,
  budget: Budget
): void {
  if (depth > MAX_DEPTH)
    failArgument(budget.target, `a property tree at most ${MAX_DEPTH} levels deep`);
  for (const key of Object.keys(tree)) {
    budget.spend();
    const value = tree[key];
    if (key === ALL) {
      if (typeof value !== 'boolean') failArgument(`${target}.${key}`, 'true or false');
      if (value) node.all = true;
    } else if (key === '$top') node.top = count(value, `${target}.${key}`);
    else if (key === '$skip') node.skip = count(value, `${target}.${key}`);
    else if (key.length > MAX_SEGMENT_LENGTH || !PROPERTY_NAME.test(key))
      failArgument(target, 'property names of letters and digits');
    else if (value === true) {
      if (!node.select.includes(key)) node.select.push(key);
    } else if (value === false || value === undefined) continue;
    else if (isPlainObject(value))
      addTree(child(node, key), value, `${target}.${key}`, depth + 1, budget);
    else failArgument(`${target}.${key}`, 'true, false, or a nested property tree');
  }
}

/**
 * The finished request. A navigation property named both alone and with a deeper path loads
 * its default set as well as the deeper path: `load('fill, fill/type')` loads all of `fill`.
 */
function finish(node: Draft): ResolvedLoadOptions {
  const itemsDraft = node.nested.get(ITEMS);
  const nested = new Map<string, ResolvedLoadOptions>();
  for (const [name, next] of node.nested) {
    if (name === ITEMS) continue;
    if (node.select.includes(name)) next.all = true;
    nested.set(name, finish(next));
  }
  const select = node.select.filter((name) => !nested.has(name));
  const items = itemsDraft ? finish(itemsDraft) : undefined;
  return {
    select,
    all: node.all,
    nested,
    ...(items ? { items } : {}),
    itemSelect: items ? items.select : [],
    ...(node.top === undefined ? {} : { top: node.top }),
    ...(node.skip === undefined ? {} : { skip: node.skip }),
  };
}

/**
 * One resolved request from any accepted call shape.
 *
 * `target` is the object being loaded, so a refusal names `presentation.slides`, not the
 * offending value — which may well be attacker-influenced data in a consumer's own pipeline.
 */
export function resolveLoadOption(
  option: LoadOption | undefined,
  target: string
): ResolvedLoadOptions {
  if (option === undefined) return DEFAULT_LOAD;
  const root = draft();
  if (typeof option === 'string' || Array.isArray(option)) {
    addPaths(root, option, true, target);
    return finish(root);
  }
  if (!isPlainObject(option))
    failArgument(target, 'a path string, an array of path strings, or an options object');

  const keys = Object.keys(option);
  if (!keys.some((key) => QUERY_KEYS.has(key))) {
    addTree(root, option, target, 1, new Budget(target));
    return finish(root);
  }
  for (const key of keys) {
    if (!QUERY_KEYS.has(key))
      failArgument(
        `${target}.${key}`,
        'the options select, expand, top, or skip',
        nearest(key, [...QUERY_KEYS])
      );
  }
  const query = option as LoadQueryOptions;
  if (query.select !== undefined) addPaths(root, query.select, true, `${target}.select`);
  if (query.expand !== undefined) addPaths(root, query.expand, false, `${target}.expand`);
  if (query.top !== undefined) root.top = count(query.top, `${target}.top`);
  if (query.skip !== undefined) root.skip = count(query.skip, `${target}.skip`);
  return finish(root);
}

/** Two requests for one object, as one. An empty request stands for the default set. */
function merge(a: ResolvedLoadOptions, b: ResolvedLoadOptions): ResolvedLoadOptions {
  const select = [...a.select];
  for (const name of b.select) if (!select.includes(name)) select.push(name);
  const nested = new Map(a.nested);
  for (const [name, next] of b.nested) {
    const mine = nested.get(name);
    nested.set(name, mine ? merge(mine, next) : next);
  }
  const items = a.items && b.items ? merge(a.items, b.items) : (a.items ?? b.items);
  return {
    select,
    all: loadsAllScalars(a) || loadsAllScalars(b),
    nested,
    ...(items ? { items } : {}),
    itemSelect: items ? items.select : [],
  };
}

/**
 * The request for each item of a collection: the `items` request merged with the names of the
 * collection's own request, because on a collection `name` means `items/name`. `load()` and
 * `load('items')` name every scalar property of each item, as in Office.js.
 */
export function itemRequest(request: ResolvedLoadOptions): ResolvedLoadOptions | undefined {
  const bare = isDefaultLoad(request) || request.select.includes(ITEMS);
  const own: ResolvedLoadOptions = {
    select: request.select.filter((name) => name !== ITEMS),
    all: request.all || bare,
    nested: request.nested,
    itemSelect: [],
  };
  const items = request.items;
  const ownNames = own.select.length > 0 || own.nested.size > 0 || own.all;
  const chosen = items ? (ownNames ? merge(own, items) : items) : ownNames ? own : undefined;
  return chosen ? { ...chosen, orNull: true } : undefined;
}
