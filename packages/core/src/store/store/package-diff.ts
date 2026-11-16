// The difference between two packages as ops: what turns the open package into another file.
//
// `replaceDocument` and document refresh apply a returned file as one transaction of ordinary
// ops, so undo, collaboration, and the part checks treat it as any other edit, and parts the
// file did not change keep their identity: the cascade, the layout, and the thumbnails reuse
// them, and nothing on screen paints again for them.
//
// Parts match by OPC part name. Equal binary parts and XML parts with the same canonical form
// produce no ops. A changed XML part becomes node ops: the children of each element match by
// identity (a shape's `p:cNvPr/@id`, a relationship's `Id`, a slide id, an extension URI),
// then by equal content in order, then by name in order. Matched children that moved get
// `moveNode`; unmatched ones are removed or inserted whole. A part whose node edit would be
// larger than its budget, whose root changed name, or whose root declares a namespace the open
// root does not (an `mc:Ignorable` list names prefixes) is replaced whole. Before it returns,
// the difference is applied to a copy and each changed part is compared with the file by its
// canonical form; a part that does not match is replaced whole. Store node ids of the two
// packages are unrelated: both come from document order. DOM-free.

import { isCanonicalPartName, partNameKey } from '@eigenpal/ooxml-core';
import { NS } from '../package/namespaces.ts';
import { mainPresentationPart, type PackagePart, type PptxPackage } from '../package/package.ts';
import { canonicalForm, serializeXmlPart } from '../package/serialize.ts';
import type { Element, Node } from '../package/tree.ts';
import { applyOps, type PptxOp } from './ops.ts';
import type { NodeSpec } from './tree-edit.ts';

/** Why two packages have no difference as ops. @internal */
export type PackageDiffRefusal = 'different-main-part' | 'unsafe-part-name';

/** The ops that turn one package into another, and the parts they change. @internal */
export type PackageDiff =
  | { readonly ok: true; readonly ops: readonly PptxOp[]; readonly parts: readonly string[] }
  | { readonly ok: false; readonly reason: PackageDiffRefusal; readonly detail?: string };

/** The most node ops one part may take before it is replaced whole. */
const PART_OP_BUDGET = 4000;
/** The largest `length(A) * length(B)` of a child list that gets an in-order content match. */
const LCS_CELLS = 250_000;
const CHUNK = 0x8000;

class ReplacePart extends Error {}

function toBase64(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i += CHUNK)
    s += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK) as unknown as number[]);
  return btoa(s);
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a === b) return true;
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i += 1) if (a[i] !== b[i]) return false;
  return true;
}

/** A content hash of each subtree: names, attributes (sorted), children; no ids or prefixes. */
class Hashes {
  readonly #memo = new WeakMap<Node, string>();

  of(node: Node): string {
    const known = this.#memo.get(node);
    if (known !== undefined) return known;
    let text: string;
    if (node.type === 'text') text = `"${node.value}`;
    else {
      const attrs = node.attrs.map((a) => `${a.ns}}${a.local}=${a.value}`).sort();
      text = `<${node.ns}}${node.local} ${attrs.join('\u0001')}>`;
      for (const child of node.children) text += this.of(child);
    }
    const hash = digest(text);
    this.#memo.set(node, hash);
    return hash;
  }
}

/** Two 32-bit FNV-1a variants over UTF-16 code units, as one 16-hex-digit string. */
function digest(text: string): string {
  let h1 = 0x811c9dc5;
  let h2 = 0x01000193 ^ text.length;
  for (let i = 0; i < text.length; i += 1) {
    const c = text.charCodeAt(i);
    h1 = Math.imul(h1 ^ c, 0x01000193);
    h2 = Math.imul(h2 ^ c, 0x5bd1e995) ^ (h2 >>> 15);
  }
  return (h1 >>> 0).toString(16).padStart(8, '0') + (h2 >>> 0).toString(16).padStart(8, '0');
}

const attrOf = (el: Element, ns: string, local: string): string | undefined =>
  el.attrs.find((a) => a.ns === ns && a.local === local)?.value;

/**
 * The identity of a child within its list, when the format gives it one: a shape's id, a slide
 * or master id, a relationship id, a content type entry, an extension URI, a section id.
 */
function identityOf(el: Element): string | undefined {
  const name = `${el.ns}}${el.local}`;
  const nv = el.children[0];
  if (nv?.type === 'element' && nv.local.startsWith('nv')) {
    const cNvPr = nv.children.find((c) => c.type === 'element' && c.local === 'cNvPr');
    const id = cNvPr?.type === 'element' ? attrOf(cNvPr, '', 'id') : undefined;
    if (id !== undefined) return `${name}#${id}`;
  }
  const key =
    attrOf(el, '', 'Id') ??
    attrOf(el, '', 'PartName') ??
    attrOf(el, '', 'Extension') ??
    attrOf(el, '', 'uri') ??
    (el.ns === NS.p &&
    /^(sldId|sldMasterId|sldLayoutId|notesMasterId|handoutMasterId)$/.test(el.local)
      ? attrOf(el, '', 'id')
      : undefined) ??
    (el.local === 'section' || el.local === 'cmAuthor' ? attrOf(el, '', 'id') : undefined);
  return key === undefined ? undefined : `${name}#${key}`;
}

/** The spec that inserts a copy of `node`. */
function specOf(el: Element): NodeSpec {
  return {
    ns: el.ns,
    local: el.local,
    ...(el.prefix ? { prefix: el.prefix } : {}),
    attrs: el.attrs.map((a) => ({
      ns: a.ns,
      local: a.local,
      ...(a.prefix ? { prefix: a.prefix } : {}),
      value: a.value,
    })),
    ...(el.nsDecls.length > 0 ? { nsDecls: el.nsDecls.map((d) => ({ ...d })) } : {}),
    children: el.children.map((c) => (c.type === 'text' ? { text: c.value } : specOf(c))),
  };
}

/** Pairs `[a index, b index]` of matched children: identity, then content in order, then name. */
function matchChildren(
  a: readonly Node[],
  b: readonly Node[],
  hashes: Hashes
): Map<number, number> {
  const pairs = new Map<number, number>(); // b index -> a index
  const usedA = new Set<number>();
  // 1. Identity, when it is unique on both sides.
  const ids = (list: readonly Node[]) => {
    const out = new Map<string, number | null>();
    list.forEach((n, i) => {
      const id = n.type === 'element' ? identityOf(n) : undefined;
      if (id !== undefined) out.set(id, out.has(id) ? null : i);
    });
    return out;
  };
  const idsA = ids(a);
  for (const [id, j] of ids(b)) {
    const i = idsA.get(id);
    if (i === undefined || i === null || j === null) continue;
    pairs.set(j, i);
    usedA.add(i);
  }
  // 2. Equal content, in order (longest common subsequence of the rest).
  const restA = a.map((_, i) => i).filter((i) => !usedA.has(i));
  const restB = b.map((_, j) => j).filter((j) => !pairs.has(j));
  if (restA.length > 0 && restB.length > 0 && restA.length * restB.length <= LCS_CELLS) {
    const ha = restA.map((i) => hashes.of(a[i]));
    const hb = restB.map((j) => hashes.of(b[j]));
    const n = ha.length;
    const m = hb.length;
    const table = new Uint32Array((n + 1) * (m + 1));
    for (let x = n - 1; x >= 0; x -= 1)
      for (let y = m - 1; y >= 0; y -= 1)
        table[x * (m + 1) + y] =
          ha[x] === hb[y]
            ? table[(x + 1) * (m + 1) + y + 1] + 1
            : Math.max(table[(x + 1) * (m + 1) + y], table[x * (m + 1) + y + 1]);
    for (let x = 0, y = 0; x < n && y < m;) {
      if (ha[x] === hb[y]) {
        pairs.set(restB[y], restA[x]);
        usedA.add(restA[x]);
        x += 1;
        y += 1;
      } else if (table[(x + 1) * (m + 1) + y] >= table[x * (m + 1) + y + 1]) x += 1;
      else y += 1;
    }
  }
  // 3. The same name, in order: an edited child of the same kind.
  for (let j = 0; j < b.length; j += 1) {
    if (pairs.has(j)) continue;
    const nb = b[j];
    if (nb.type !== 'element') continue;
    for (let i = 0; i < a.length; i += 1) {
      const na = a[i];
      if (usedA.has(i) || na.type !== 'element') continue;
      if (na.ns !== nb.ns || na.local !== nb.local) continue;
      // An identified child pairs only with the same identity, done above.
      if (identityOf(na) !== undefined || identityOf(nb) !== undefined) continue;
      pairs.set(j, i);
      usedA.add(i);
      break;
    }
  }
  return pairs;
}

/** The node ops of one part. Throws `ReplacePart` when the part should be replaced whole. */
class PartDiff {
  readonly ops: PptxOp[] = [];

  constructor(
    readonly part: string,
    readonly hashes: Hashes
  ) {}

  #push(op: PptxOp): void {
    this.ops.push(op);
    if (this.ops.length > PART_OP_BUDGET) throw new ReplacePart();
  }

  element(a: Element, b: Element): void {
    if (a.ns !== b.ns || a.local !== b.local) throw new ReplacePart();
    if (this.hashes.of(a) === this.hashes.of(b)) return;
    this.#attributes(a, b);
    const textA = a.children.every((c) => c.type === 'text');
    const textB = b.children.every((c) => c.type === 'text');
    if (textA && textB) {
      const before = a.children.map((c) => (c.type === 'text' ? c.value : '')).join('');
      const after = b.children.map((c) => (c.type === 'text' ? c.value : '')).join('');
      if (before !== after || a.children.length > 1)
        this.#push({ op: 'setText', part: this.part, node: a.id, text: after });
      return;
    }
    // Text on one side and elements on the other: the text goes first, or comes last.
    if (textA && a.children.length > 0)
      this.#push({ op: 'setText', part: this.part, node: a.id, text: '' });
    if (textB) {
      for (const child of a.children)
        this.#push({ op: 'removeNode', part: this.part, node: child.id });
      const text = b.children.map((c) => (c.type === 'text' ? c.value : '')).join('');
      if (text) this.#push({ op: 'setText', part: this.part, node: a.id, text });
      return;
    }
    this.#children(a, textA ? [] : a.children, b.children);
  }

  #attributes(a: Element, b: Element): void {
    const key = (x: { ns: string; local: string }) => `${x.ns}}${x.local}`;
    const before = new Map(a.attrs.map((x) => [key(x), x]));
    const after = new Set<string>();
    for (const attr of b.attrs) {
      after.add(key(attr));
      if (before.get(key(attr))?.value === attr.value) continue;
      const name = {
        ns: attr.ns,
        local: attr.local,
        ...(attr.prefix ? { prefix: attr.prefix } : {}),
      };
      this.#push({ op: 'setAttribute', part: this.part, node: a.id, name, value: attr.value });
    }
    for (const attr of a.attrs) {
      if (after.has(key(attr))) continue;
      const name = {
        ns: attr.ns,
        local: attr.local,
        ...(attr.prefix ? { prefix: attr.prefix } : {}),
      };
      this.#push({ op: 'setAttribute', part: this.part, node: a.id, name, value: null });
    }
  }

  #children(parent: Element, a: readonly Node[], b: readonly Node[]): void {
    const pairs = matchChildren(a, b, this.hashes);
    const matchedA = new Set(pairs.values());
    // The children as the ops leave them, by open node id; `null` marks an inserted child.
    const order: (string | null)[] = [];
    for (let i = 0; i < a.length; i += 1) {
      if (matchedA.has(i)) order.push(a[i].id);
      else this.#push({ op: 'removeNode', part: this.part, node: a[i].id });
    }
    for (let j = 0; j < b.length; j += 1) {
      const nb = b[j];
      const i = pairs.get(j);
      if (i === undefined) {
        // A text child among elements has no insert op: the part goes whole.
        if (nb.type !== 'element') throw new ReplacePart();
        this.#push({
          op: 'insertNode',
          part: this.part,
          parent: parent.id,
          index: j,
          spec: specOf(nb),
        });
        order.splice(j, 0, null);
        continue;
      }
      const id = a[i].id;
      const at = order.indexOf(id, j);
      if (at !== j) {
        order.splice(at, 1);
        order.splice(j, 0, id);
        this.#push({ op: 'moveNode', part: this.part, node: id, parent: parent.id, index: j });
      }
    }
    for (const [j, i] of pairs) {
      const na = a[i];
      const nb = b[j];
      if (na.type === 'element' && nb.type === 'element') this.element(na, nb);
      else if (na.type !== nb.type) throw new ReplacePart();
    }
  }
}

/** The ops that replace part `name` of `open` with `next`'s part, whole. */
function replaceWhole(name: string, hasOpen: boolean, part: PackagePart, out: PptxOp[]): void {
  if (hasOpen) out.push({ op: 'removePart', part: name });
  out.push(
    part.kind === 'xml'
      ? { op: 'addPart', part: name, xml: serializeXmlPart({ ...part.xml, partName: name }) }
      : { op: 'addPart', part: name, base64: toBase64(part.bytes) }
  );
}

/** Whether `next`'s root declares a prefix that `open`'s root does not bind the same way. */
function newRootPrefixes(open: Element, next: Element): boolean {
  const have = new Set(open.nsDecls.map((d) => `${d.prefix}=${d.uri}`));
  return next.nsDecls.some((d) => !have.has(`${d.prefix}=${d.uri}`));
}

/** The node ops of one XML part, or null when it should be replaced whole. */
function xmlPartOps(name: string, open: Element, next: Element, hashes: Hashes): PptxOp[] | null {
  if (open.ns !== next.ns || open.local !== next.local || newRootPrefixes(open, next)) return null;
  const diff = new PartDiff(name, hashes);
  try {
    diff.element(open, next);
  } catch (e) {
    if (e instanceof ReplacePart) return null;
    throw e;
  }
  return diff.ops;
}

/**
 * The ops that turn `open` into `next`, as one transaction. Parts match by name, folded as OPC
 * folds them; the open package's spelling stays. Refuses when the main presentation part of the
 * two packages differs.
 *
 * @internal
 */
export function diffPackages(open: PptxPackage, next: PptxPackage): PackageDiff {
  const mainOpen = mainPresentationPart(open);
  const mainNext = mainPresentationPart(next);
  if (!mainOpen || !mainNext || partNameKey(mainOpen) !== partNameKey(mainNext))
    return { ok: false, reason: 'different-main-part' };
  const openByKey = new Map<string, string>();
  for (const name of open.parts.keys()) openByKey.set(partNameKey(name), name);
  const hashes = new Hashes();
  const ops: PptxOp[] = [];
  const parts: string[] = [];
  const whole: [string, boolean, PackagePart][] = [];
  const seen = new Set<string>();
  for (const [nextName, part] of next.parts) {
    const key = partNameKey(nextName);
    seen.add(key);
    const name = openByKey.get(key) ?? nextName;
    if (!isCanonicalPartName(name)) return { ok: false, reason: 'unsafe-part-name', detail: name };
    const before = open.parts.get(name);
    if (before === part) continue;
    if (!before) {
      whole.push([name, false, part]);
      continue;
    }
    if (before.kind === 'binary' && part.kind === 'binary') {
      if (!sameBytes(before.bytes, part.bytes)) whole.push([name, true, part]);
      continue;
    }
    if (before.kind === 'xml' && part.kind === 'xml') {
      const a = before.xml.root;
      const b = part.xml.root;
      if (hashes.of(a) === hashes.of(b)) continue;
      const nodeOps = xmlPartOps(name, a, b, hashes);
      if (nodeOps === null) whole.push([name, true, part]);
      else if (nodeOps.length > 0) {
        for (const op of nodeOps) ops.push(op);
        parts.push(name);
      }
      continue;
    }
    whole.push([name, true, part]);
  }
  for (const name of open.parts.keys())
    if (!seen.has(partNameKey(name))) {
      ops.push({ op: 'removePart', part: name });
      parts.push(name);
    }
  for (const [name, hasOpen, part] of whole) {
    replaceWhole(name, hasOpen, part, ops);
    parts.push(name);
  }
  return verified(open, next, ops, parts, openByKey);
}

/**
 * Apply `ops` to a copy and compare each node-edited part with `next` by canonical form. A part
 * that differs, or ops that do not apply, fall back to whole-part replacement for those parts.
 */
function verified(
  open: PptxPackage,
  next: PptxPackage,
  ops: PptxOp[],
  parts: string[],
  openByKey: ReadonlyMap<string, string>
): PackageDiff {
  const nextByKey = new Map<string, PackagePart>();
  for (const [name, part] of next.parts) nextByKey.set(partNameKey(name), part);
  const edited = new Set<string>();
  for (const op of ops) if (op.op !== 'addPart' && op.op !== 'removePart') edited.add(op.part);
  const applied = applyOps(open, ops);
  const bad = new Set<string>();
  // Ops that do not apply: every node-edited part goes whole.
  if (!applied.ok) for (const name of edited) bad.add(name);
  else {
    for (const name of edited) {
      const want = nextByKey.get(partNameKey(name));
      const got = applied.pkg.parts.get(name);
      if (!want || !got || want.kind !== 'xml' || got.kind !== 'xml') continue;
      if (canonicalForm(want.xml.root) !== canonicalForm(got.xml.root)) bad.add(name);
    }
  }
  if (bad.size === 0) return { ok: true, ops, parts };
  const kept = ops.filter((op) => !bad.has(op.part));
  for (const name of bad) {
    const part = nextByKey.get(partNameKey(name));
    if (part) replaceWhole(openByKey.get(partNameKey(name)) ?? name, true, part, kept);
  }
  return { ok: true, ops: kept, parts };
}
