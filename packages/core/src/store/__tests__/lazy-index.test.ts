// A transaction learns the new places of shifted siblings only when one of them is read, and
// layers its index over the index it opened. Random runs of inserts, moves, removals, and
// attribute edits in one transaction give the tree that one transaction per op gives, and an
// index that matches a fresh walk of that tree. Many commits on one large part keep the index
// work linear.

import { expect, test } from 'bun:test';
import {
  applyOps,
  nodeIndex,
  NS,
  parseXmlPart,
  serializeXmlPart,
  xmlPartOf,
  type Element,
  type NodeId,
  type PptxOp,
  type PptxPackage,
  type XmlPart,
} from '../index.ts';
import { nodeIndexFlattenCount, nodeIndexTestRecorder } from '../store/node-index.ts';
import { openTransaction } from '../store/ops.ts';
import { workingTreeTestRecorder } from '../store/working-tree.ts';

const PART = '/ppt/slides/slide1.xml';

/** A package with one part: a root with `groups` children, each with `kids` children. */
function pkgOf(groups: number, kids: number): PptxPackage {
  const g = (i: number) =>
    `<p:g n="${i}">${Array.from({ length: kids }, (_, k) => `<p:k n="${i}.${k}"/>`).join('')}</p:g>`;
  const r = parseXmlPart(
    PART,
    `<p:root xmlns:p="${NS.p}">${Array.from({ length: groups }, (_, i) => g(i)).join('')}</p:root>`
  );
  if (!r.ok) throw new Error('fixture');
  return { parts: new Map([[PART, { kind: 'xml', partName: PART, xml: r.part }]]) };
}

/** The ids whose index entry differs from a fresh walk of the part's tree. */
function stale(part: XmlPart): NodeId[] {
  const kept = nodeIndex(part);
  const fresh = nodeIndex({ ...part, root: { ...part.root } });
  const out: NodeId[] = [...kept.keys()].filter((id) => !fresh.has(id));
  for (const [id, loc] of fresh) {
    const k = kept.get(id);
    const same = k?.node === (loc.parent === null ? part.root : loc.node);
    if (k?.parent !== loc.parent || k?.index !== loc.index || !same) out.push(id);
  }
  return out;
}

function random(seed: number) {
  let x = seed;
  return (n: number) => {
    x = (x * 1103515245 + 12345) & 0x7fffffff;
    return x % n;
  };
}

/** Random valid ops against the tree as `pkg` holds it, each checked by applying it. */
function randomOps(pkg: PptxPackage, r: (n: number) => number, count: number): PptxOp[] {
  const ops: PptxOp[] = [];
  let now = pkg;
  for (let i = 0; i < count; i += 1) {
    const xml = xmlPartOf(now, PART)!;
    const groups = xml.root.children.filter((c): c is Element => c.type === 'element');
    const group = groups[r(groups.length)];
    if (!group) break;
    const kids = group.children.filter((c): c is Element => c.type === 'element');
    const kid = kids[r(Math.max(1, kids.length))];
    const pick = r(4);
    let op: PptxOp;
    if (pick === 0 || !kid)
      op = {
        op: 'insertNode',
        part: PART,
        parent: group.id,
        index: r(group.children.length + 1),
        spec: {
          ns: NS.p,
          local: 'k',
          prefix: 'p',
          attrs: [{ ns: '', local: 'n', value: `new${i}` }],
        },
      };
    else if (pick === 1) op = { op: 'removeNode', part: PART, node: kid.id };
    else if (pick === 2) {
      const to = groups[r(groups.length)]!;
      op = {
        op: 'moveNode',
        part: PART,
        node: kid.id,
        parent: to.id,
        index: r(to.children.length),
      };
    } else
      op = {
        op: 'setAttribute',
        part: PART,
        node: kid.id,
        name: { ns: '', local: 'v' },
        value: String(i),
      };
    const applied = applyOps(now, [op]);
    if (!applied.ok) continue;
    now = applied.pkg;
    ops.push(op);
  }
  return ops;
}

test('one transaction gives the tree and the index of one transaction per op', () => {
  for (let seed = 1; seed <= 60; seed += 1) {
    const r = random(seed);
    const pkg = pkgOf(1 + r(4), r(40));
    const ops = randomOps(pkg, r, 80);
    let stepwise = pkg;
    for (const op of ops) {
      const s = applyOps(stepwise, [op]);
      if (!s.ok) throw new Error('step');
      stepwise = s.pkg;
    }
    const before = serializeXmlPart(xmlPartOf(pkg, PART)!);
    const once = applyOps(pkg, ops);
    expect(once.ok).toBe(true);
    // The transaction writes in place only into the copies it made, never into `pkg`.
    expect(serializeXmlPart(xmlPartOf(pkg, PART)!)).toBe(before);
    if (!once.ok) continue;
    const part = xmlPartOf(once.pkg, PART)!;
    expect(serializeXmlPart(part)).toBe(serializeXmlPart(xmlPartOf(stepwise, PART)!));
    expect(stale(part)).toEqual([]);
    expect(stale(xmlPartOf(stepwise, PART)!)).toEqual([]);
  }
});

test('2000 commits of one edit each on a 20000-node part write about what they change', () => {
  let pkg = pkgOf(1, 20_000);
  const work = workingTreeTestRecorder();
  const kids = (xmlPartOf(pkg, PART)!.root.children[0] as Element).children;
  work.reset();
  const flattened = nodeIndexFlattenCount();
  for (let i = 0; i < 2000; i += 1) {
    const r = applyOps(pkg, [
      {
        op: 'setAttribute',
        part: PART,
        node: kids[(i * 7) % kids.length]!.id,
        name: { ns: '', local: 'v' },
        value: String(i),
      },
    ]);
    if (!r.ok) throw new Error('op');
    pkg = r.pkg;
  }
  // Each commit writes its path (3 entries); a flatten every 9 layers copies 20002 entries.
  // A copy per commit would write 2000 × 20002.
  expect(work.indexWrites).toBe(2000 * 3);
  expect(nodeIndexFlattenCount() - flattened).toBeLessThanOrEqual((2000 / 9 + 1) * 20_002);
  expect(stale(xmlPartOf(pkg, PART)!)).toEqual([]);
});

test('an element read during a transaction stays as it was read', () => {
  for (let seed = 1; seed <= 60; seed += 1) {
    const r = random(seed);
    const pkg = pkgOf(1 + r(3), 1 + r(30));
    const ops = randomOps(pkg, r, 60);
    const tx = openTransaction(pkg);
    const held: { el: Element; xml: string }[] = [];
    const show = (el: Element) => serializeXmlPart({ partName: PART, root: el, nextId: 0 });
    for (const op of ops) {
      expect(tx.apply(op)).toBeNull();
      // Read the root, a group, or a kid between ops, as a command reads what it edits next.
      const root = tx.element(PART, xmlPartOf(pkg, PART)!.root.id)!;
      const groups = root.children.filter((c): c is Element => c.type === 'element');
      const group = groups[r(groups.length)];
      const pick = r(3);
      const id = pick === 0 || !group ? root.id : pick === 1 ? group.id : group.children[0]?.id;
      const el = id === undefined ? undefined : tx.element(PART, id);
      if (el) held.push({ el, xml: show(el) });
    }
    const done = tx.finish();
    for (const { el, xml } of held) expect(show(el)).toBe(xml);
    const once = applyOps(pkg, ops);
    if (!once.ok) throw new Error('ops');
    expect(serializeXmlPart(xmlPartOf(done, PART)!)).toBe(
      serializeXmlPart(xmlPartOf(once.pkg, PART)!)
    );
  }
});

test('ops on the runs of one long paragraph copy its child list once', () => {
  const pkg = pkgOf(1, 20_000);
  const group = xmlPartOf(pkg, PART)!.root.children[0] as Element;
  const ops: PptxOp[] = group.children.map((kid, i) => ({
    op: 'setAttribute',
    part: PART,
    node: kid.id,
    name: { ns: '', local: 'v' },
    value: String(i),
  }));
  const work = workingTreeTestRecorder();
  work.reset();
  const start = performance.now();
  const done = applyOps(pkg, ops);
  expect(done.ok).toBe(true);
  // The first op copies the path (3 entries); each later one writes its kid and then writes
  // into the group copy in place. A copy of the 20000 kids per op took seconds.
  expect(work.indexWrites).toBe(3 + (ops.length - 1));
  expect(performance.now() - start).toBeLessThan(2000);
});

test('a held element keeps its subtree when a later edit copies a path through it', () => {
  const r = parseXmlPart(
    PART,
    `<p:sld xmlns:p="${NS.p}"><p:t><p:p><p:r/></p:p><p:p><p:r/></p:p></p:t></p:sld>`
  );
  if (!r.ok) throw new Error('fixture');
  const t = r.part.root.children[0] as Element;
  const [p1, p2] = t.children as Element[];
  const pkg: PptxPackage = {
    parts: new Map([[PART, { kind: 'xml', partName: PART, xml: r.part }]]),
  };
  const tx = openTransaction(pkg);
  const attr = (node: string, value: string) =>
    tx.apply({ op: 'setAttribute', part: PART, node, name: { ns: '', local: 'v' }, value });
  const show = (el: Element) => serializeXmlPart({ partName: PART, root: el, nextId: 0 });
  // The tree owns copies of the second paragraph and its path after this edit.
  expect(attr(p2!.children[0]!.id, '1')).toBeNull();
  const held = tx.element(PART, t.id)!;
  const before = show(held);
  // A copy of the held text body shares the owned second paragraph with it.
  expect(attr(p1!.children[0]!.id, '2')).toBeNull();
  const spec = { ns: NS.p, local: 'r', prefix: 'p', attrs: [] };
  expect(tx.apply({ op: 'insertNode', part: PART, parent: p2!.id, index: 0, spec })).toBeNull();
  expect(show(held)).toBe(before);
});

/** A three-level tree: 3 `p:a`, each with 4 `p:b`, each with 3 `p:c` that hold text. */
function deepPkg(): PptxPackage {
  const c = '<p:c>t</p:c>'.repeat(3);
  const a = `<p:a>${`<p:b>${c}</p:b>`.repeat(4)}</p:a>`;
  const r = parseXmlPart(PART, `<p:root xmlns:p="${NS.p}">${a.repeat(3)}</p:root>`);
  if (!r.ok) throw new Error('fixture');
  return { parts: new Map([[PART, { kind: 'xml', partName: PART, xml: r.part }]]) };
}

const allElements = (el: Element, out: Element[] = []): Element[] => {
  out.push(el);
  for (const k of el.children) if (k.type === 'element') allElements(k, out);
  return out;
};

test('elements held between random ops, revisions, and the base never change', () => {
  const show = (el: Element) => serializeXmlPart({ partName: PART, root: el, nextId: 0 });
  for (let seed = 1; seed < 300; seed += 1) {
    const r = random(seed);
    const pkg = deepPkg();
    const base = show(xmlPartOf(pkg, PART)!.root);
    const tx = openTransaction(pkg);
    let step = pkg;
    const held: { el: Element; xml: string }[] = [];
    for (let i = 0; i < 60; i += 1) {
      const els = allElements(xmlPartOf(step, PART)!.root);
      const el = els[r(els.length)]!;
      const other = els[r(els.length)]!;
      const pick = r(6);
      const op: PptxOp =
        pick === 0
          ? {
              op: 'insertNode',
              part: PART,
              parent: el.id,
              index: r(el.children.length + 1),
              spec: { ns: NS.p, local: 'n', prefix: 'p', attrs: [] },
            }
          : pick === 1
            ? { op: 'removeNode', part: PART, node: el.id }
            : pick === 2
              ? {
                  op: 'moveNode',
                  part: PART,
                  node: el.id,
                  parent: other.id,
                  index: r(other.children.length + 1),
                }
              : pick === 3
                ? { op: 'setText', part: PART, node: el.id, text: `x${i}` }
                : {
                    op: 'setAttribute',
                    part: PART,
                    node: el.id,
                    name: { ns: '', local: 'v' },
                    value: String(i),
                  };
      const applied = applyOps(step, [op]);
      if (!applied.ok) continue;
      step = applied.pkg;
      expect(tx.apply(op)).toBeNull();
      if (r(2)) {
        const ids = allElements(xmlPartOf(step, PART)!.root);
        const h = tx.element(PART, ids[r(ids.length)]!.id);
        if (h) held.push({ el: h, xml: show(h) });
      }
    }
    const done = tx.finish();
    expect(show(xmlPartOf(pkg, PART)!.root)).toBe(base);
    for (const { el, xml } of held) expect(show(el)).toBe(xml);
    const part = xmlPartOf(done, PART)!;
    expect(show(part.root)).toBe(show(xmlPartOf(step, PART)!.root));
    expect(stale(part)).toEqual([]);
  }
});

test('an insert near the root of a large part, never committed, walks no index', () => {
  // A toolbar's `can` probe of Insert Text Box on a slide of 20,000 shapes.
  const pkg = pkgOf(2, 10_000);
  const root = xmlPartOf(pkg, PART)!.root;
  const spec = { ns: NS.p, local: 'k', prefix: 'p', attrs: [] };
  const recorder = nodeIndexTestRecorder();
  recorder.reset();
  const tx = openTransaction(pkg);
  for (let i = 0; i < 3; i += 1)
    expect(tx.apply({ op: 'insertNode', part: PART, parent: root.id, index: 2, spec })).toBeNull();
  expect(recorder.completeBuilds).toBe(0);
  // A commit hands an exact index to the new root.
  const done = tx.finish();
  expect(stale(xmlPartOf(done, PART)!)).toEqual([]);
  expect(xmlPartOf(done, PART)!.root.children.length).toBe(5);
});

test('a lookup deep in a part without an index builds the index once', () => {
  const pkg = pkgOf(2, 5_000);
  const deep = (xmlPartOf(pkg, PART)!.root.children[1] as Element).children[4_999]!;
  const recorder = nodeIndexTestRecorder();
  recorder.reset();
  const tx = openTransaction(pkg);
  for (let i = 0; i < 100; i += 1) expect(tx.element(PART, deep.id)).toBe(deep as Element);
  expect(recorder.completeBuilds).toBe(1);
});
