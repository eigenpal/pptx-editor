// Regression tests for audit findings in the op and store layer: S4 (depth
// through ops), K3 (case-variant part names), K5 (listener isolation), K7 (kinds after a move),
// K8 (node ids after undo), P3 (one id index per transaction), malformed op shapes, and the
// serializer's prefix choice.

import { beforeAll, describe, expect, test } from 'bun:test';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  applyOps,
  find,
  nodeById,
  nodeIndex,
  NS,
  parseXmlPart,
  PptxDocumentStore,
  readPptx,
  serializeXmlPart,
  writePptx,
  xmlPartOf,
  type Element,
  type Node,
  type NodeId,
  type NodeSpec,
  type PptxOp,
  type PptxPackage,
  type StoreChange,
  type XmlPart,
} from '../index.ts';
import { nodeIndexTestRecorder } from '../store/node-index.ts';
import { workingTreeTestRecorder } from '../store/working-tree.ts';
import { MAX_XML_DEPTH } from '@eigenpal/ooxml-core';

const SLIDE = '/ppt/slides/slide1.xml';

function deck(): PptxPackage {
  const r = readPptx(
    buildPptx({ slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100 }) }] })
  );
  if (!r.ok) throw new Error('fixture');
  return r.pkg;
}

const root = (pkg: PptxPackage) => xmlPartOf(pkg, SLIDE)!.root;

function reopen(pkg: PptxPackage): PptxPackage {
  const r = readPptx(writePptx(pkg));
  if (!r.ok) throw new Error(`reopen failed: ${JSON.stringify(r.error)}`);
  return r.pkg;
}

describe('S4: ops cannot build a tree the reader refuses', () => {
  const chain = (n: number): NodeSpec =>
    n === 1 ? { ns: 'urn:x', local: 'e' } : { ns: 'urn:x', local: 'e', children: [chain(n - 1)] };

  test('nested inserts stop at the reader depth, and what is accepted reopens', () => {
    const s = new PptxDocumentStore(deck());
    let parent = root(s.package).id;
    let level = 1;
    let refused = false;
    // Bounded: the old store never refused, so an unbounded loop would hang instead of failing.
    for (let attempt = 0; attempt < 5 && !refused; attempt += 1) {
      const before = xmlPartOf(s.package, SLIDE)!.nextId;
      const r = s.transact([{ op: 'insertNode', part: SLIDE, parent, index: 0, spec: chain(100) }]);
      if (!r.ok) {
        expect(r.rejection.code).toBe('too-deep');
        refused = true;
      } else {
        parent = `n${before + 99}`;
        level += 100;
      }
    }
    expect(refused).toBe(true);
    expect(level).toBeLessThanOrEqual(MAX_XML_DEPTH);
    const fill = MAX_XML_DEPTH - level;
    expect(
      s.transact([{ op: 'insertNode', part: SLIDE, parent, index: 0, spec: chain(fill) }]).ok
    ).toBe(true);
    expect(
      s.transact([{ op: 'insertNode', part: SLIDE, parent, index: 0, spec: chain(fill + 1) }]).ok
    ).toBe(false);
    expect(() => reopen(s.package)).not.toThrow();
  });

  test('a move that would exceed the depth is refused', () => {
    const s = new PptxDocumentStore(deck());
    const top = root(s.package).id;
    const first = xmlPartOf(s.package, SLIDE)!.nextId;
    expect(
      s.transact([{ op: 'insertNode', part: SLIDE, parent: top, index: 0, spec: chain(200) }]).ok
    ).toBe(true);
    const second = xmlPartOf(s.package, SLIDE)!.nextId;
    expect(
      s.transact([{ op: 'insertNode', part: SLIDE, parent: top, index: 0, spec: chain(100) }]).ok
    ).toBe(true);
    const deepest = `n${first + 199}`;
    const r = s.transact([
      { op: 'moveNode', part: SLIDE, node: `n${second}`, parent: deepest, index: 0 },
    ]);
    expect(r.ok ? null : r.rejection.code).toBe('too-deep');
  });
});

describe('K3: part names are compared as OPC does, ignoring ASCII case', () => {
  test('addPart refuses a case variant of an existing part', () => {
    const s = new PptxDocumentStore(deck());
    const r = s.transact([{ op: 'addPart', part: '/PPT/Presentation.xml', xml: '<a/>' }]);
    expect(r.ok ? null : r.rejection).toEqual({
      code: 'part-exists',
      part: '/PPT/Presentation.xml',
    });
    expect(() => writePptx(s.package)).not.toThrow();
  });

  test('the check covers parts added earlier in the same transaction', () => {
    const s = new PptxDocumentStore(deck());
    const r = s.transact([
      { op: 'addPart', part: '/ppt/media/image1.png', base64: 'AAEC' },
      { op: 'addPart', part: '/ppt/media/Image1.png', base64: 'AAEC' },
    ]);
    expect(r.ok ? null : r.rejection.code).toBe('part-exists');
    // A part needs a content type: declare one for `.png` first.
    const types = xmlPartOf(s.package, '/[Content_Types].xml')!.root.id;
    const ok = s.transact([
      {
        op: 'insertNode',
        part: '/[Content_Types].xml',
        parent: types,
        index: 0,
        spec: {
          ns: NS.ct,
          local: 'Default',
          attrs: [
            { ns: '', local: 'Extension', value: 'png' },
            { ns: '', local: 'ContentType', value: 'image/png' },
          ],
        },
      },
      { op: 'addPart', part: '/ppt/media/image1.png', base64: 'AAEC' },
      { op: 'removePart', part: '/ppt/media/image1.png' },
      { op: 'addPart', part: '/ppt/media/IMAGE1.png', base64: 'AAEC' },
    ]);
    expect(ok.ok).toBe(true);
  });
});

describe('K5: listeners are isolated and see changes in order', () => {
  const textOp = (pkg: PptxPackage, value: string): PptxOp => ({
    op: 'setAttribute',
    part: SLIDE,
    node: root(pkg).id,
    name: { ns: '', local: 'k' },
    value,
  });

  test('a throwing listener does not stop delivery, and transact still reports the commit', () => {
    const errors: unknown[] = [];
    const s = new PptxDocumentStore(deck(), { onListenerError: (e) => errors.push(e) });
    const got: number[] = [];
    s.subscribe(() => {
      throw new Error('boom');
    });
    s.subscribe((c) => got.push(c.toRevision));
    const r = s.transact([textOp(s.package, '1')]);
    expect(r).toEqual({ ok: true, revision: 1 });
    expect(got).toEqual([1]);
    expect(errors).toHaveLength(1);
    expect(s.canUndo).toBe(true);
  });

  test('a transaction inside a listener is delivered after the current change, in order', () => {
    const s = new PptxDocumentStore(deck());
    const seen: string[] = [];
    let once = true;
    s.subscribe((c: StoreChange) => {
      seen.push(`A:${c.fromRevision}->${c.toRevision}`);
      if (once) {
        once = false;
        s.transact([textOp(s.package, '2')]);
      }
    });
    s.subscribe((c: StoreChange) => seen.push(`B:${c.fromRevision}->${c.toRevision}`));
    s.transact([textOp(s.package, '1')]);
    expect(seen).toEqual(['A:0->1', 'B:0->1', 'A:1->2', 'B:1->2']);
  });
});

describe('K7: a moved node gets the kind a reopen would give it', () => {
  test('a shape moved out of the shape tree is generic, and back in it is a shape again', () => {
    const s = new PptxDocumentStore(deck());
    const sp = find(root(s.package), 'p:sp')!;
    const cSld = find(root(s.package), 'p:cSld')!;
    const tree = find(root(s.package), 'p:spTree')!;
    expect(sp.kind).toBe('sp');
    expect(
      s.transact([{ op: 'moveNode', part: SLIDE, node: sp.id, parent: cSld.id, index: 0 }]).ok
    ).toBe(true);
    const moved = nodeById(xmlPartOf(s.package, SLIDE)!, sp.id) as Element;
    expect(moved.kind).toBe('generic');
    const reopened = find(find(root(reopen(s.package)), 'p:cSld')!, 'p:sp')!;
    expect(reopened.kind).toBe(moved.kind);
    // Descendants are typed again too: the text body inside is no longer under a shape.
    expect(
      s.transact([{ op: 'moveNode', part: SLIDE, node: sp.id, parent: tree.id, index: 0 }]).ok
    ).toBe(true);
    expect((nodeById(xmlPartOf(s.package, SLIDE)!, sp.id) as Element).kind).toBe('sp');
  });

  test('a shape inserted inside a markup-compatibility wrapper is a shape, as on reopen', () => {
    const s = new PptxDocumentStore(deck());
    const tree = find(root(s.package), 'p:spTree')!;
    const before = xmlPartOf(s.package, SLIDE)!.nextId;
    const spec: NodeSpec = {
      ns: NS.mc,
      local: 'AlternateContent',
      prefix: 'mc',
      children: [
        {
          ns: NS.mc,
          local: 'Fallback',
          prefix: 'mc',
          children: [{ ns: NS.p, local: 'sp', prefix: 'p' }],
        },
      ],
    };
    expect(
      s.transact([{ op: 'insertNode', part: SLIDE, parent: tree.id, index: 0, spec }]).ok
    ).toBe(true);
    expect((nodeById(xmlPartOf(s.package, SLIDE)!, `n${before + 2}`) as Element).kind).toBe('sp');
    const inner = reopen(s.package);
    const fallback = find(find(root(inner), 'p:spTree')!, 'mc:AlternateContent')!;
    expect(find(fallback, 'p:sp')!.kind).toBe('sp');
  });
});

describe('K8: node ids are not reused after undo', () => {
  test('an insert after an undo mints a new id', () => {
    const s = new PptxDocumentStore(deck());
    const top = root(s.package).id;
    const ins = (local: string) =>
      s.transact([
        { op: 'insertNode', part: SLIDE, parent: top, index: 0, spec: { ns: 'urn:x', local } },
      ]);
    ins('first');
    const firstId = root(s.package).children[0].id;
    s.undo();
    ins('second');
    const second = root(s.package).children[0];
    expect(second.type === 'element' && second.local).toBe('second');
    expect(second.id).not.toBe(firstId);
    // Redo is gone, and a further undo and insert also stays clear of both.
    s.undo();
    ins('third');
    expect([firstId, second.id]).not.toContain(root(s.package).children[0].id);
  });

  test('undo still restores the same content and selection-relevant ids', () => {
    const s = new PptxDocumentStore(deck());
    const sp = find(root(s.package), 'p:sp')!;
    s.transact([{ op: 'removeNode', part: SLIDE, node: sp.id }]);
    s.undo();
    expect(find(root(s.package), 'p:sp')!.id).toBe(sp.id);
  });
});

describe('malformed op shapes are typed rejections, never a TypeError', () => {
  test.each([
    ['an unknown op kind', { op: 'bogus', part: SLIDE }],
    ['a null op', null],
    ['setAttribute without a name', { op: 'setAttribute', part: SLIDE, node: 'n1', value: 'x' }],
    ['insertNode without a spec', { op: 'insertNode', part: SLIDE, parent: 'n1', index: 0 }],
    [
      'spec attrs that are not an array',
      {
        op: 'insertNode',
        part: SLIDE,
        parent: 'n1',
        index: 0,
        spec: { ns: '', local: 'e', attrs: 5 },
      },
    ],
    ['a part that is not a string', { op: 'removePart', part: 7 }],
    [
      'an index that is not a number',
      { op: 'moveNode', part: SLIDE, node: 'n2', parent: 'n1', index: '0' },
    ],
    ['an op kind inherited from Object', { op: 'constructor', part: SLIDE }],
  ])('%s', (_label, op) => {
    const pkg = deck();
    const r = applyOps(pkg, [op as PptxOp]);
    expect(r.ok ? null : r.rejection.code).toBe('malformed-op');
  });

  test('ops that are not an array', () => {
    const r = applyOps(deck(), 'x' as unknown as PptxOp[]);
    expect(r.ok ? null : r.rejection.code).toBe('malformed-op');
  });
});

describe('P3: one id index per transaction', () => {
  // Work counts, not wall-clock time: a shared CI runner makes any time limit flaky. Rebuilding
  // the index per op walked the whole slide once for each op (2.6 s for 1,000 ops). The counters
  // see id index work: builds, the transaction's copy, and every write to the kept index. A walk
  // of the whole tree that does not touch the index, or a copy read out of it, is outside them,
  // and only the per-test timeout stops it.
  const indexWork = nodeIndexTestRecorder();
  const treeWork = workingTreeTestRecorder();
  const resetWork = () => {
    indexWork.reset();
    treeWork.reset();
  };
  const work = () => ({
    completeBuilds: indexWork.completeBuilds,
    completeVisits: indexWork.completeVisits,
    opens: treeWork.opens,
    subtreeVisits: treeWork.subtreeVisits,
    indexWrites: treeWork.indexWrites,
  });

  /** The ids whose kept index entry differs from a new walk of the part's tree. */
  const staleEntries = (part: XmlPart): NodeId[] => {
    const kept = nodeIndex(part);
    // A shallow copy of the root has no cached index, so this walks the whole tree again. Its
    // root entry holds the copy, so the kept root entry must hold the part's own root.
    const rebuilt = nodeIndex({ ...part, root: { ...part.root } });
    const stale = [...kept.keys()].filter((id) => !rebuilt.has(id));
    for (const [id, loc] of rebuilt) {
      const k = kept.get(id);
      const sameNode = k?.node === (loc.parent === null ? part.root : loc.node);
      if (k?.parent !== loc.parent || k?.index !== loc.index || !sameNode) stale.push(id);
    }
    return stale;
  };

  const subtreeSize = (node: Node): number =>
    node.type === 'element' ? node.children.reduce((n, kid) => n + subtreeSize(kid), 1) : 1;

  // One slide part of 60,003 nodes, parsed once before the tests that use it.
  const SHAPES = 5000;
  let base: PptxPackage;
  let big: XmlPart;
  let slideIndex: ReturnType<typeof nodeIndex>;
  beforeAll(() => {
    base = deck();
    const sp =
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="x"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:t>hi</a:t></a:r></a:p></p:txBody></p:sp>';
    const parsed = parseXmlPart(
      SLIDE,
      `<p:sld xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:p="${NS.p}">` +
        `<p:cSld><p:spTree>${sp.repeat(SHAPES)}</p:spTree></p:cSld></p:sld>`
    );
    if (!parsed.ok) throw new Error('fixture');
    big = parsed.part;
    slideIndex = nodeIndex(big);
  });
  const idsOf = (local: string) =>
    [...slideIndex.values()]
      .filter((l) => l.node.type === 'element' && l.node.local === local)
      .map((l) => l.node.id);
  /** The slide on a new root: same nodes and ids, no cached index yet. */
  const coldPackage = (): PptxPackage => {
    const parts = new Map(base.parts);
    const xml = { ...big, root: { ...big.root } };
    parts.set(SLIDE, { kind: 'xml', partName: SLIDE, xml });
    return { parts };
  };
  const counted = (ops: readonly PptxOp[]) => {
    const pkg = coldPackage();
    resetWork();
    const result = applyOps(pkg, ops);
    expect(result.ok ? null : result.rejection).toBeNull();
    return { part: result.ok ? xmlPartOf(result.pkg, SLIDE)! : null, work: work() };
  };
  /**
   * One part: the transaction indexes it once and never copies the index. `editWrites` are the
   * writes the edits make; the commit layers them over the index it opened.
   */
  const openedOnce = (edits: { subtreeVisits: number; editWrites: number }) => ({
    completeBuilds: 1,
    completeVisits: slideIndex.size,
    opens: 1,
    subtreeVisits: edits.subtreeVisits,
    indexWrites: edits.editWrites,
  });
  /** The number of nodes from `id` up to the root, both included. */
  const pathLength = (id: NodeId): number => {
    let length = 0;
    for (let cur: NodeId | null = id; cur !== null; cur = slideIndex.get(cur)!.parent) length += 1;
    return length;
  };

  test('1,000 ops on a 60,000-node slide index it once, then write their paths', () => {
    expect(slideIndex.size).toBeGreaterThan(60_000);
    const names = idsOf('cNvPr');
    // An attribute edit adds no node to the kept index and removes none. The first one writes
    // the entries from the edited node up to the root, which it copied: 6 for a `p:cNvPr`. The
    // transaction owns those copies, so each later edit copies its path up to its shape (3) and
    // writes into the owned `p:spTree` in place.
    const path = pathLength(names[0]);
    expect(path).toBe(6);
    const writes = (n: number): number => path + (n - 1) * (path - 3);
    const ops: PptxOp[] = names.slice(0, 1000).map((node, i) => ({
      op: 'setAttribute',
      part: SLIDE,
      node,
      name: { ns: '', local: 'name' },
      value: `n${i}`,
    }));
    // 10 ops first, so a per-op rebuild fails in well under a second.
    expect(counted(ops.slice(0, 10)).work).toEqual(
      openedOnce({ subtreeVisits: 0, editWrites: writes(10) })
    );
    const { part, work: thousand } = counted(ops);
    expect(thousand).toEqual(openedOnce({ subtreeVisits: 0, editWrites: writes(1000) }));
    if (!part) return;
    // The committed root takes over the kept index: reading it builds nothing.
    resetWork();
    nodeIndex(part);
    expect(work()).toEqual({
      completeBuilds: 0,
      completeVisits: 0,
      opens: 0,
      subtreeVisits: 0,
      indexWrites: 0,
    });
    // The ops wrote, and the index the transaction kept is exact.
    const last = nodeById(part, names[999]) as Element;
    expect(last.attrs.find((a) => a.local === 'name')?.value).toBe('n999');
    expect(staleEntries(part)).toEqual([]);
    // The store's transactions take the same path.
    const store = new PptxDocumentStore(coldPackage());
    resetWork();
    expect(store.transact(ops).ok).toBe(true);
    expect(work()).toEqual(openedOnce({ subtreeVisits: 0, editWrites: writes(1000) }));
  });

  test('inserts, moves, removals, and text edits on a 60,000-node slide keep that one index', () => {
    const [tree] = idsOf('spTree');
    const shapes = idsOf('sp');
    const texts = idsOf('t');
    const shape = subtreeSize(slideIndex.get(shapes[1])!.node);
    const { part, work: done } = counted([
      {
        op: 'insertNode',
        part: SLIDE,
        parent: tree,
        index: 0,
        spec: { ns: NS.p, local: 'sp', prefix: 'p', children: [{ text: 'x' }] },
      },
      { op: 'moveNode', part: SLIDE, node: shapes[1], parent: tree, index: 0 },
      { op: 'setText', part: SLIDE, node: texts[2], text: 'changed' },
      { op: 'removeNode', part: SLIDE, node: shapes[3] },
    ]);
    // The edits visit only the subtrees that come or go: the new shape and its text (2), the
    // moved shape out and back in (2 × 12), the old and the new text (2), the removed shape (12).
    // The first edit writes those nodes and its path to the root (3 from `p:spTree`). The
    // transaction owns that copy of `p:spTree` and above, so later edits write into it in place:
    // a later edit writes no path above `p:spTree`, and 5 entries from `a:t` up to its shape.
    // The siblings after an insert or a removal learn their new places once, when one of them
    // is read next or when the transaction ends. These are today's exact counts: a cheaper
    // edit lowers them, and then the formulas change.
    expect(shape).toBe(12);
    const subtreeVisits = 2 + 2 * shape + 2 + shape;
    // Insert at 0: the new shape and its text, the path. The shapes after it shift later.
    const insert = 2 + 3;
    // Move: reading the moved shape updates every shape from the insert place on (all 5001),
    // then out (its subtree) and in at 0 (its subtree), both in place in `p:spTree`.
    const move = SHAPES + 1 + shape + shape;
    // Text: its path reads the shapes from the move's insert place on (all 5001), then the two
    // text nodes, the path up to its shape.
    const text = SHAPES + 1 + 1 + 1 + 5;
    // Remove at position 4: its subtree, in place in `p:spTree`.
    const remove = shape;
    // The end: the shapes after the removed one learn their places.
    const end = SHAPES - 4;
    const editWrites = insert + move + text + remove + end;
    expect(done).toEqual(openedOnce({ subtreeVisits, editWrites }));
    expect(part && staleEntries(part)).toEqual([]);
  });

  test('the kept index stays exact through inserts, moves, removals, and text edits', () => {
    const s = new PptxDocumentStore(deck());
    const tree = find(root(s.package), 'p:spTree')!;
    const sp = find(root(s.package), 'p:sp')!;
    const t0 = xmlPartOf(s.package, SLIDE)!.nextId;
    const r = s.transact([
      {
        op: 'insertNode',
        part: SLIDE,
        parent: tree.id,
        index: 0,
        spec: { ns: NS.p, local: 'sp', prefix: 'p', children: [{ text: 'x' }] },
      },
      { op: 'moveNode', part: SLIDE, node: sp.id, parent: tree.id, index: 0 },
      { op: 'setText', part: SLIDE, node: `n${t0}`, text: 'changed' },
      {
        op: 'insertNode',
        part: SLIDE,
        parent: `n${t0}`,
        index: 0,
        spec: { ns: 'urn:x', local: 'q' },
      },
      { op: 'removeNode', part: SLIDE, node: find(root(s.package), 'p:nvGrpSpPr')!.id },
    ]);
    expect(r.ok).toBe(true);
    expect(staleEntries(xmlPartOf(s.package, SLIDE)!)).toEqual([]);
  });
});

describe('the serializer never rebinds a prefix it is using', () => {
  const single = (xml: string): PptxPackage => {
    const t = parseXmlPart('/x.xml', xml);
    if (!t.ok) throw new Error('fixture');
    return { parts: new Map([['/x.xml', { kind: 'xml', partName: '/x.xml', xml: t.part }]]) };
  };
  const roundTrip = (pkg: PptxPackage) => {
    const xml = serializeXmlPart(xmlPartOf(pkg, '/x.xml')!);
    const again = parseXmlPart('/x.xml', xml);
    if (!again.ok) throw new Error(`${again.reason}: ${xml}`);
    return again.part.root;
  };

  test('an attribute whose preferred prefix is the element’s own prefix', () => {
    const pkg = single('<a:r xmlns:a="urn:A"><a:k/></a:r>');
    const r = applyOps(pkg, [
      {
        op: 'setAttribute',
        part: '/x.xml',
        node: root2(pkg).id,
        name: { ns: 'urn:x', local: 'q', prefix: 'a' },
        value: '1',
      },
    ]);
    if (!r.ok) throw new Error('op');
    const back = roundTrip(r.pkg);
    expect(back.ns).toBe('urn:A');
    expect(back.attrs.map((a) => [a.ns, a.local])).toEqual([['urn:x', 'q']]);
  });

  test('a generated prefix never collides with one already in scope', () => {
    const pkg = single('<r xmlns:ns2="urn:B" ns2:z="0"><k/></r>');
    const r = applyOps(pkg, [
      {
        op: 'setAttribute',
        part: '/x.xml',
        node: root2(pkg).id,
        name: { ns: 'urn:zzz', local: 'q' },
        value: '1',
      },
    ]);
    if (!r.ok) throw new Error('op');
    const back = roundTrip(r.pkg);
    expect(back.attrs.map((a) => [a.ns, a.local]).sort()).toEqual([
      ['urn:B', 'z'],
      ['urn:zzz', 'q'],
    ]);
  });

  test('an element in no namespace under a default namespace keeps no namespace', () => {
    const pkg = single('<r xmlns="urn:D"><k/></r>');
    const r = applyOps(pkg, [
      {
        op: 'insertNode',
        part: '/x.xml',
        parent: root2(pkg).id,
        index: 0,
        spec: { ns: '', local: 'plain' },
      },
    ]);
    if (!r.ok) throw new Error('op');
    const back = roundTrip(r.pkg);
    const kids = back.children.map((c) => (c.type === 'element' ? [c.ns, c.local] : null));
    expect(kids).toEqual([
      ['', 'plain'],
      ['urn:D', 'k'],
    ]);
  });
});

function root2(pkg: PptxPackage): Element {
  return xmlPartOf(pkg, '/x.xml')!.root;
}
