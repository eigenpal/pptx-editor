// `diffPackages`: the ops that turn the open package into a returned file. After the ops, every
// part equals the file's by canonical form, unchanged parts keep their identity, and a small
// edit is a small set of node ops.

import { describe, expect, test } from 'bun:test';
import { partNameKey } from '@eigenpal/ooxml-core';
import {
  buildPptx,
  para,
  run,
  shape,
  type FixtureSlide,
} from '../../../../../scripts/fixtures/builder.ts';
import {
  canonicalForm,
  diffPackages,
  PptxDocumentStore,
  readPptx,
  writePptx,
  type PptxPackage,
} from '../index.ts';

const box = (id: number, text: string, x = 100_000) =>
  shape({ id, x, y: 100_000, cx: 2_000_000, cy: 500_000, text: para(run(text)) });

const deck = (slides: readonly FixtureSlide[], media?: Parameters<typeof buildPptx>[0]['media']) =>
  read(buildPptx({ slides, ...(media ? { media } : {}) }));

function read(bytes: Uint8Array): PptxPackage {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error(`fixture did not open: ${r.error.reason}`);
  return r.pkg;
}

/** Apply the difference in a store and check every part against the target. */
function applyAndCheck(open: PptxPackage, next: PptxPackage) {
  const diff = diffPackages(open, next);
  if (!diff.ok) throw new Error(diff.reason);
  const store = new PptxDocumentStore(open);
  const r = store.transact(diff.ops, { exact: true });
  if (!r.ok) throw new Error(JSON.stringify(r.rejection));
  const got = store.package;
  const keys = (pkg: PptxPackage) => [...pkg.parts.keys()].map(partNameKey).sort();
  expect(keys(got)).toEqual(keys(next));
  for (const [name, part] of next.parts) {
    const mine = got.parts.get(name);
    expect(mine?.kind).toBe(part.kind);
    if (part.kind === 'xml' && mine?.kind === 'xml')
      expect(canonicalForm(mine.xml.root)).toBe(canonicalForm(part.xml.root));
    if (part.kind === 'binary' && mine?.kind === 'binary') expect(mine.bytes).toEqual(part.bytes);
  }
  // A save of the result reopens to the same content as a save of the file.
  const saved = read(writePptx(got));
  const reference = read(writePptx(next));
  for (const [name, part] of reference.parts) {
    const mine = saved.parts.get(name);
    if (part.kind === 'xml' && mine?.kind === 'xml')
      expect(canonicalForm(mine.xml.root)).toBe(canonicalForm(part.xml.root));
  }
  return { diff, store, got };
}

describe('diffPackages', () => {
  test('the same file has no difference', () => {
    const slides = [{ shapes: box(2, 'Hello') }];
    const diff = diffPackages(deck(slides), deck(slides));
    expect(diff.ok && diff.ops).toEqual([]);
  });

  test('one changed word is one text op in one part', () => {
    const open = deck([{ shapes: box(2, 'Hello') }, { shapes: box(2, 'Other') }]);
    const next = deck([{ shapes: box(2, 'Hello world') }, { shapes: box(2, 'Other') }]);
    const { diff, got } = applyAndCheck(open, next);
    expect(diff.ok && diff.ops.map((op) => op.op)).toEqual(['setText']);
    // The unchanged slide keeps its part object: nothing reads or paints it again.
    expect(got.parts.get('/ppt/slides/slide2.xml')).toBe(open.parts.get('/ppt/slides/slide2.xml'));
  });

  test('moved, added, removed, and restyled shapes', () => {
    const open = deck([{ shapes: box(2, 'A') + box(3, 'B') + box(4, 'C') }]);
    const next = deck([{ shapes: box(4, 'C', 900_000) + box(2, 'A') + box(5, 'New') }]);
    const { diff } = applyAndCheck(open, next);
    const kinds = diff.ok ? diff.ops.map((op) => op.op) : [];
    expect(kinds).toContain('moveNode');
    expect(kinds).toContain('insertNode');
    expect(kinds).toContain('removeNode');
    expect(kinds).toContain('setAttribute');
    expect(kinds).not.toContain('addPart');
  });

  test('added and removed slides, media, and notes', () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3]);
    const open = deck([{ shapes: box(2, 'One') }, { shapes: box(2, 'Two') }]);
    const next = deck(
      [
        { shapes: box(2, 'One'), notes: 'Speaker notes' },
        { shapes: box(2, 'Three') },
        { shapes: box(2, 'Four') },
      ],
      [{ name: 'image1.png', bytes: png, contentType: 'image/png' }]
    );
    applyAndCheck(open, next);
    applyAndCheck(next, open);
  });

  test('a returned file with another main part is refused', () => {
    const open = deck([{ shapes: box(2, 'A') }]);
    const parts = new Map(open.parts);
    const rels = parts.get('/_rels/.rels');
    if (rels?.kind !== 'xml') throw new Error('no package relationships');
    const xml = canonicalForm(rels.xml.root);
    expect(xml).toContain('presentation.xml');
    parts.delete('/_rels/.rels');
    const diff = diffPackages(open, { parts });
    expect(diff.ok).toBe(false);
    if (!diff.ok) expect(diff.reason).toBe('different-main-part');
  });

  test('undo of the applied difference restores the open package', () => {
    const open = deck([{ shapes: box(2, 'Before') }]);
    const next = deck([{ shapes: box(2, 'After') + box(3, 'More') }, { shapes: box(2, 'X') }]);
    const { store } = applyAndCheck(open, next);
    expect(store.undo().ok).toBe(true);
    expect([...store.package.parts.keys()]).toEqual([...open.parts.keys()]);
    for (const [name, part] of open.parts) {
      const back = store.package.parts.get(name);
      if (part.kind === 'xml' && back?.kind === 'xml')
        expect(canonicalForm(back.xml.root)).toBe(canonicalForm(part.xml.root));
      else expect(back?.kind).toBe(part.kind);
    }
  });

  test('a large rewrite of one part replaces that part whole', () => {
    const many = (prefix: string) =>
      Array.from({ length: 4100 }, (_, i) => box(i + 2, `${prefix}${i}`)).join('');
    const open = deck([{ shapes: many('a') }]);
    const next = deck([{ shapes: many('b') }]);
    const { diff } = applyAndCheck(open, next);
    expect(diff.ok && diff.ops.some((op) => op.op === 'addPart')).toBe(true);
  });

  test('a namespace declaration that changes below the root replaces the part whole', () => {
    const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
    const choice = (prefix: string, uri: string) =>
      `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:${prefix}="${uri}" ` +
      `Requires="${prefix}"><p:sp><p:nvSpPr><p:cNvPr id="9" name="X"/><p:cNvSpPr/><p:nvPr/>` +
      `</p:nvSpPr><p:spPr/></p:sp></mc:Choice><mc:Fallback/></mc:AlternateContent>`;
    const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
    const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
    const open = deck([{ shapes: box(2, 'A') + choice('p14', P14) }]);
    const next = deck([{ shapes: box(2, 'A') + choice('a14', A14) }]);
    const { store } = applyAndCheck(open, next);
    const written = read(writePptx(store.package)).parts.get('/ppt/slides/slide1.xml');
    if (written?.kind !== 'xml') throw new Error('no slide');
    type El = typeof written.xml.root;
    const all = (el: El): El[] => [
      el,
      ...el.children.flatMap((c) => (c.type === 'element' ? all(c) : [])),
    ];
    const ac = all(written.xml.root).find((el) => el.local === 'Choice');
    expect(ac?.nsDecls.map((d) => `${d.prefix}=${d.uri}`)).toEqual([`a14=${A14}`]);
  });

  test('30000 changed paragraphs take linear time and replace the part whole', () => {
    const pair = (n: number) => {
      const body = (tag: string) =>
        shape({
          id: 2,
          x: 1,
          y: 1,
          cx: 10,
          cy: 10,
          text: Array.from({ length: n }, (_, i) => para(run(`${tag}${i}`))).join(''),
        });
      return [deck([{ shapes: body('a') }]), deck([{ shapes: body('b') }])] as const;
    };
    // The best of three runs, so a slow runner or a collection pause does not decide it.
    const fastest = (open: PptxPackage, next: PptxPackage) => {
      let best = Infinity;
      for (let i = 0; i < 3; i++) {
        const start = performance.now();
        diffPackages(open, next);
        best = Math.min(best, performance.now() - start);
      }
      return best;
    };
    const [smallOpen, smallNext] = pair(3_000);
    const [open, next] = pair(30_000);
    const diff = diffPackages(open, next);
    if (!diff.ok) throw new Error(diff.reason);
    expect(diff.ops.some((op) => op.op === 'addPart')).toBe(true);
    // Ten times the paragraphs take about ten times as long. A quadratic name match took
    // about 100 times as long. A ratio does not depend on the speed of the machine.
    expect(fastest(open, next) / Math.max(fastest(smallOpen, smallNext), 1)).toBeLessThan(25);
  });

  test('children of one name pair in order by name when their content differs', () => {
    const texts = (tag: string) => ['x', 'y', 'z'].map((t) => para(run(`${tag}${t}`))).join('');
    const body = (tag: string) => shape({ id: 2, x: 1, y: 1, cx: 10, cy: 10, text: texts(tag) });
    const open = deck([{ shapes: body('a') }]);
    const next = deck([{ shapes: body('b') }]);
    const { diff } = applyAndCheck(open, next);
    if (!diff.ok) throw new Error('no difference');
    expect(diff.ops.every((op) => op.op === 'setText')).toBe(true);
    expect(diff.ops.length).toBe(3);
  });
});
