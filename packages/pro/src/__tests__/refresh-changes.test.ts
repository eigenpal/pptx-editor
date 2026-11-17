// Computed document refresh changes (review/refresh-changes.ts): what changed from a captured
// presentation to a returned one, with stable ids and locations the editor accepts, and
// `comparePresentations` on its own.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createDocumentRefresh, createPptxEditor } from '@pptx-editor.dev/core';
import { readPptx, type PptxPackage } from '@pptx-editor.dev/core/store';
import {
  buildPptx,
  para,
  run,
  shape,
  type FixtureSlide,
} from '../../../../scripts/fixtures/builder.ts';
import { comparePresentations, reviewModule } from '../index.ts';
import { computeRefreshChanges } from '../review/refresh-changes.ts';

const box = (id: number, text: string, x = 400_000) =>
  shape({ id, x, y: 400_000, cx: 4_000_000, cy: 800_000, text: para(run(text)) });
const bytesOf = (
  slides: readonly FixtureSlide[],
  layouts?: Parameters<typeof buildPptx>[0]['layouts']
) => buildPptx({ slides, ...(layouts ? { layouts } : {}) });
const pkg = (bytes: Uint8Array): PptxPackage => {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('fixture did not open');
  return r.pkg;
};
const compare = (a: readonly FixtureSlide[], b: readonly FixtureSlide[]) =>
  computeRefreshChanges(pkg(bytesOf(a)), pkg(bytesOf(b))).changes;

describe('computed refresh changes', () => {
  test('the same presentation has no changes', () => {
    const slides = [{ shapes: box(2, 'A') }, { shapes: box(2, 'B') }];
    expect(compare(slides, slides)).toEqual([]);
  });

  test('a text edit is one text change located at the new text', () => {
    const changes = compare(
      [{ shapes: box(2, 'Hello there') }],
      [{ shapes: box(2, 'Hello world') }]
    );
    expect(changes).toEqual([
      {
        id: expect.stringMatching(/^text:/),
        location: {
          slideIndex: 0,
          shapeId: 2,
          paragraphIndex: 0,
          start: 6,
          end: 11,
          text: 'world',
        },
        computed: { scope: 'text', action: 'changed', aspects: ['text'] },
      },
    ]);
  });

  test('a moved shape reports its aspect; added and removed shapes report themselves', () => {
    const changes = compare(
      [{ shapes: box(2, 'A') + box(3, 'B') }],
      [{ shapes: box(2, 'A', 900_000) + box(4, 'New') }]
    );
    const by = changes.map((c) => [c.computed.scope, c.computed.action, c.computed.aspects ?? []]);
    expect(by).toContainEqual(['shape', 'changed', ['position']]);
    expect(by).toContainEqual(['shape', 'added', []]);
    expect(by).toContainEqual(['shape', 'removed', []]);
  });

  test('added, removed, and moved slides; an insertion does not move the slides after it', () => {
    const a = [{ shapes: box(2, 'One') }, { shapes: box(2, 'Two') }, { shapes: box(2, 'Three') }];
    const inserted = [a[0]!, { shapes: box(2, 'New') }, a[1]!, a[2]!];
    const ins = compare(a, inserted);
    expect(ins.map((c) => c.computed.action)).toEqual(['added']);
    expect(ins[0]!.location).toEqual({ slideIndex: 1 });
    const removed = compare(a, [a[0]!, a[2]!]);
    expect(removed).toHaveLength(1);
    expect(removed[0]).toMatchObject({
      unavailableReason: 'deleted',
      computed: { scope: 'slide', action: 'removed', previousSlideIndex: 1 },
    });
  });

  test('a new slide in the part of a removed slide is an addition and a removal', () => {
    const a = [
      { shapes: box(2, 'One') },
      { shapes: box(2, 'Two') },
      { shapes: box(2, 'Appendix') },
    ];
    const b = [a[0]!, a[1]!, { shapes: box(2, 'Summary') }];
    // The writer keeps slide ids: the new slide has a new id in the same part name.
    const files = unzipSync(bytesOf(b));
    const xml = strFromU8(files['ppt/presentation.xml']!);
    files['ppt/presentation.xml'] = strToU8(xml.replace('<p:sldId id="258"', '<p:sldId id="300"'));
    const changes = computeRefreshChanges(pkg(bytesOf(a)), pkg(zipSync(files))).changes;
    expect(changes.map((c) => [c.computed.scope, c.computed.action])).toEqual([
      ['slide', 'removed'],
      ['slide', 'added'],
    ]);
  });

  test('notes changes and layout changes, with the slides that show the layout', () => {
    const notes = compare(
      [{ shapes: box(2, 'A'), notes: 'Old' }],
      [{ shapes: box(2, 'A'), notes: 'New' }]
    );
    expect(notes).toEqual([
      {
        id: expect.stringMatching(/^notes:/),
        location: { slideIndex: 0, notes: true },
        computed: { scope: 'notes', action: 'changed', aspects: ['text'] },
      },
    ]);
    const layouts = (name: string) => [{ name: 'Mine', shapes: box(9, name) }];
    const slides = [
      { shapes: box(2, 'A'), layoutName: 'Mine' },
      { shapes: box(2, 'B') },
      { shapes: box(2, 'C'), layoutName: 'Mine' },
    ];
    const changes = computeRefreshChanges(
      pkg(bytesOf(slides, layouts('one'))),
      pkg(bytesOf(slides, layouts('two')))
    ).changes;
    expect(changes).toHaveLength(1);
    expect(changes[0]!.computed).toEqual({
      scope: 'layout',
      action: 'changed',
      affectedSlides: [0, 2],
    });
  });

  test('comparePresentations refuses a file that is not a presentation', () => {
    const r = comparePresentations(new Uint8Array([1, 2, 3]), bytesOf([{ shapes: '' }]));
    expect(r).toMatchObject({ ok: false, code: 'invalidDocument' });
    const ok = comparePresentations(
      bytesOf([{ shapes: box(2, 'A') }]),
      bytesOf([{ shapes: box(2, 'B') }])
    );
    expect(ok.ok && ok.changes.length).toBe(1);
  });
});

describe('reviewModule computes refresh changes', () => {
  test('a result without changes gets located, highlighted, stable changes', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, modules: [reviewModule()] });
    await editor.load(bytesOf([{ shapes: box(2, 'Hello there') }, { shapes: box(2, 'Two') }]));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const r1 = await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: bytesOf([{ shapes: box(2, 'Hello world') }, { shapes: box(2, 'Two') }]),
    });
    if (!r1.ok) throw new Error(r1.code);
    expect(r1.changeInformation).toBe('available');
    expect(r1.changes.map((c) => [c.status, c.isNew, c.computed?.scope])).toEqual([
      ['available', true, 'text'],
    ]);
    // The next cumulative result keeps the first change and adds one.
    const r2 = await refresh.applyUpdate({
      submission,
      sequence: 2,
      bytes: bytesOf([
        { shapes: box(2, 'Hello world') },
        { shapes: box(2, 'Two') },
        { shapes: box(2, 'Three') },
      ]),
      mode: 'edit',
    });
    if (!r2.ok) throw new Error(r2.code);
    expect(r2.changes.map((c) => [c.computed?.action, c.isNew])).toEqual([
      ['changed', false],
      ['added', true],
    ]);
    expect(refresh.highlightChanges({ animation: false, includePrevious: true })).toBe(2);
    refresh.finish(submission);
    editor.destroy();
  });
});
