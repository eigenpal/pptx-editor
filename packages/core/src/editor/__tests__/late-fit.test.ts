// Text that a person edits while the deck measures with the fixed measurer fits again once the
// substitutes arrive: the saved `normAutofit` values are the ones shaped text needs. The fit is
// its own undo step, and a moved shape keeps the fit that the file has.
//
// The module mock needs its own process: `bun run test` gives each file one.

import { expect, mock, test } from 'bun:test';
import * as realFonts from '@docx-editor.dev/fonts';

const original = realFonts.loadDefaultFonts;
const copy = { ...realFonts };
// The substitutes load when the test calls `release`.
let release: () => void = () => {};
let held = Promise.resolve();
const hold = (): void => void (held = new Promise<void>((resolve) => (release = resolve)));
mock.module('@docx-editor.dev/fonts', () => ({
  ...copy,
  loadDefaultFonts: (o: Parameters<typeof realFonts.loadDefaultFonts>[0] = {}) =>
    held.then(() => original(o)),
}));

const settle = async (editor: { fontMeasurement(): { resolving: boolean } }) => {
  for (let i = 0; i < 400 && editor.fontMeasurement().resolving; i += 1)
    await new Promise((r) => setTimeout(r, 5));
};

test('text edited on the fixed measurer fits again with shaped text', async () => {
  const { createPptxEditor } = await import('../index.ts');
  const { buildPptx, para, run, shape } =
    await import('../../../../../scripts/fixtures/builder.ts');
  const { xmlPartOf } = await import('../../store/index.ts');
  type Editor = ReturnType<typeof createPptxEditor>;
  const IN = 914_400;
  const deck = buildPptx({
    slides: [
      {
        shapes:
          shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 3 * IN,
            cy: 0.6 * IN,
            bodyPr: '<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>',
            text: para(run('Fit '.repeat(12).trim())),
          }) +
          // A shape nobody edits keeps the fit the file has, even one that its text does not need.
          shape({
            id: 3,
            x: IN,
            y: 3 * IN,
            cx: 3 * IN,
            cy: IN,
            bodyPr: '<a:bodyPr wrap="square"><a:normAutofit fontScale="62500"/></a:bodyPr>',
            text: para(run('Kept')),
          }),
      },
    ],
  });
  const typeInto = (editor: Editor) => {
    const item = editor.session!.layout(0)!.items.find((i) => i.kind === 'shape')!;
    const end = { paragraph: 0, offset: 0 };
    editor.setSelection({ kind: 'text', nodeId: item.element.nodeId, anchor: end, focus: end });
    for (const key of ' this text grows'.split('')) editor.input('text', key);
  };
  /** The written `fontScale` and `lnSpcReduction` of shape `n` (0 is the edited one). */
  const fit = (editor: Editor, n = 0): string => {
    const part = editor.session!.info.slides[0]!.partName;
    const xml = JSON.stringify(xmlPartOf(editor.session!.store.package, part)!.root);
    const all = [...xml.matchAll(/"local":"normAutofit".*?"attrs":(\[.*?\])/g)];
    const attrs = all[n]?.[1] ?? '[]';
    const list = JSON.parse(attrs) as { local: string; value: string }[];
    return list
      .map((a) => `${a.local}=${a.value}`)
      .sort()
      .join(' ');
  };

  hold();
  const late = createPptxEditor();
  expect((await late.load(deck)).ok).toBe(true);
  expect(late.fontMeasurement().measurer).toBe('fixed');
  // A move changes the shape's XML but fits no text: the late fit leaves it alone.
  const kept = late.session!.layout(0)!.items.filter((i) => i.kind === 'shape')[1]!;
  late.setSelection({ kind: 'shapes', nodeIds: [kept.element.nodeId] });
  expect(late.exec({ type: 'nudge', dx: IN / 10, dy: 0 }).ok).toBe(true);
  // Typing that is undone leaves the opened text: that shape keeps its fit too.
  const start = { paragraph: 0, offset: 0 };
  late.setSelection({ kind: 'text', nodeId: kept.element.nodeId, anchor: start, focus: start });
  late.input('text', 'x');
  expect(late.exec({ type: 'undo' }).ok).toBe(true);
  typeInto(late);
  const fixedFit = fit(late);
  const undoable = late.snapshot().canUndo;
  release();
  await settle(late);
  expect(late.fontMeasurement()).toEqual({ measurer: 'shaped', resolving: false });

  const early = createPptxEditor();
  expect((await early.load(deck)).ok).toBe(true);
  expect(early.fontMeasurement().measurer).toBe('shaped');
  typeInto(early);

  // The fit with estimated widths differs, and the late fit writes the shaped one.
  expect(fixedFit).not.toBe(fit(early));
  expect(fit(late)).toBe(fit(early));
  expect(fit(late, 1)).toBe('fontScale=62500');
  expect(undoable).toBe(true);
  // The fit is its own undo step: the first undo gives back the estimated fit.
  expect(late.exec({ type: 'undo' }).ok).toBe(true);
  expect(fit(late)).toBe(fixedFit);
  while (late.snapshot().canUndo) expect(late.exec({ type: 'undo' }).ok).toBe(true);
  expect(fit(late)).toBe('');
  expect(fit(late, 1)).toBe('fontScale=62500');
  late.destroy();
  early.destroy();
}, 15_000);
