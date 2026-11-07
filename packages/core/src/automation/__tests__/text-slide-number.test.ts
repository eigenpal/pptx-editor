// A `slidenum` field is as long as the slide number it shows, so a text edit reads the slide's
// number. One batch (one automation `context.sync()`) can add or move slides between two text
// edits: the edits after it read the new numbers.

import { expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  find,
  presentationInfo,
  readPptx,
  xmlPartOf,
  type PptxPackage,
} from '../../store/index.ts';
import { OpBatch } from '../commands/op-batch.ts';
import { addSlide, moveSlide } from '../commands/slide-commands.ts';
import { insertBreak, insertText, slideNumberOf } from '../commands/text-model.ts';

const IN = 914_400;
const FIELD =
  '<a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/>' +
  '<a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p>';

/** Twelve slides, each with a slide number field. */
function deck(): PptxPackage {
  const bytes = buildPptx({
    slides: Array.from({ length: 12 }, (_, i) => ({
      shapes: shape({ id: 2, name: `N${i}`, x: IN, y: IN, cx: 4 * IN, cy: IN, text: FIELD }),
    })),
  });
  const read = readPptx(bytes);
  if (!read.ok) throw new Error('deck does not read');
  return read.pkg;
}

const txBodyOf = (batch: OpBatch, part: string): string =>
  find(xmlPartOf(batch.pkg, part)!.root, 'p:txBody')!.id;

test('a text edit after a slide moves in the same batch reads the new slide number', () => {
  const pkg = deck();
  const parts = presentationInfo(pkg).slides.map((s) => s.partName);
  const batch = new OpBatch(pkg);
  const first = parts[0]!;
  const last = parts[11]!;
  // A text edit reads the numbers: the last slide shows 12.
  insertText(batch, first, txBodyOf(batch, first), { paragraph: 0, offset: 1 }, 'a');
  expect(slideNumberOf(batch, last)).toBe(12);
  moveSlide(batch, 11, 0);
  expect(slideNumberOf(batch, last)).toBe(1);
  expect(slideNumberOf(batch, first)).toBe(2);
  // After the field, which shows "1": the caret moves past the typed text only.
  const at = insertText(batch, last, txBodyOf(batch, last), { paragraph: 0, offset: 1 }, 'x');
  expect(at).toEqual({ paragraph: 0, offset: 2 });
});

test('a text edit on a slide added in the same batch reads its slide number', () => {
  const pkg = deck();
  const batch = new OpBatch(pkg);
  const first = presentationInfo(pkg).slides[0]!.partName;
  insertText(batch, first, txBodyOf(batch, first), { paragraph: 0, offset: 1 }, 'a');
  const index = addSlide(batch, 0);
  const added = presentationInfo(batch.pkg).slides[index]!.partName;
  expect(slideNumberOf(batch, added)).toBe(index + 1);
  expect(slideNumberOf(batch, presentationInfo(batch.pkg).slides[11]!.partName)).toBe(12);
});

/** The fastest of five runs of 1,000 line breaks on a slide with `extra` other shapes, in ms. */
function breaksMs(extra: number): number {
  let shapes = shape({ id: 2, name: 'Body', x: IN, y: IN, cx: IN, cy: IN, text: para(run('ab')) });
  for (let i = 0; i < extra; i += 1)
    shapes += shape({ id: 10 + i, x: i * 10, y: 0, cx: IN, cy: IN, text: para(run(`s${i}`)) });
  const read = readPptx(buildPptx({ slides: [{ shapes }] }));
  if (!read.ok) throw new Error('deck does not read');
  const part = presentationInfo(read.pkg).slides[0]!.partName;
  let best = Infinity;
  for (let k = 0; k < 5; k += 1) {
    const batch = new OpBatch(read.pkg);
    const body = txBodyOf(batch, part);
    const t0 = performance.now();
    let at = { paragraph: 0, offset: 1 };
    for (let i = 0; i < 1_000; i += 1) at = insertBreak(batch, part, body, at);
    best = Math.min(best, performance.now() - t0);
  }
  return best;
}

// Each line break reads its slide's number. A read of `batch.pkg` closes the open transaction,
// and the next op opens one on a copy of the slide part, so the cost grew with the slide: 15
// times as long beside 30,000 shapes as beside 100. A limit of 5 leaves room for a busy machine.
test('line breaks beside 30,000 shapes cost less than five times what they cost beside 100', () => {
  const small = breaksMs(100);
  const large = breaksMs(30_000);
  expect(large).toBeLessThan(small * 5);
});
