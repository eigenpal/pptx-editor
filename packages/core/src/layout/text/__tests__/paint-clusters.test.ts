import { expect, test } from 'bun:test';
import { graphemeCount } from '@eigenpal/ooxml-core';
import { buildPptx, shape, para, run } from '../../../../../../scripts/fixtures/builder.ts';
import { readPptx, writePptx } from '../../../store/index.ts';
import { resolveSlide } from '../../../cascade/index.ts';
import { layoutSlide, createFixedMeasurer } from '../../index.ts';
import { paintClusterRanges, MAX_PAINT_CLUSTER_UNITS } from '../paint-clusters.ts';

test('ordinary graphemes keep their full shaping context', () => {
  const text = 'مرحبا👩🏽‍🚀á'.repeat(100);
  expect(paintClusterRanges(text)).toEqual([{ start: 0, end: text.length }]);
});

test('oversized clusters keep every code point without splitting surrogate pairs', () => {
  const text = 'prefix ' + 'a' + '\u{1D185}'.repeat(500) + ' suffix';
  const pieces = paintClusterRanges(text).map(({ start, end }) => text.slice(start, end));
  expect(pieces.join('')).toBe(text);
  expect(pieces.every((piece) => piece.length <= MAX_PAINT_CLUSTER_UNITS)).toBe(true);
  expect(pieces.every((piece) => !/[\uD800-\uDFFF]/u.test(piece))).toBe(true);
  expect(pieces[0]).toBe('prefix ');
  expect(pieces.at(-1)).toBe(' suffix');
});

test('bounded layout runs preserve source text, positions, and semantic caret boundaries', () => {
  const text = 'a' + '\u0301'.repeat(5000);
  const opened = readPptx(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: 0,
            y: 0,
            cx: 6_000_000,
            cy: 4_000_000,
            text: para(run(text)),
            bodyPr: '<a:bodyPr wrap="none"/>',
          }),
        },
      ],
    })
  );
  if (!opened.ok) throw new Error('open failed');
  const slide = resolveSlide(opened.pkg, 0)!;
  const layout = layoutSlide(slide, createFixedMeasurer());
  const item = layout.items[0];
  if (item.kind !== 'shape' || !item.text) throw new Error('no text');
  expect(layout.diagnostics.some((d) => d.includes('long-text-clusters'))).toBe(true);
  const line = item.text.lines[0];
  expect(line.runs.map((r) => r.text).join('')).toBe(text);
  expect(line.runs.every((r) => r.text.length <= MAX_PAINT_CLUSTER_UNITS)).toBe(true);
  expect(line.caretX).toHaveLength(text.length + 1);
  for (const r of line.runs)
    expect(r.x).toEqual(line.caretX.slice(r.start, r.start + r.text.length + 1));
  const reopened = readPptx(writePptx(opened.pkg));
  if (!reopened.ok) throw new Error('reopen failed');
  const original = resolveSlide(reopened.pkg, 0)!.elements[0];
  if (original.kind !== 'shape') throw new Error('not shape');
  expect(original.text!.paragraphs[0].runs[0].text).toBe(text);
  expect(graphemeCount(text)).toBe(1);
});
