// Text warp layout (`a:prstTxWarp`): the output budget, the pairs of a warp with a path of zero
// length, the arcs with one zero radius, and picture bullets.

import { describe, expect, test } from 'bun:test';
import { PNG } from 'pngjs';
import type { PathCommand } from '@eigenpal/ooxml-core';
import { evaluatePresetTextWarp } from '@eigenpal/ooxml-core';
import { buildPptx, para, run, shape } from '../../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../../store/index.ts';
import { resolveSlide, type ResolvedTextBody } from '../../../cascade/index.ts';
import { createFixedMeasurer, layoutSlide, type TextLayout } from '../../index.ts';
import { warpTextLayout } from '../text-warp.ts';

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const CX = 3 * IN;
const CY = 2 * IN;

function png(): Uint8Array<ArrayBuffer> {
  const image = new PNG({ width: 4, height: 4 });
  image.data.fill(200);
  return new Uint8Array(PNG.sync.write(image));
}

/** The laid-out text and the resolved text body of one warped shape. */
function lay(warp: string, text: string): { text: TextLayout; body: ResolvedTextBody } {
  const r = readPptx(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: CX,
            cy: CY,
            bodyPr: `<a:bodyPr wrap="none" rtlCol="0">${warp}<a:noAutofit/></a:bodyPr>`,
            text,
          }),
          rels: [['rIdBu', R_IMAGE, '../media/image1.png']],
        },
      ],
      media: [{ name: 'image1.png', bytes: png(), contentType: 'image/png' }],
    })
  );
  if (!r.ok) throw new Error('open failed');
  const item = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer()).items[0]!;
  if (item.kind !== 'shape' || !item.text || item.element.kind !== 'shape' || !item.element.text)
    throw new Error('no text');
  return { text: item.text, body: item.element.text };
}

const warpOf = (prst: string, gd = ''): string =>
  `<a:prstTxWarp prst="${prst}"><a:avLst>${gd}</a:avLst></a:prstTxWarp>`;

/** The largest y of the warped glyph points, in the text frame. */
function lowestPoint(t: TextLayout): number {
  let most = -Infinity;
  for (const w of t.warped ?? [])
    for (const c of w.glyphs) if (c.op !== 'Z' && c.y > most) most = c.y;
  return most;
}

describe('text warp layout', () => {
  test('a font with huge glyphs spends the budget and draws flat before it cuts every glyph', () => {
    const { text, body } = lay(warpOf('textWave1'), para(run('A'.repeat(200))));
    // A glyph of 1,999 curves with far control points: each curve cuts into 64 pieces.
    const huge: PathCommand[] = [{ op: 'M', x: 0, y: 0 }];
    for (let i = 0; i < 1_999; i += 1)
      huge.push({ op: 'C', x1: 100, y1: 100, x2: -100, y2: 100, x: (i % 2) * 0.5, y: 0.5 });
    huge.push({ op: 'Z' });
    let calls = 0;
    const measurer = {
      ...createFixedMeasurer(),
      glyphOutline: () => {
        calls += 1;
        return huge;
      },
    };
    const started = performance.now();
    const result = warpTextLayout(
      text,
      body,
      { preset: 'textWave1', adjust: new Map() },
      CX,
      CY,
      measurer
    );
    expect(result).toEqual({ flat: 'too much text' });
    // About 128,000 points a glyph against a budget of 400,000: the fourth glyph stops it.
    expect(calls).toBeLessThanOrEqual(4);
    expect(performance.now() - started).toBeLessThan(2_000);
  });

  test('a pour warp with a zero inner radius pairs the outer arc with the center point', () => {
    for (const prst of ['textArchUpPour', 'textCirclePour']) {
      const gd = '<a:gd name="adj2" fmla="val 0"/>';
      // The inner path has zero length, but it still counts: two paths make one pair.
      expect(evaluatePresetTextWarp(prst, CX, CY, new Map([['adj2', 0]]))!.length).toBe(2);
      const { text } = lay(warpOf(prst, gd), para(run('Pour')));
      expect(text.warped?.length).toBeGreaterThan(0);
      // The band narrows to the shape's center: the text reaches down to it.
      const center = CY / 2 - text.box.y;
      expect(lowestPoint(text)).toBeGreaterThan(center - CY * 0.02);
    }
  });

  test('a can warp with a zero arc height keeps straight paths', () => {
    for (const [prst, adj] of [
      ['textCanDown', 0],
      ['textCanUp', 100_000],
    ] as const) {
      const paths = evaluatePresetTextWarp(prst, CX, CY, new Map([['adj', adj]]))!;
      // Each arc of zero height is a line from the left edge to the right edge.
      for (const p of paths) {
        expect(p.commands.map((c) => c.op)).toEqual(['M', 'L']);
        const [m, l] = p.commands;
        if (m?.op !== 'M' || l?.op !== 'L') throw new Error('not a line');
        expect(Math.round(l.x - m.x)).toBe(CX);
      }
      const { text } = lay(warpOf(prst, `<a:gd name="adj" fmla="val ${adj}"/>`), para(run('Can')));
      expect(text.warped?.length).toBeGreaterThan(0);
    }
  });

  test('each pair of a warp with two pairs keeps the text size of the largest pair', () => {
    // textDeflateInflate has two pairs: one line in each. The short line grows about its center
    // to the bounds of the long line, so it does not stretch to the ends of its band.
    const { text } = lay(
      warpOf('textDeflateInflate'),
      para(run('Long line'), '<a:pPr algn="ctr"/>') + para(run('ab'), '<a:pPr algn="ctr"/>')
    );
    const span = (line: number): [number, number] => {
      let [a, b] = [Infinity, -Infinity];
      for (const w of text.warped ?? [])
        if (w.line === line)
          for (const c of w.glyphs) if (c.op !== 'Z') [a, b] = [Math.min(a, c.x), Math.max(b, c.x)];
      return [a, b];
    };
    const [l0, r0] = span(0);
    const [l1, r1] = span(1);
    // The long line spans the band. The short one keeps about its width relative to it (0.9 em
    // of ink against 4.15 em); the curved band edge moves it a little.
    expect(r0 - l0).toBeCloseTo(CX, -3);
    expect((r1 - l1) / (r0 - l0)).toBeCloseTo(0.9 / 4.15, 1);
    expect((l1 + r1) / 2).toBeCloseTo((l0 + r0) / 2, -3);
  });

  test('a picture bullet maps its corners into the warp', () => {
    const bullet =
      '<a:pPr marL="342900" indent="-342900"><a:buBlip><a:blip r:embed="rIdBu"/></a:buBlip></a:pPr>';
    const { text } = lay(warpOf('textWave1'), para(run('Item'), bullet));
    const pic = text.warped?.find((w) => w.picture)?.picture;
    expect(pic).toBeDefined();
    // A wave turns the picture: its top edge is not level.
    expect(Math.abs(pic!.across[0])).toBeGreaterThan(0);
    expect(Math.abs(pic!.across[1])).toBeGreaterThan(0);
    expect(pic!.down[1]).toBeGreaterThan(0);
  });
});
