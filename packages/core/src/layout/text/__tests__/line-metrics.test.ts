// The line model, calibrated on reference PDFs. The expected values follow from the rules that the
// reference PDFs show; the comments name the fixture and PDF value.

import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../../store/index.ts';
import { resolveSlide } from '../../../cascade/index.ts';
import { createFixedMeasurer, layoutSlide, type TextLayout } from '../../index.ts';
import type { FaceMetrics } from '../measurer.ts';
import {
  autofitSizePt,
  baselineOffsetPt,
  faceLinePt,
  lineHeightPt,
  paragraphSpacingPt,
  roundToWholePoint,
  WHOLE_POINT_BASELINES,
} from '../line-metrics.ts';

const EMU_PER_PT = 12_700;
const face = (ascent: number, descent: number): FaceMetrics => ({
  ascent,
  descent,
  lineGap: 0,
  underlinePosition: -0.1,
  underlineThickness: 0.05,
});
/** Calibri's horizontal metrics (ascent share 0.78). */
const CALIBRI = face(0.9521, 0.2686);

describe('line pitch', () => {
  test('percent spacing advances pct × 1.2 × size', () => {
    expect(lineHeightPt(18, { pct: 0.9 }, 0)).toBeCloseTo(19.44, 6);
    expect(lineHeightPt(18, { pct: 1.5 }, 0)).toBeCloseTo(32.4, 6);
  });

  test('lnSpcReduction is subtracted from the percentage', () => {
    expect(lineHeightPt(20, { pct: 1.1 }, 0.1)).toBeCloseTo(24, 6);
    expect(lineHeightPt(26, { pct: 0.9 }, 0.2)).toBeCloseTo(21.84, 6);
  });

  test('exact spacing rounds to a whole point', () => {
    expect(lineHeightPt(13, { pts: 1680 }, 0)).toBe(17);
    expect(lineHeightPt(14, { pts: 1560 }, 0)).toBe(16);
    expect(lineHeightPt(28, { pts: 4120 }, 0)).toBe(41);
  });

  test('paragraph spacing in points rounds to a whole point', () => {
    expect(paragraphSpacingPt({ pts: 640 }, 28)).toBe(6);
    expect(paragraphSpacingPt({ pts: 750 }, 23)).toBe(8);
    expect(paragraphSpacingPt({ pct: 0.2 }, 20)).toBeCloseTo(4.8, 6);
  });
});

describe('baseline in the line box', () => {
  test('at 100 % the face ascent share of the single line', () => {
    expect(baselineOffsetPt(18, 21.6, CALIBRI)).toBeCloseTo(16.848, 2);
  });

  test('below 100 % the descent stays: the space comes from above', () => {
    // vertical-columns: 18 pt Calibri at 90 % → first baseline 14.688 (PDF: 15 after rounding).
    expect(baselineOffsetPt(18, 19.44, CALIBRI)).toBeCloseTo(14.688, 2);
  });

  test('the descent is never more than 25 % of the box', () => {
    // autofit: 18 pt at 90 % − 20 % → 0.75 × 15.12 = 11.34 (PDF: 11).
    expect(baselineOffsetPt(18, 15.12, CALIBRI)).toBeCloseTo(11.34, 6);
  });

  test('the deprecated spacing argument is ignored: the line height decides', () => {
    expect(baselineOffsetPt(18, 32.4, CALIBRI, { pct: 1.5 })).toBe(
      baselineOffsetPt(18, 32.4, CALIBRI)
    );
    expect(baselineOffsetPt(18, 19.44, CALIBRI, { pts: 1944 })).toBe(
      baselineOffsetPt(18, 19.44, CALIBRI)
    );
  });

  test('above the single line the baseline sits at 75 % of the box for a face of share 0.75 or more', () => {
    expect(baselineOffsetPt(18, 32.4, CALIBRI)).toBeCloseTo(24.3, 6);
    expect(baselineOffsetPt(18, 30, face(0.9053, 0.2119))).toBeCloseTo(22.5, 6);
  });

  // Cordia New: win 3678 and 2000 of 4096 (share 0.648). Measured on 178 single and 51 1.5-line
  // probe lines: the baseline sits at the face's share, never raised to 75 % of the box.
  const CORDIA = face(3678 / 4096, 2000 / 4096);
  const share = 3678 / 5678;

  test('a face of share under 0.75 keeps its own share at single spacing', () => {
    expect(baselineOffsetPt(18, 21.6, CORDIA)).toBeCloseTo(21.6 * share, 6);
    expect(baselineOffsetPt(18, 21.6, CORDIA)).toBeLessThan(0.75 * 21.6);
  });

  test('above the single line the extra space goes 75 % above the baseline', () => {
    // 1.5 lines: 0.75 × (32.4 − 21.6) + 21.6 × share, below 75 % of the box.
    expect(baselineOffsetPt(18, 32.4, CORDIA)).toBeCloseTo(0.75 * 10.8 + 21.6 * share, 6);
    expect(baselineOffsetPt(18, 32.4, CORDIA)).toBeLessThan(0.75 * 32.4);
  });
});

describe('whole points', () => {
  test('autofit font scale gives whole-point sizes, half up', () => {
    expect(autofitSizePt(28, 0.625)).toBe(18);
    expect(autofitSizePt(20, 0.925)).toBe(19);
    expect(autofitSizePt(18, 0.925)).toBe(17);
    expect(autofitSizePt(20, 0.475)).toBe(10);
    expect(autofitSizePt(10.5, 1)).toBe(10.5);
  });

  test('roundToWholePoint rounds baselines to whole points', () => {
    expect(roundToWholePoint(14.688 * EMU_PER_PT)).toBe(15 * EMU_PER_PT);
    expect(roundToWholePoint(53.568 * EMU_PER_PT)).toBe(54 * EMU_PER_PT);
  });
});

const shape = (body: string, bodyPr: string): string =>
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="4000000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
  `<p:txBody>${bodyPr}<a:lstStyle/>${body}</p:txBody></p:sp>`;

function lay(body: string, bodyPr = '<a:bodyPr tIns="0"/>'): TextLayout {
  const r = readPptx(buildPptx({ slides: [{ shapes: shape(body, bodyPr), layout: 'blank' }] }));
  if (!r.ok) throw new Error('open failed');
  const item = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer()).items[0];
  if (item.kind !== 'shape' || !item.text) throw new Error('no text');
  return item.text;
}

const para = (text: string, pPr = '', size = 1800): string =>
  `<a:p>${pPr}<a:r><a:rPr lang="en-US" sz="${size}"/><a:t>${text}</a:t></a:r></a:p>`;
const baselinesPt = (t: TextLayout): number[] => t.lines.map((l) => l.baseline / EMU_PER_PT);

describe('layout', () => {
  test('each baseline rounds to a whole point from the top of the text', () => {
    expect(WHOLE_POINT_BASELINES).toBe(true);
    const t = lay(
      ['a', 'b', 'c', 'd']
        .map((x) => para(x, '<a:pPr><a:lnSpc><a:spcPct val="90000"/></a:lnSpc></a:pPr>'))
        .join('')
    );
    // 18 pt at 90 %: the line tops step 19.44 pt (they do not round), and each baseline
    // rounds on its own: the rounding does not add up from line to line.
    const tops = t.lines.map((l) => l.top / EMU_PER_PT);
    expect(tops[3] - tops[0]).toBeCloseTo(3 * 19.44, 6);
    const baselines = baselinesPt(t);
    t.lines.forEach((line, i) => {
      const exact = tops[i] + baselineOffsetPt(18, 19.44, line.runs[0].face.metrics);
      expect(baselines[i]).toBe(Math.round(exact));
    });
  });

  test('vertical anchoring moves rounded baselines by an offset that does not round', () => {
    const top = lay(para('a'), '<a:bodyPr tIns="0" bIns="0"/>');
    const centered = lay(para('a'), '<a:bodyPr tIns="0" bIns="0" anchor="ctr"/>');
    expect(top.lines[0].baseline / EMU_PER_PT).toBe(Math.round(top.lines[0].baseline / EMU_PER_PT));
    const offset = (4_000_000 - top.contentHeight) / 2;
    expect(centered.lines[0].baseline - top.lines[0].baseline).toBeCloseTo(offset, 6);
  });

  test('in columns, baselines are placed from the top of each column', () => {
    const t = lay(
      Array.from({ length: 20 }, (_, i) =>
        para(`line ${i}`, '<a:pPr><a:lnSpc><a:spcPct val="90000"/></a:lnSpc></a:pPr>')
      ).join(''),
      '<a:bodyPr tIns="0" bIns="0" numCol="2" spcCol="0"/>'
    );
    const firstOf = (column: number) => t.lines.find((l) => l.column === column)!;
    expect(firstOf(1).baseline - firstOf(1).top).toBeCloseTo(
      firstOf(0).baseline - firstOf(0).top,
      6
    );
    expect(firstOf(1).top).toBe(firstOf(0).top);
  });

  test('a bullet larger than the text does not make its line taller', () => {
    const plain = lay(para('a', '<a:pPr marL="342900" indent="-342900"><a:buNone/></a:pPr>'));
    const big = lay(
      para(
        'a',
        '<a:pPr marL="342900" indent="-342900"><a:buSzPct val="150000"/><a:buChar char="–"/></a:pPr>'
      )
    );
    expect(big.lines[0].height).toBe(plain.lines[0].height);
  });

  test('stored fontScale draws whole-point sizes', () => {
    const t = lay(
      para('a', '', 2800),
      '<a:bodyPr tIns="0"><a:normAutofit fontScale="62500"/></a:bodyPr>'
    );
    const run = t.lines[0].runs.find((x) => x.kind === 'text')!;
    expect(run.sizePt).toBe(18);
  });
});

describe('chart text lines (faceLinePt)', () => {
  // Calibri's win ascent and descent, 1950 and 550 of 2048 units.
  const calibri = face(1950 / 2048, 550 / 2048);
  test('a single line is the ascent plus the descent, with the baseline at the ascent', () => {
    const line = faceLinePt(10, { pct: 1 }, 0, calibri)!;
    expect(line.height).toBeCloseTo(12.207, 3);
    expect(line.baseline).toBeCloseTo(9.521, 3);
    // The whole-point rule puts it at 10 pt below the line top, as `slide-structure.pdf` shows
    // for its 10 pt axis labels; the 1.2 × size line puts it at 9 pt.
    expect(roundToWholePoint(line.baseline * EMU_PER_PT) / EMU_PER_PT).toBe(10);
    expect(
      roundToWholePoint(
        baselineOffsetPt(10, lineHeightPt(10, { pct: 1 }, 0), calibri) * EMU_PER_PT
      ) / EMU_PER_PT
    ).toBe(9);
  });
  test('percent spacing puts 75 % of the extra space above the baseline', () => {
    const line = faceLinePt(10, { pct: 2 }, 0, calibri)!;
    expect(line.height).toBeCloseTo(24.414, 3);
    expect(line.baseline).toBeCloseTo(9.521 + 0.75 * 12.207, 2);
  });
  test('exact spacing keeps the slide line model', () => {
    expect(faceLinePt(10, { pts: 1200 }, 0, calibri)).toBeUndefined();
  });
});
