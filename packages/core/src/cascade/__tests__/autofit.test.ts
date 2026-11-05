// `a:normAutofit` values clamp at the parse boundary (ECMA-376 Part 1, CT_TextNormalAutofit):
// fontScale to 1 %–100 %, lnSpcReduction to 0 %–100 %.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import { resolveSlide, type Autofit } from '../index.ts';

function autofitOf(attrs: string): Autofit {
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: 0,
          y: 0,
          cx: 914_400,
          cy: 914_400,
          text: para(run('x')),
          bodyPr: `<a:bodyPr><a:normAutofit ${attrs}/></a:bodyPr>`,
        }),
      },
    ],
  });
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('open failed');
  const el = resolveSlide(r.pkg, 0)!.elements[0];
  if (el.kind !== 'shape' || !el.text) throw new Error('no text');
  return el.text.body.autofit;
}

describe('normAutofit values', () => {
  test('keep values in range', () => {
    expect(autofitOf('fontScale="62500" lnSpcReduction="20000"')).toEqual({
      kind: 'normal',
      fontScale: 0.625,
      lineSpaceReduction: 0.2,
    });
  });

  test('clamp values out of range', () => {
    expect(autofitOf('fontScale="0" lnSpcReduction="200000"')).toEqual({
      kind: 'normal',
      fontScale: 0.01,
      lineSpaceReduction: 1,
    });
    expect(autofitOf('fontScale="200000" lnSpcReduction="-5000"')).toEqual({
      kind: 'normal',
      fontScale: 1,
      lineSpaceReduction: 0,
    });
  });
});
