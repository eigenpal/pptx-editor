// Bounds on file numbers in the text cascade: a crafted value reads within its schema type, so
// layout and autofit never loop on it.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import { resolveSlide, type ResolvedShape } from '../index.ts';

const IN = 914_400;

function body(pPr: string, bodyPr = '<a:bodyPr/>') {
  const r = readPptx(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, bodyPr, text: para(run('x'), pPr) }),
        },
      ],
    })
  );
  if (!r.ok) throw new Error('unreadable');
  return (resolveSlide(r.pkg, 0)!.elements[0] as ResolvedShape).text!;
}

describe('text cascade bounds', () => {
  test('a percent string in a:spcPct reads as its percent (ECMA-376 Part 1, 20.1.10.77)', () => {
    const at = (val: string) =>
      body(`<a:pPr><a:lnSpc><a:spcPct val="${val}"/></a:lnSpc></a:pPr>`).paragraphs[0]!.lineSpacing;
    expect(at('150%')).toEqual({ pct: 1.5 });
    expect(at('150000')).toEqual({ pct: 1.5 });
    expect(at('150.5%')).toEqual({ pct: 1.505 });
    expect(at('-50%')).toEqual({ pct: 0 });
    // A value that is not a number keeps single spacing.
    expect(at('abc')).toEqual({ pct: 1 });
    const before = body('<a:pPr><a:spcBef><a:spcPct val="50%"/></a:spcBef></a:pPr>');
    expect(before.paragraphs[0]!.spaceBefore).toEqual({ pct: 0.5 });
  });

  test('line spacing in percent reads within ST_TextSpacingPercentOrPercentString', () => {
    const huge = body('<a:pPr><a:lnSpc><a:spcPct val="999999999999999"/></a:lnSpc></a:pPr>');
    expect(huge.paragraphs[0]!.lineSpacing).toEqual({ pct: 132 });
    const negative = body('<a:pPr><a:lnSpc><a:spcPct val="-50000"/></a:lnSpc></a:pPr>');
    expect(negative.paragraphs[0]!.lineSpacing).toEqual({ pct: 0 });
  });

  test('a column count reads as a whole number from 1 to 16', () => {
    expect(body('', '<a:bodyPr numCol="99"/>').body.numCol).toBe(16);
    expect(body('', '<a:bodyPr numCol="2.7"/>').body.numCol).toBe(2);
    expect(body('', '<a:bodyPr numCol="0"/>').body.numCol).toBe(1);
  });
});
