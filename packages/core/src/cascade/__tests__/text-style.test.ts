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
