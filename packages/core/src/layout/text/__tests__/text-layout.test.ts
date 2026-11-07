import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../../store/index.ts';
import { resolveSlide } from '../../../cascade/index.ts';
import {
  createFixedMeasurer,
  layoutSlide,
  type TextLayout,
  type TextMeasurer,
} from '../../index.ts';

const shape = (body: string, bodyPr = '<a:bodyPr/>'): string =>
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="4000000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
  `<p:txBody>${bodyPr}<a:lstStyle/>${body}</p:txBody></p:sp>`;

function lay(body: string, bodyPr?: string): TextLayout {
  const r = readPptx(buildPptx({ slides: [{ shapes: shape(body, bodyPr), layout: 'blank' }] }));
  if (!r.ok) throw new Error('open failed');
  const item = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer()).items[0];
  if (item.kind !== 'shape' || !item.text) throw new Error('no text');
  return item.text;
}

const r = (t: string, size = 1800): string =>
  `<a:r><a:rPr lang="en-US" sz="${size}"/><a:t>${t}</a:t></a:r>`;
const br = (size = 1800): string => `<a:br><a:rPr lang="en-US" sz="${size}"/></a:br>`;
const end = (size = 1800): string => `<a:endParaRPr lang="en-US" sz="${size}"/>`;

describe('a trailing line break (R4)', () => {
  test('starts a new, empty line', () => {
    const t = lay(`<a:p>${r('abc')}${br()}${end()}</a:p>`);
    expect(t.lines.length).toBe(2);
    const last = t.lines[1];
    expect(last.start).toBe(4);
    expect(last.end).toBe(4);
    expect(last.runs).toEqual([]);
    expect(last.caretX.length).toBe(1);
  });

  test('is as tall as text followed by an empty paragraph', () => {
    const withBreak = lay(`<a:p>${r('abc')}${br()}${end()}</a:p>`);
    const twoParas = lay(`<a:p>${r('abc')}${end()}</a:p><a:p>${end()}</a:p>`);
    expect(withBreak.contentHeight).toBeCloseTo(twoParas.contentHeight, 6);
  });

  test('the paragraph mark sizes the empty line', () => {
    const t = lay(`<a:p>${r('abc', 1800)}${br(1800)}${end(3600)}</a:p>`);
    expect(t.lines[1].height).toBeGreaterThan(t.lines[0].height * 1.9);
  });

  test('moves bottom-anchored text up by one line', () => {
    const plain = lay(`<a:p>${r('abc')}${end()}</a:p>`, '<a:bodyPr anchor="b"/>');
    const broken = lay(`<a:p>${r('abc')}${br()}${end()}</a:p>`, '<a:bodyPr anchor="b"/>');
    expect(plain.lines[0].top - broken.lines[0].top).toBeCloseTo(broken.lines[1].height, 3);
  });

  test('a break inside a paragraph adds no extra line', () => {
    expect(lay(`<a:p>${r('abc')}${br()}${r('def')}${end()}</a:p>`).lines.length).toBe(2);
  });
});

describe('bullets on aligned lines (R5)', () => {
  const bulleted = (algn: string): TextLayout =>
    lay(
      `<a:p><a:pPr algn="${algn}" marL="342900" indent="-342900"><a:buChar char="•"/></a:pPr>${r('Item')}</a:p>`
    );
  const gap = (t: TextLayout): { bullet: number; gap: number } => {
    const runs = t.lines[0].runs;
    const b = runs.find((x) => x.kind === 'bullet')!;
    const text = runs.find((x) => x.kind === 'text')!;
    return { bullet: b.x[0], gap: text.x[0] - b.x[0] };
  };

  test('centered and right-aligned bullets keep their distance to the text', () => {
    const left = gap(bulleted('l'));
    for (const algn of ['ctr', 'r']) {
      const g = gap(bulleted(algn));
      expect(g.gap).toBeCloseTo(left.gap, 6);
      expect(g.bullet).toBeGreaterThan(left.bullet + 1_000_000);
    }
  });

  test('a right-aligned line ends at the right inset', () => {
    const t = bulleted('r');
    const text = t.lines[0].runs.find((x) => x.kind === 'text')!;
    // Box 6 000 000 wide, default insets 91 440 on each side.
    expect(text.x[text.x.length - 1]).toBeCloseTo(6_000_000 - 91_440, 0);
  });
});

describe('autonumber sequences (R6)', () => {
  const numbered = (lvl: number, t: string, scheme: string, startAt?: number): string =>
    `<a:p><a:pPr lvl="${lvl}" marL="${(lvl + 1) * 342900}" indent="-342900">` +
    `<a:buAutoNum type="${scheme}"${startAt ? ` startAt="${startAt}"` : ''}/></a:pPr>${r(t)}</a:p>`;
  const markers = (t: TextLayout): string[] =>
    t.lines.map((l) => l.runs.find((x) => x.kind === 'bullet')?.text ?? '');

  test('deeper levels restart under each parent', () => {
    const t = lay(
      numbered(0, 'A', 'arabicPeriod') +
        numbered(1, 'x', 'alphaLcPeriod') +
        numbered(1, 'y', 'alphaLcPeriod') +
        numbered(0, 'B', 'arabicPeriod') +
        numbered(1, 'z', 'alphaLcPeriod')
    );
    expect(markers(t)).toEqual(['1.', 'a.', 'b.', '2.', 'a.']);
  });

  test('a scheme change restarts the sequence', () => {
    const t = lay(
      numbered(0, 'A', 'arabicPeriod') +
        numbered(0, 'B', 'arabicPeriod') +
        numbered(0, 'C', 'romanUcPeriod') +
        numbered(0, 'D', 'romanUcPeriod')
    );
    expect(markers(t)).toEqual(['1.', '2.', 'I.', 'II.']);
  });

  test('a startAt change restarts the sequence', () => {
    const t = lay(
      numbered(0, 'A', 'arabicPeriod') +
        numbered(0, 'B', 'arabicPeriod', 5) +
        numbered(0, 'C', 'arabicPeriod', 5)
    );
    expect(markers(t)).toEqual(['1.', '5.', '6.']);
  });

  test('empty paragraphs keep the sequence; unnumbered text ends it', () => {
    const t = lay(
      numbered(0, 'A', 'arabicParenR') +
        `<a:p><a:pPr lvl="0"><a:buAutoNum type="arabicParenR"/></a:pPr>${end()}</a:p>` +
        numbered(0, 'B', 'arabicParenR') +
        `<a:p><a:pPr lvl="0"><a:buNone/></a:pPr>${r('plain')}</a:p>` +
        numbered(0, 'C', 'arabicParenR')
    );
    expect(markers(t)).toEqual(['1)', '', '2)', '', '1)']);
  });

  test('an Arabic or East Asian number draws in the first run font for its script', () => {
    const run = (t: string, fonts: string): string =>
      `<a:r><a:rPr lang="ar-SA" sz="1800">${fonts}</a:rPr><a:t>${t}</a:t></a:r>`;
    const fonts =
      '<a:latin typeface="Liberation Sans"/><a:ea typeface="Noto Sans CJK JP"/>' +
      '<a:cs typeface="Noto Sans Arabic"/>';
    const para = (scheme: string): string =>
      `<a:p><a:pPr marL="457200" indent="-457200"><a:buFont typeface="Carlito"/>` +
      `<a:buAutoNum type="${scheme}"/></a:pPr>${run('نص', fonts)}</a:p>`;
    const t = lay(para('arabic1Minus') + para('ea1ChsPeriod') + para('arabicPeriod'));
    const families = t.lines.map((l) => l.runs.find((x) => x.kind === 'bullet')?.face.family);
    expect(markers(t)).toEqual(['أ‌-', '一.', '1.']);
    expect(families).toEqual(['Noto Sans Arabic', 'Noto Sans CJK JP', 'Carlito']);
  });
});

describe('symbol font bullets', () => {
  /** A measurer that has no Wingdings: the face for it is a substitute. */
  const noWingdings = (): TextMeasurer => {
    const base = createFixedMeasurer();
    return {
      ...base,
      resolve: (req) => ({ ...base.resolve(req), substituted: req.family === 'Wingdings' }),
    };
  };
  const marker = (char: string, measurer: TextMeasurer): string | undefined => {
    const body = `<a:p><a:pPr marL="342900" indent="-342900"><a:buFont typeface="Wingdings"/><a:buChar char="${char}"/></a:pPr>${r('x')}</a:p>`;
    const p = readPptx(buildPptx({ slides: [{ shapes: shape(body), layout: 'blank' }] }));
    if (!p.ok) throw new Error('open failed');
    const item = layoutSlide(resolveSlide(p.pkg, 0)!, measurer).items[0];
    if (item.kind !== 'shape' || !item.text) throw new Error('no text');
    return item.text.lines[0].runs.find((x) => x.kind === 'bullet')?.text;
  };

  test("a substituted Wingdings draws the gallery bullets' Unicode twins", () => {
    expect(['§', 'q', 'v', 'Ø', 'ü'].map((c) => marker(c, noWingdings()))).toEqual([
      '▪',
      '❑',
      '❖',
      '➢',
      '✔',
    ]);
  });

  test('Wingdings itself draws its own character', () => {
    expect(marker('v', createFixedMeasurer())).toBe('v');
  });
});

describe('tab stops (ECMA-376 Part 1, 21.1.2.1.36 tab)', () => {
  const LEFT_INSET = 91_440;
  const tabbed = (tabs: string, text: string): readonly number[] => {
    const t = lay(
      `<a:p><a:pPr><a:tabLst>${tabs}</a:tabLst></a:pPr>${r(text)}</a:p>`,
      '<a:bodyPr wrap="square"/>'
    );
    expect(t.lines.length).toBe(1);
    return t.lines[0].runs.find((x) => x.kind === 'text')!.x;
  };

  test('right, centered, and decimal stops place the text after the tab', () => {
    const x = tabbed(
      '<a:tab pos="2000000" algn="r"/><a:tab pos="3000000" algn="ctr"/><a:tab pos="4000000" algn="dec"/>',
      'a\tbb\tcc\t1.5'
    );
    // "bb" ends at the right stop, "cc" centers on its stop, and the full stop of "1.5" starts
    // at the decimal stop.
    expect(x[4]).toBeCloseTo(LEFT_INSET + 2_000_000, 0);
    expect((x[5] + x[7]) / 2).toBeCloseTo(LEFT_INSET + 3_000_000, 0);
    expect(x[9]).toBeCloseTo(LEFT_INSET + 4_000_000, 0);
  });

  test('a decimal stop with no full stop ends the text there', () => {
    const x = tabbed('<a:tab pos="2000000" algn="dec"/>', 'a\t15');
    expect(x[4]).toBeCloseTo(LEFT_INSET + 2_000_000, 0);
  });

  test('text at a stop near the right edge moves left to end at the edge', () => {
    const x = tabbed('<a:tab pos="5900000" algn="dec"/>', 'a\t12.5');
    expect(x[x.length - 1]).toBeCloseTo(6_000_000 - LEFT_INSET, 0);
  });

  test('text wider than the room before a right stop starts at the tab', () => {
    const x = tabbed('<a:tab pos="200000" algn="r"/>', 'a\tlonger text');
    expect(x[1]).toBeLessThan(LEFT_INSET + 200_000);
    expect(x[2]).toBeCloseTo(x[1], 0);
  });
});

describe('justified lines (ECMA-376 Part 1, 21.1.2.2.7 pPr algn)', () => {
  const words = 'justified text that wraps onto more than one line in the box '.repeat(4);

  test('a wrapped line keeps finite caret positions after its trailing space', () => {
    const t = lay(`<a:p><a:pPr algn="just"/>${r(words)}${end()}</a:p>`);
    expect(t.lines.length).toBeGreaterThan(1);
    for (const line of t.lines) {
      for (const x of line.caretX) expect(Number.isFinite(x)).toBe(true);
      for (const run of line.runs) for (const x of run.x) expect(Number.isFinite(x)).toBe(true);
    }
    // The caret after the trailing space sits at or after the caret before it.
    const first = t.lines[0].caretX;
    expect(first[first.length - 1]).toBeGreaterThanOrEqual(first[first.length - 2]);
  });
});
