import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../../store/index.ts';
import { resolveSlide } from '../../../cascade/index.ts';
import { createFixedMeasurer, layoutSlide, type Line, type TextLayout } from '../../index.ts';
import { paragraphBidi, visualOrderOfLevels } from '../rtl-paragraph.ts';
import { runGlyphs } from '../bidi-line.ts';
import { bidiAlgorithm } from '../bidi.ts';
import { leadingOffset, paragraphAlignment } from '../paragraph-alignment.ts';

// The shape is 6,000,000 EMU wide with the default insets (91,440 EMU on each side).
const BOX_RIGHT = 6_000_000 - 91_440;
const BOX_LEFT = 91_440;

const shape = (body: string): string =>
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6000000" cy="4000000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
  `<p:txBody><a:bodyPr/><a:lstStyle/>${body}</p:txBody></p:sp>`;

function lay(body: string): TextLayout {
  const r = readPptx(buildPptx({ slides: [{ shapes: shape(body), layout: 'blank' }] }));
  if (!r.ok) throw new Error('open failed');
  const item = layoutSlide(resolveSlide(r.pkg, 0)!, createFixedMeasurer()).items[0];
  if (item.kind !== 'shape' || !item.text) throw new Error('no text');
  return item.text;
}

const r = (t: string, lang = 'ar-SA', size = 1800): string =>
  `<a:r><a:rPr lang="${lang}" sz="${size}"/><a:t xml:space="preserve">${t}</a:t></a:r>`;
const para = (runs: string, pPr = '<a:pPr rtl="1" algn="r"/>'): string =>
  `<a:p>${pPr}${runs}</a:p>`;

/** The characters of a line from left to right, as paint draws them. */
function visual(line: Line): string {
  return [...line.runs]
    .sort((a, b) => a.x[0] - b.x[0])
    .map((run) => (run.rtl ? [...run.text].reverse().join('') : run.text))
    .join('');
}

const reversed = (text: string): string => [...text].reverse().join('');
const left = (line: Line): number => Math.min(...line.runs.map((run) => run.x[0]));
const right = (line: Line): number => Math.max(...line.runs.map((run) => run.x[run.x.length - 1]));

// Arabic and Hebrew words of the tests.
const SALAM = 'سلام';
const DAR = 'دار';
const SHALOM = 'שלום';
const DATE = 'תאריך';

describe('bidi levels (UAX #9)', () => {
  test('a left-to-right paragraph without right-to-left letters needs no levels', () => {
    expect(paragraphBidi([{ text: 'Hello 123' }], false)).toBeUndefined();
  });

  test('digits in a right-to-left run take an even level', () => {
    const bidi = paragraphBidi([{ text: `${SALAM} 12` }], true)!;
    expect([...bidi.levels]).toEqual([1, 1, 1, 1, 1, 2, 2]);
  });

  test('L2 reverses each run of a level and higher', () => {
    expect(visualOrderOfLevels([1, 1, 2, 2, 1])).toEqual([4, 2, 3, 1, 0]);
    expect(visualOrderOfLevels([0, 1, 1, 0])).toEqual([0, 2, 1, 3]);
  });
});

describe('visual order of a line', () => {
  test('Arabic with a Latin word reads from the right, the Latin word left to right', () => {
    const line = lay(para(r(`${SALAM} `) + r('PPTX', 'en-US') + r(` ${DAR}.`))).lines[0];
    expect(visual(line)).toBe(`.${reversed(DAR)} PPTX ${reversed(SALAM)}`);
    // Runs stay in source order for the caret and the selection.
    expect(line.runs.map((run) => run.start)).toEqual(
      [...line.runs.map((run) => run.start)].sort((a, b) => a - b)
    );
  });

  test('Hebrew with a number keeps the number left to right', () => {
    const line = lay(para(r(`${SHALOM} 120.50 ${DATE}`, 'he-IL'))).lines[0];
    expect(visual(line)).toBe(`${reversed(DATE)} 120.50 ${reversed(SHALOM)}`);
  });

  test('Hebrew and Arabic words in a left-to-right paragraph read right to left in place', () => {
    const line = lay(
      para(
        r('The word ', 'en-US') + r(SALAM) + r(' and ', 'en-US') + r(SHALOM, 'he-IL'),
        '<a:pPr/>'
      )
    ).lines[0];
    expect(visual(line)).toBe(`The word ${reversed(SALAM)} and ${reversed(SHALOM)}`);
  });

  test('neutrals at a run edge stay in their run', () => {
    // A left-to-right run keeps its colon and quotation marks inside it.
    const a = lay(para(r('phrase: ', 'en-US') + r(SALAM))).lines[0];
    expect(visual(a)).toBe(`${reversed(SALAM)} phrase:`);
    const b = lay(para(r(`${SHALOM}: `, 'he-IL') + r('"Hi, all!"', 'en-US'))).lines[0];
    expect(visual(b)).toBe(`"Hi, all!" :${reversed(SHALOM)}`);
    // A run with no strong letter reads in the paragraph's direction.
    const c = lay(para(r('ECMA-376', 'en-US') + r('.', 'he-IL'))).lines[0];
    expect(visual(c)).toBe('.ECMA-376');
  });

  test('a slash between digits separates numbers in an Arabic run, not in a Hebrew run', () => {
    const arabic = lay(para(r(`${SALAM} 17/08/96`))).lines[0];
    expect(visual(arabic)).toStartWith('96/08/17');
    const split = lay(para(r(`${SALAM} 17/`) + r('08/96'))).lines[0];
    expect(visual(split)).toContain('96/08/17');
    // The language of the run that holds the slash decides.
    const mixed = lay(para(r(`${SHALOM} 17/08`, 'he-IL') + r('/96'))).lines[0];
    expect(visual(mixed)).toContain('96/17/08');
    const hebrew = lay(para(r(`${SHALOM} 852/20`, 'he-IL'))).lines[0];
    expect(visual(hebrew)).toContain('852/20');
  });
});

describe('runs with more than one script', () => {
  test('a run that mixes scripts reads in the paragraph direction (UAX #9)', () => {
    const words = '\u05d4\u05d5\u05d0 \u05db\u05dc\u05d9';
    const line = lay(para(r(`Excel ${words}`, 'he-IL'))).lines[0];
    expect(visual(line)).toBe(`${reversed(words)} Excel`);
  });

  test('right-to-left words in two runs of a left-to-right paragraph keep their order', () => {
    const line = lay(para(r(`Hello ${SHALOM} `, 'he-IL') + r(DATE, 'he-IL'), '<a:pPr/>')).lines[0];
    expect(visual(line)).toBe(`Hello ${reversed(`${SHALOM} ${DATE}`)}`);
  });

  test('bidi controls reorder a left-to-right paragraph without right-to-left letters', () => {
    const line = lay(para(r('abc\u202edef\u202c', 'en-US'), '<a:pPr/>')).lines[0];
    expect(visual(line).replace(/[\u202c\u202e]/gu, '')).toBe('abcfed');
  });

  test('a right-to-left script outside the Basic Multilingual Plane reads right to left', () => {
    const adlam = '\u{1e900}\u{1e901}';
    expect(paragraphBidi([{ text: `a ${adlam}` }], false)).toBeDefined();
  });
});

describe('a run that reads against the paragraph resolves like an isolate', () => {
  const ltr = (runs: string): string => visual(lay(para(runs, '<a:pPr/>')).lines[0]);
  const rtl = (runs: string): string => visual(lay(para(runs)).lines[0]);
  const he = (t: string): string => r(t, 'he-IL');

  test('numbers at the edges of a right-to-left run keep their order', () => {
    expect(ltr(he(`${SHALOM} 123`))).toBe(`123 ${reversed(SHALOM)}`);
    expect(ltr(he(`2024 ${DATE}`))).toBe(`${reversed(DATE)} 2024`);
    expect(ltr(he(`${DATE} 12-14`))).toBe(`12-14 ${reversed(DATE)}`);
    expect(ltr(he('\u05d1-50%'))).toBe('50%-\u05d1');
    // An Arabic run separates numbers at a slash.
    expect(ltr(r(`${SALAM} 12/34`))).toBe(`34/12 ${reversed(SALAM)}`);
  });

  test('spaces and digits inside a left-to-right run keep it whole', () => {
    expect(rtl(he(`${SHALOM} `) + r('"5 apples"', 'en-US'))).toBe(`"5 apples" ${reversed(SHALOM)}`);
    expect(rtl(he(`${DATE} `) + r('$5 USD', 'en-US'))).toBe(`$5 USD ${reversed(DATE)}`);
  });

  test('spaces beside a bracket paired outside the run stay with the bracket', () => {
    // Brackets at level 1 draw mirrored (L4), so `) abc (` shows as ( abc ) from the left.
    expect(rtl(he(`${SHALOM} `) + r('( abc ', 'en-US') + r(')', 'en-US'))).toBe(
      `) abc ( ${reversed(SHALOM)}`
    );
    // The space after the opening bracket stands beside it: `(abc )` from the left, as UAX #9
    // orders the paragraph's text when it resolves it as one.
    const runs = [{ text: `${SHALOM} ` }, { text: '(' }, { text: ' abc)' }];
    expect([...paragraphBidi(runs, true)!.levels]).toEqual([
      ...bidiAlgorithm.getEmbeddingLevels(runs.map((run) => run.text).join(''), 'rtl').levels,
    ]);
    expect(rtl(he(`${SHALOM} `) + r('(', 'en-US') + r(' abc)', 'en-US'))).toBe(
      `)abc ( ${reversed(SHALOM)}`
    );
    expect(ltr(r('abc (', 'en-US') + he(`${SHALOM} )`) + r(' def', 'en-US'))).toBe(
      `abc (${reversed(SHALOM)} ) def`
    );
  });

  test('a directional mark does not turn off the run isolates', () => {
    const line = rtl(he(`${SHALOM}\u200f `) + r('"5 apples"', 'en-US'));
    expect(line.replace('\u200f', '')).toBe(`"5 apples" ${reversed(SHALOM)}`);
  });

  test('a bracket pair across two runs resolves as one pair (UAX #9 N0)', () => {
    // Both brackets take level 1, so paint mirrors them (L4) and the line reads (abc).
    expect(rtl(he(`${SHALOM} (`) + r('abc)', 'en-US'))).toBe(`)abc( ${reversed(SHALOM)}`);
  });
});

describe('physical alignment', () => {
  for (const [algn, side] of [
    ['ctr', 'center'],
    ['dist', 'fill'],
    ['justLow', 'fill'],
    ['thaiDist', 'fill'],
  ] as const)
    test(`algn="${algn}" in a right-to-left paragraph`, () => {
      const words = Array.from({ length: 12 }, () => SALAM.repeat(4)).join(' ');
      const t = lay(para(r(words), `<a:pPr rtl="1" algn="${algn}"/>`));
      const first = t.lines[0];
      const SPACE = 0.25 * 18 * 12_700;
      if (side === 'center') {
        const last = t.lines[t.lines.length - 1];
        expect((left(last) + right(last)) / 2).toBeCloseTo((BOX_LEFT + BOX_RIGHT) / 2, 3);
      } else {
        expect(right(first)).toBeCloseTo(BOX_RIGHT, 3);
        expect(left(first) + SPACE).toBeCloseTo(BOX_LEFT, 3);
      }
      const last = t.lines[t.lines.length - 1];
      // A justified last line starts at the leading (right) edge; a distributed one fills.
      if (algn === 'justLow') expect(right(last)).toBeCloseTo(BOX_RIGHT, 3);
    });

  test('DrawingML algn names physical edges, also in a right-to-left paragraph', () => {
    expect(paragraphAlignment('l')).toBe('left');
    expect(paragraphAlignment('r')).toBe('right');
    expect(leadingOffset('right', true, 10)).toBe(0);
    expect(leadingOffset('left', true, 10)).toBe(10);
    expect(leadingOffset('right', false, 10)).toBe(10);
    expect(leadingOffset('both', true, 10)).toBe(0);
  });

  test('algn="r" ends a right-to-left line at the right inset', () => {
    const line = lay(para(r(SALAM), '<a:pPr rtl="1" algn="r"/>')).lines[0];
    expect(right(line)).toBeCloseTo(BOX_RIGHT, 3);
  });

  test('algn="l" starts a right-to-left line at the left inset', () => {
    const line = lay(para(r(SALAM), '<a:pPr rtl="1" algn="l"/>')).lines[0];
    expect(left(line)).toBeCloseTo(BOX_LEFT, 3);
  });

  test('marL is measured from the right in a right-to-left paragraph', () => {
    const line = lay(para(r(SALAM), '<a:pPr rtl="1" algn="r" marL="457200"/>')).lines[0];
    expect(right(line)).toBeCloseTo(BOX_RIGHT - 457_200, 3);
  });

  test('a number bullet stands right of the text and reads in the paragraph direction', () => {
    const t = lay(
      para(
        r(SALAM),
        '<a:pPr rtl="1" algn="r" marL="457200" indent="-457200"><a:buAutoNum type="arabicPeriod"/></a:pPr>'
      )
    );
    const line = t.lines[0];
    const bullets = line.runs.filter((run) => run.kind === 'bullet');
    expect(visual({ ...line, runs: bullets })).toBe('.1');
    expect(Math.max(...bullets.map((b) => b.x[b.x.length - 1]))).toBeCloseTo(BOX_RIGHT, 3);
    const text = line.runs.filter((run) => run.kind === 'text');
    expect(Math.max(...text.map((b) => b.x[b.x.length - 1]))).toBeCloseTo(BOX_RIGHT - 457_200, 3);
  });
});

describe('trailing spaces', () => {
  // 18 pt text in the fixed measurer: letters 9 pt, spaces 4.5 pt wide.
  const word = SALAM.repeat(4);
  const words = Array.from({ length: 12 }, () => word).join(' ');

  test('a wrapped right-to-left line ends at the right inset, its space hanging on the left', () => {
    const t = lay(para(r(words)));
    expect(t.lines.length).toBeGreaterThan(1);
    const first = t.lines[0];
    expect(right(first)).toBeCloseTo(BOX_RIGHT, 3);
    // The space that ends the line in source order stands at its left end.
    const run = first.runs[0];
    expect(run.rtl).toBe(true);
    expect(run.text.endsWith(' ')).toBe(true);
  });

  test('a left-to-right run that ends a wrapped line ends at the right inset', () => {
    const t = lay(
      para(r('phrase phrase phrase phrase phrase phrase phrase phrase ', 'en-US') + r(words))
    );
    const first = t.lines[0];
    const latin = first.runs.find((run) => !run.rtl)!;
    const lastLetter = latin.x[latin.text.trimEnd().length];
    expect(lastLetter).toBeCloseTo(BOX_RIGHT, 3);
  });

  test('a justified right-to-left line fills the line, and its last line starts at the right', () => {
    const t = lay(para(r(words), '<a:pPr rtl="1" algn="just"/>'));
    // The space that ends the line in source order hangs past the left inset.
    const SPACE = 0.25 * 18 * 12_700;
    expect(left(t.lines[0]) + SPACE).toBeCloseTo(BOX_LEFT, 3);
    expect(right(t.lines[0])).toBeCloseTo(BOX_RIGHT, 3);
    expect(right(t.lines[t.lines.length - 1])).toBeCloseTo(BOX_RIGHT, 3);
  });
});

describe('caret positions', () => {
  test('offsets of a right-to-left line run from the right edge to the left', () => {
    const line = lay(para(r(SALAM))).lines[0];
    expect(line.caretX[0]).toBeCloseTo(BOX_RIGHT, 3);
    for (let i = 1; i < line.caretX.length; i += 1)
      expect(line.caretX[i]).toBeLessThan(line.caretX[i - 1]);
  });

  test('an empty right-to-left paragraph puts the caret at the right', () => {
    const line = lay(para('<a:endParaRPr lang="ar-SA" sz="1800"/>')).lines[0];
    expect(line.caretX[0]).toBeCloseTo(BOX_RIGHT, 3);
  });
});

describe('glyph positions for readers that place each glyph', () => {
  test('a right-to-left run puts unit i in cell n - 1 - i and mirrors brackets', () => {
    const glyphs = runGlyphs({ text: 'a(b', x: [0, 10, 20, 30], rtl: true });
    expect(glyphs.map((g) => [g.char, g.x])).toEqual([
      ['a', 20],
      [')', 10],
      ['b', 0],
    ]);
  });

  test('a surrogate pair is one character in its two cells', () => {
    const adlam = '\u{1e900}';
    const glyphs = runGlyphs({ text: `${adlam}a`, x: [0, 10, 10, 20], rtl: true });
    expect(glyphs.map((g) => [g.char, g.x])).toEqual([
      [adlam, 10],
      ['a', 0],
    ]);
  });

  test('a left-to-right run keeps its characters and positions', () => {
    const glyphs = runGlyphs({ text: 'a(b', x: [0, 10, 20, 30] });
    expect(glyphs.map((g) => [g.char, g.x])).toEqual([
      ['a', 0],
      ['(', 10],
      ['b', 20],
    ]);
  });
});

describe('bracket pairs across runs (UAX #9 BD16)', () => {
  // `(` ends the first run and `)` ends the second, a left-to-right run of a right-to-left
  // paragraph. A paired `)` keeps the paragraph level with its `(`; an unpaired one reads
  // inside its run, one level above the paragraph.
  const levelsOf = (between: string, after = ''): number[] => {
    const runs = [{ text: `${SHALOM} (` }, { text: `${between}abc)` }, { text: after }];
    return [...paragraphBidi(runs, true)!.levels];
  };
  const open = SHALOM.length + 1;
  const close = (between: string): number => open + 1 + between.length + 3;
  const LINE_FEED = String.fromCharCode(0x0a);
  const PARAGRAPH_SEPARATOR = String.fromCharCode(0x2029);

  test('a pair found before 63 more openings stays a pair', () => {
    const levels = levelsOf('', ` ${'('.repeat(63)}`);
    expect(levels[close('')]).toBe(levels[open]);
    expect(levels[close('')]).toBe(1);
  });

  test('a pair found before the 64th opening stays a pair, as bidi-js keeps it', () => {
    const levels = levelsOf('', ` ${'('.repeat(64)}`);
    expect(levels[close('')]).toBe(levels[open]);
    expect(levels[close('')]).toBe(1);
  });

  test('a line feed between the brackets ends the paragraph, so they do not pair', () => {
    const levels = levelsOf(LINE_FEED);
    expect(levels[open]).toBe(1);
    expect(levels[close(LINE_FEED)]).toBe(2);
  });

  test('a paragraph separator between the brackets ends the paragraph too', () => {
    const levels = levelsOf(PARAGRAPH_SEPARATOR);
    expect(levels[open]).toBe(1);
    expect(levels[close(PARAGRAPH_SEPARATOR)]).toBe(2);
  });
});
