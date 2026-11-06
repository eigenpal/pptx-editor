// East Asian text: the font slot of each character (`a:latin`, `a:ea`, `a:cs`, `a:sym`), the
// theme script fonts that a run's language picks, the East Asian fallback faces, the line
// breaking rules of `eaLnBrk`, `latinLnBrk`, and `p:kinsoku` (measured on the reference PDFs of
// the `east-asian-line-breaking` and `east-asian-typography` fixtures), and East Asian
// justification.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, shape } from '../../../../../scripts/fixtures/builder.ts';
import { escapeText } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import { resolveSlide, type ResolvedShape, type RunProps } from '../../cascade/index.ts';
import {
  breakOpportunities,
  createFixedMeasurer,
  layoutSlide,
  type FontRequest,
  type LayoutItem,
  type TextLayout,
  type TextMeasurer,
} from '../index.ts';
import { fontSlotPieces } from '../text/font-slots.ts';

const IN = 914_400;

const run = (text: string, rPr: string, fonts = ''): string =>
  `<a:r><a:rPr ${rPr} dirty="0">${fonts}</a:rPr><a:t>${escapeText(text)}</a:t></a:r>`;

const SCRIPT_FONTS = [
  ['Jpan', 'Yu Gothic'],
  ['Hans', 'DengXian'],
  ['Hant', 'Microsoft JhengHei'],
  ['Hang', 'Malgun Gothic'],
  ['Arab', 'Tahoma'],
] as const;

function deck(paragraphs: string, cx = 6 * IN, presentationLists = ''): Uint8Array {
  return buildPptx({
    scriptFonts: SCRIPT_FONTS,
    presentationLists,
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx,
          cy: 3 * IN,
          bodyPr: '<a:bodyPr wrap="square" lIns="0" rIns="0"><a:noAutofit/></a:bodyPr>',
          text: paragraphs,
        }),
      },
    ],
  });
}

function runs(bytes: Uint8Array): RunProps[] {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('unreadable');
  const sp = resolveSlide(r.pkg, 0)!.elements[0] as ResolvedShape;
  return sp.text!.paragraphs.flatMap((p) => p.runs.map((x) => x.props));
}

/** A measurer that records each request: East Asian characters 1 em, others 0.5 em. */
function recordingMeasurer(): TextMeasurer & { readonly requests: FontRequest[] } {
  const inner = createFixedMeasurer();
  const requests: FontRequest[] = [];
  return {
    requests,
    resolve: (r) => {
      requests.push(r);
      return inner.resolve(r);
    },
    advances: (face, text) => {
      const out = inner.advances(face, text);
      for (let i = 0; i < text.length; i += 1)
        if ((text.codePointAt(i) ?? 0) >= 0x2e80 && out[i] > 0) out[i] = 1;
      return out;
    },
  };
}

function textOf(bytes: Uint8Array, measurer: TextMeasurer = createFixedMeasurer()): TextLayout {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('unreadable');
  const items = layoutSlide(resolveSlide(r.pkg, 0)!, measurer).items;
  const first = items[0] as Extract<LayoutItem, { kind: 'shape' }>;
  return first.text!;
}

const lineTexts = (layout: TextLayout): string[] =>
  layout.lines.map((l) => l.runs.map((g) => g.text).join(''));

describe('theme script fonts (ECMA-376 Part 1, 20.1.4.1.16 font)', () => {
  test("a run's language picks the theme's East Asian face when a:ea is empty", () => {
    const props = runs(
      deck(
        para(run('日本語', 'lang="ja-JP"')) +
          para(run('简体中文', 'lang="zh-CN"')) +
          para(run('繁體中文', 'lang="zh-TW"')) +
          para(run('한국어', 'lang="ko-KR"')) +
          para(run('English', 'lang="en-US"'))
      )
    );
    expect(props.map((p) => p.ea)).toEqual([
      'Yu Gothic',
      'DengXian',
      'Microsoft JhengHei',
      'Malgun Gothic',
      'Carlito',
    ]);
  });

  test('the alternate language picks the face when the language is not East Asian', () => {
    const [p] = runs(deck(para(run('日本語', 'lang="en-US" altLang="ja-JP"'))));
    expect(p.ea).toBe('Yu Gothic');
    expect(p.altLang).toBe('ja-JP');
  });

  test("a run's language picks the theme's complex-script face", () => {
    const [p] = runs(deck(para(run('عربي', 'lang="ar-SA"'))));
    expect(p.cs).toBe('Tahoma');
  });

  test('a face the run names wins over the script font', () => {
    const [p] = runs(deck(para(run('日本語', 'lang="ja-JP"', '<a:ea typeface="Meiryo"/>'))));
    expect(p.ea).toBe('Meiryo');
    const [ref] = runs(deck(para(run('日本語', 'lang="ja-JP"', '<a:ea typeface="+mj-ea"/>'))));
    expect(ref.ea).toBe('Yu Gothic');
  });
});

describe('font slots', () => {
  const props = (o: Partial<RunProps> = {}) => ({ lang: 'en-US', ...o });

  test('East Asian, complex-script, symbol, and Latin characters take their own slots', () => {
    const pieces = fontSlotPieces('Arial 日本語 عربي ', props({ sym: 'Wingdings' }));
    expect(pieces.map((p) => p.slot)).toEqual(['latin', 'ea', 'latin', 'cs', 'latin', 'sym']);
  });

  test('a space is Latin between East Asian words, complex-script between such words', () => {
    // Reference PDF of `east-asian-theme-fonts`: Korean words are an Arial space apart.
    expect(fontSlotPieces('맑은 고딕', props()).map((p) => p.slot)).toEqual(['ea', 'latin', 'ea']);
    expect(fontSlotPieces('نص عربي', props()).map((p) => p.slot)).toEqual(['cs']);
  });

  test('shared symbols take the East Asian slot only in East Asian text', () => {
    expect(fontSlotPieces('“引用”', props()).map((p) => p.slot)).toEqual(['latin', 'ea', 'latin']);
    expect(fontSlotPieces('“引用”', props({ lang: 'ja-JP' })).map((p) => p.slot)).toEqual(['ea']);
  });

  test('a surrogate pair stays whole and a combining mark keeps its base slot', () => {
    const pieces = fontSlotPieces('a𠀋b', props());
    expect(pieces).toEqual([
      { start: 0, end: 1, slot: 'latin' },
      { start: 1, end: 3, slot: 'ea' },
      { start: 3, end: 4, slot: 'latin' },
    ]);
    expect(fontSlotPieces('が', props()).map((p) => p.slot)).toEqual(['ea']);
  });

  test('symbol characters stay Latin when the run names no a:sym face', () => {
    expect(fontSlotPieces('', props()).map((p) => p.slot)).toEqual(['latin']);
  });

  test('layout measures each stretch with the face of its slot', () => {
    const measurer = recordingMeasurer();
    const layout = textOf(
      deck(
        para(
          run(
            '日本語 Arial',
            'lang="ja-JP"',
            '<a:latin typeface="Arial"/><a:ea typeface="Yu Mincho"/>'
          )
        )
      ),
      measurer
    );
    const runs = layout.lines[0].runs;
    expect(runs.map((g) => [g.text, g.face.family])).toEqual([
      ['日本語', 'Yu Mincho'],
      [' Arial', 'Arial'],
    ]);
    // The East Asian face is asked for as East Asian text in its language, so a missing face
    // falls back to an East Asian face for that language.
    expect(measurer.requests).toContainEqual({
      family: 'Yu Mincho',
      bold: false,
      italic: false,
      script: 'eastAsian',
      lang: 'ja-JP',
    });
    // Source offsets stay continuous across the stretches.
    expect(runs[1].start).toBe(3);
  });
});

describe('East Asian line breaking (eaLnBrk, latinLnBrk)', () => {
  const at = (text: string, ea = true, latin = false): number[] => {
    const b = breakOpportunities(text, ea, latin);
    const out: number[] = [];
    for (let i = 1; i < text.length; i += 1) if (b[i]) out.push(i);
    return out;
  };

  test('eaLnBrk on: no break before closing or after opening punctuation', () => {
    // 「括弧」や: no break after 「, none before 」.
    expect(at('や「括弧」や')).toEqual([1, 3, 5]);
  });

  test('eaLnBrk off: East Asian text still breaks per character, with no kinsoku', () => {
    expect(at('や「括弧」や', false)).toEqual([1, 2, 3, 4, 5]);
  });

  test('Latin words between East Asian text stay whole unless latinLnBrk is on', () => {
    expect(at('のWord', true, false)).toEqual([1]);
    expect(at('のWord', true, true)).toEqual([1, 2, 3, 4]);
    // A space still ends the line before the next word; it never starts one.
    expect(at('ab cd', true, true)).toEqual([1, 3, 4]);
  });

  test('no break inside a surrogate pair or before a combining mark', () => {
    expect(at('𠀋𠀋')).toEqual([2]);
    expect(at('ab́c', true, true)).toEqual([1, 3]);
  });

  test('a narrow box: eaLnBrk off lets opening punctuation end a line (reference PDF)', () => {
    // Four characters fit a line: 4 em at 20 pt in a box 80 pt wide.
    const text = (ea: 0 | 1) =>
      deck(
        para(run('ですや「括弧」', 'lang="ja-JP" sz="2000"'), `<a:pPr eaLnBrk="${ea}"/>`),
        Math.round(80.5 * 12_700)
      );
    const measurer = () => recordingMeasurer();
    expect(lineTexts(textOf(text(1), measurer()))).toEqual(['ですや', '「括弧」']);
    expect(lineTexts(textOf(text(0), measurer()))).toEqual(['ですや「', '括弧」']);
  });
});

describe('custom kinsoku lists (p:kinsoku, ECMA-376 Part 1, 19.2.1.17)', () => {
  const KINSOKU = '<p:kinsoku lang="ja-JP" invalStChars="か" invalEndChars="お"/>';
  const at = (text: string, lists?: { notStart: string; notEnd: string }): number[] => {
    const b = breakOpportunities(text, true, false, lists);
    const out: number[] = [];
    for (let i = 1; i < text.length; i += 1) if (b[i]) out.push(i);
    return out;
  };

  test('the lists replace the built-in lists', () => {
    const lists = { notStart: 'か', notEnd: 'お' };
    // No break between お and か; 」 may start a line, and 「 may end one.
    expect(at('えおかき', lists)).toEqual([1, 3]);
    expect(at('や「括弧」や', lists)).toEqual([1, 2, 3, 4, 5]);
  });

  test('the cascade gives the lists to paragraphs of their language only', () => {
    const r = readPptx(
      deck(para(run('日本語', 'lang="ja-JP"')) + para(run('中文', 'lang="zh-CN"')), 6 * IN, KINSOKU)
    );
    if (!r.ok) throw new Error('unreadable');
    const sp = resolveSlide(r.pkg, 0)!.elements[0] as ResolvedShape;
    const [ja, zh] = sp.text!.paragraphs;
    expect(ja.kinsoku).toEqual({ lang: 'ja-JP', notStart: 'か', notEnd: 'お' });
    expect(zh.kinsoku).toBeUndefined();
  });

  test('a narrow box breaks where the lists allow (reference PDF)', () => {
    // Four characters fit a line: 4 em at 20 pt in a box 80 pt wide.
    const text = (lists: string) =>
      deck(
        para(run('あいうえおかきくけこ', 'lang="ja-JP" sz="2000"')),
        Math.round(80.5 * 12_700),
        lists
      );
    expect(lineTexts(textOf(text(''), recordingMeasurer()))).toEqual([
      'あいうえ',
      'おかきく',
      'けこ',
    ]);
    const kinsoku = '<p:kinsoku lang="ja-JP" invalStChars="か" invalEndChars="く"/>';
    expect(lineTexts(textOf(text(kinsoku), recordingMeasurer()))).toEqual([
      'あいうえ',
      'おかき',
      'くけこ',
    ]);
  });
});

describe('East Asian justification (algn="just", algn="dist")', () => {
  const xs = (layout: TextLayout, line: number): number[] =>
    layout.lines[line].runs.flatMap((g) => g.x.slice(0, g.text.length).map((x) => x / 12_700));

  test('a justified East Asian line spreads its room between all characters', () => {
    // Five 20 pt characters in a 104 pt box: 4 pt of room, 1 pt in each of 4 gaps.
    const layout = textOf(
      deck(
        para(run('日本語の文字と', 'lang="ja-JP" sz="2000"'), '<a:pPr algn="just"/>'),
        104 * 12_700
      ),
      recordingMeasurer()
    );
    expect(lineTexts(layout)).toEqual(['日本語の文', '字と']);
    expect(xs(layout, 0)).toEqual([0, 21, 42, 63, 84]);
    // The last line of a justified paragraph stays left aligned.
    expect(xs(layout, 1)).toEqual([0, 20]);
  });

  test('a space widens to half the font size first, then the characters spread', () => {
    // Spaces 0.25 em (5 pt) wide. Four 20 pt characters and a space in a 100 pt box: 15 pt of
    // room. The space first widens by 5 pt to half the font size, then 10 pt spread over 4 gaps.
    const inner = recordingMeasurer();
    const measurer: TextMeasurer = {
      ...inner,
      advances: (face, text) => {
        const out = inner.advances(face, text);
        for (let i = 0; i < text.length; i += 1) if (text[i] === ' ') out[i] = 0.25;
        return out;
      },
    };
    const layout = textOf(
      deck(para(run('한국 어휘', 'lang="ko-KR" sz="2000"'), '<a:pPr algn="dist"/>'), 100 * 12_700),
      measurer
    );
    expect(lineTexts(layout)).toEqual(['한국 어휘']);
    // 한, 국, and the space, then 어 after the 10 pt space and its 2.5 pt gap.
    expect(xs(layout, 0)).toEqual([0, 22.5, 45, 57.5, 80]);
  });

  test('a line with no East Asian text widens its spaces only', () => {
    const layout = textOf(
      deck(para(run('ab cd ef gh', 'lang="en-US" sz="2000"'), '<a:pPr algn="just"/>'), 60 * 12_700),
      recordingMeasurer()
    );
    const first = xs(layout, 0);
    expect(first[1] - first[0]).toBe(10);
  });
});

describe('glyph fallback', () => {
  test("a character that the run's face lacks takes the measurer's fallback face", () => {
    const inner = recordingMeasurer();
    const measurer: TextMeasurer = {
      ...inner,
      fallbackFace: (face, cp) =>
        cp === 0x2460
          ? inner.resolve({ family: 'Fallback', bold: face.bold, italic: face.italic })
          : undefined,
    };
    const layout = textOf(
      deck(para(run('a①b', 'lang="en-US"', '<a:latin typeface="Arial"/>'))),
      measurer
    );
    expect(layout.lines[0].runs.map((g) => [g.text, g.face.family])).toEqual([
      ['a', 'Arial'],
      ['①', 'Fallback'],
      ['b', 'Arial'],
    ]);
  });
});
