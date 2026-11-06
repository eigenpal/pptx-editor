// The `rtl-text` and `east-asian-line-breaking` fixtures. `rtl-text` holds right-to-left
// paragraphs (`a:pPr/@rtl`, ECMA-376 Part 1, 21.1.2.2.7 pPr) with Arabic and Hebrew runs, mixed
// directions in one line, numbers in right-to-left text, right-to-left bullets and numbering,
// and alignment and indents. `east-asian-line-breaking` holds CJK text with `eaLnBrk`,
// `latinLnBrk`, and `hangingPunct` on and off, punctuation at line ends, and the `a:ea` and
// `a:cs` run fonts (21.1.2.3.3 ea, 21.1.2.3.1 cs). Every run names its open fonts, so no theme
// fallback decides the face. `east-asian-theme-fonts` names no run font: its
// theme's `a:ea` is empty, and the theme script fonts (20.1.4.1.16 font) with each run's `lang`
// and `altLang` decide the face. It also holds shared symbols in East Asian and Latin runs, and
// Korean and Chinese text in narrow boxes. `east-asian-typography` holds custom `p:kinsoku`
// lists and justified and distributed East Asian text.

import { escapeText, shape, solid, type FixtureOptions, type FixtureSlide } from './builder.ts';

const IN = 914_400;

/** One run with its language and fonts. `fonts` is `a:latin`, `a:ea`, `a:cs`, `a:sym` XML. */
function r(text: string, lang: string, fonts: string, attrs = 'sz="2000"'): string {
  return `<a:r><a:rPr lang="${lang}" ${attrs} dirty="0">${fonts}</a:rPr><a:t>${escapeText(text)}</a:t></a:r>`;
}

const p = (runs: string, pPr = ''): string => `<a:p>${pPr}${runs}</a:p>`;

/** A text box with a fixed frame and no autofit, so the frame decides every line break. */
function box(id: number, x: number, y: number, cx: number, cy: number, text: string, fill = 'F7F7F7'): string {
  return shape({
    id,
    x: x * IN,
    y: y * IN,
    cx: cx * IN,
    cy: cy * IN,
    fill: solid(fill),
    bodyPr: '<a:bodyPr wrap="square" rtlCol="0" anchor="t"><a:noAutofit/></a:bodyPr>',
    text,
  });
}

// --- rtl-text ----------------------------------------------------------------

const LATIN = '<a:latin typeface="Liberation Sans"/><a:cs typeface="Liberation Sans"/>';
/** Arabic: Liberation Sans has no Arabic glyphs, so the Arabic face is Noto Sans Arabic. */
const ARABIC = '<a:latin typeface="Liberation Sans"/><a:cs typeface="Noto Sans Arabic"/>';
const HEBREW_SERIF = '<a:latin typeface="Liberation Serif"/><a:cs typeface="Liberation Serif"/>';
const ar = (text: string, fonts = ARABIC, attrs?: string): string => r(text, 'ar-SA', fonts, attrs);
const he = (text: string, fonts = LATIN, attrs?: string): string => r(text, 'he-IL', fonts, attrs);
const en = (text: string, fonts = LATIN, attrs?: string): string => r(text, 'en-US', fonts, attrs);
const RTL_R = '<a:pPr algn="r" rtl="1"/>';

const rtlText: FixtureSlide[] = [
  {
    // Right-to-left paragraphs in Arabic and Hebrew, right aligned; the last Hebrew paragraph in a serif face.
    shapes:
      box(
        2,
        0.5,
        0.5,
        6,
        3,
        p(ar('هذا نص عربي يبدأ من اليمين إلى اليسار.'), RTL_R) +
          p(ar('تتدفق الفقرة من اليمين، وتلتف الأسطر داخل المربع حسب عرضه دون أن تتغير الكلمات.'), RTL_R) +
          p(ar('خط آخر للنص نفسه في فقرة ثالثة.', ARABIC), RTL_R)
      ) +
      box(
        3,
        6.9,
        0.5,
        5.9,
        3,
        p(he('זהו טקסט בעברית שנכתב מימין לשמאל.'), RTL_R) +
          p(he('השורות נשברות בתוך התיבה לפי הרוחב שלה, והמילים נשארות שלמות.'), RTL_R) +
          p(he('גופן אחר לאותו טקסט בפסקה שלישית.', HEBREW_SERIF), RTL_R)
      ) +
      box(
        4,
        0.5,
        3.9,
        12.3,
        3,
        p(ar('عنوان كبير بخط عريض', ARABIC, 'sz="3600" b="1"'), RTL_R) +
          p(he('כותרת משנה בגופן נטוי', LATIN, 'sz="2800" i="1"'), RTL_R) +
          p(ar('فقرة من اليمين إلى اليسار بمحاذاة إلى اليسار.'), '<a:pPr algn="l" rtl="1"/>') +
          p(en('A left-to-right paragraph in the same box, aligned left.'), '<a:pPr algn="l" rtl="0"/>')
      ),
  },
  {
    // Mixed directions in one line, and numbers in right-to-left text.
    shapes:
      box(
        2,
        0.5,
        0.5,
        6,
        3.1,
        p(ar('يحفظ المحرر ملفات ') + en('PPTX') + ar(' دون فقدان أي بيانات.'), RTL_R) +
          p(he('העורך פותח קבצי ') + en('PPTX') + he(' ושומר אותם לפי ') + en('ECMA-376') + he('.'), RTL_R) +
          p(ar('اسم الملف هو ') + en('report final.pptx') + ar(' في المجلد.'), RTL_R)
      ) +
      box(
        3,
        6.9,
        0.5,
        5.9,
        3.1,
        p(en('The word ') + ar('سلام') + en(' means peace, and ') + he('שלום') + en(' means the same.'), '<a:pPr algn="l" rtl="0"/>') +
          p(en('An English phrase in a right-to-left paragraph: ') + ar('مرحبا') + en('.'), RTL_R) +
          p(he('ציטוט: ') + en('"Hello, world!"') + he(' בתוך משפט.'), RTL_R)
      ) +
      box(
        4,
        0.5,
        3.9,
        12.3,
        3,
        p(ar('بلغ العدد 2026 في السنة، بنسبة 45.5% من 1,200 مشارك.'), RTL_R) +
          p(ar('الصفحة ١٢ من ٣٤، والتاريخ ٢٠٢٦/١١/٠٥.'), RTL_R) +
          p(he('המחיר הוא 120.50 ש״ח, הנחה של 15% עד 31.12.2026.'), RTL_R) +
          p(he('טלפון 03-1234567, חדר 4ב, קומה -2.'), RTL_R)
      ),
  },
  {
    // Right-to-left bullets and numbering, with a second level.
    shapes:
      box(
        2,
        0.5,
        0.5,
        4,
        3.4,
        ['البند الأول من القائمة', 'البند الثاني أطول قليلا ويلتف إلى سطر جديد', 'البند الثالث']
          .map((t) =>
            p(ar(t), '<a:pPr marL="342900" indent="-342900" algn="r" rtl="1"><a:buFont typeface="Liberation Sans"/><a:buChar char="•"/></a:pPr>')
          )
          .join('') +
          p(ar('بند فرعي في المستوى الثاني'), '<a:pPr marL="800100" lvl="1" indent="-342900" algn="r" rtl="1"><a:buFont typeface="Liberation Sans"/><a:buChar char="–"/></a:pPr>')
      ) +
      box(
        3,
        4.7,
        0.5,
        4,
        3.4,
        ['الخطوة الأولى', 'الخطوة الثانية', 'الخطوة الثالثة', 'الخطوة الرابعة']
          .map((t) =>
            p(ar(t), '<a:pPr marL="457200" indent="-457200" algn="r" rtl="1"><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/></a:pPr>')
          )
          .join('')
      ) +
      box(
        4,
        8.9,
        0.5,
        3.9,
        3.4,
        ['שלב ראשון', 'שלב שני', 'שלב שלישי', 'שלב רביעי']
          .map((t) =>
            p(he(t), '<a:pPr marL="457200" indent="-457200" algn="r" rtl="1"><a:buFont typeface="+mj-lt"/><a:buAutoNum type="hebrew2Minus"/></a:pPr>')
          )
          .join('')
      ) +
      box(
        5,
        0.5,
        4.2,
        12.3,
        2.8,
        ['פריט עם תבליט', 'פריט נוסף']
          .map((t) =>
            p(he(t), '<a:pPr marL="342900" indent="-342900" algn="r" rtl="1"><a:buFont typeface="Liberation Sans"/><a:buChar char="•"/></a:pPr>')
          )
          .join('') +
          p(ar('أ. بند مرقم بالحروف'), '<a:pPr marL="457200" indent="-457200" algn="r" rtl="1"><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabic1Minus"/></a:pPr>') +
          p(ar('ب. بند مرقم آخر'), '<a:pPr marL="457200" indent="-457200" algn="r" rtl="1"><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabic1Minus"/></a:pPr>')
      ),
  },
  {
    // Alignment and indents of right-to-left paragraphs.
    shapes:
      box(
        2,
        0.5,
        0.5,
        6,
        3.2,
        p(ar('محاذاة إلى اليمين.'), RTL_R) +
          p(ar('محاذاة في الوسط.'), '<a:pPr algn="ctr" rtl="1"/>') +
          p(ar('محاذاة إلى اليسار.'), '<a:pPr algn="l" rtl="1"/>') +
          p(ar('فقرة مضبوطة الطرفين تمتد على أكثر من سطر واحد حتى يظهر توزيع المسافات بين الكلمات.'), '<a:pPr algn="just" rtl="1"/>')
      ) +
      box(
        3,
        6.9,
        0.5,
        5.9,
        3.2,
        p(he('הזחה של שורה ראשונה בפסקה שנמשכת על פני כמה שורות בתוך התיבה.'), '<a:pPr marL="0" indent="457200" algn="r" rtl="1"/>') +
          p(he('הזחה תלויה בפסקה שנמשכת גם היא על פני כמה שורות בתוך התיבה.'), '<a:pPr marL="685800" indent="-457200" algn="r" rtl="1"/>')
      ) +
      box(
        4,
        0.5,
        4,
        12.3,
        3,
        p(ar('هامش أيمن وأيسر: فقرة بهامش من الجهتين.'), '<a:pPr marL="914400" marR="1828800" algn="r" rtl="1"/>') +
          p(ar('فقرة بمسافة قبلها وبعدها.'), '<a:pPr algn="r" rtl="1"><a:spcBef><a:spcPts val="1200"/></a:spcBef><a:spcAft><a:spcPts val="1200"/></a:spcAft></a:pPr>') +
          p(he('פסקה מיושרת לימין עם שוליים מימין.'), '<a:pPr marL="1828800" algn="r" rtl="1"/>')
      ),
  },
];

// --- east-asian-line-breaking ------------------------------------------------

const YU_GOTHIC = '<a:latin typeface="Noto Sans CJK JP"/><a:ea typeface="Noto Sans CJK JP"/>';
const ja = (text: string, fonts = YU_GOTHIC, attrs?: string): string => r(text, 'ja-JP', fonts, attrs);
const brk = (ea: 0 | 1, latin: 0 | 1, hang: 0 | 1): string =>
  `<a:pPr eaLnBrk="${ea}" latinLnBrk="${latin}" hangingPunct="${hang}"/>`;

const JA_LONG =
  '日本語の文章は単語の間に空白がないため、どの文字の間でも改行できます。英単語のPresentationやEditorは途中で分けません。';
const KINSOKU =
  'これは禁則処理の例です。「括弧」や（丸括弧）、句点。読点、長音ー、小さいっやゅ、感嘆符！疑問符？は行頭に来ません。';
const LATIN_MIX = 'このファイルはinternationalizationとlocalizationの両方に対応しています。';

/** A row of boxes with the same text: one per `a:pPr`, at one width. */
const row = (id: number, y: number, cx: number, text: string, pPrs: readonly string[]): string =>
  pPrs.map((pPr, i) => box(id + i, 0.4 + i * (cx + 0.2), y, cx, 3.2, p(ja(text, YU_GOTHIC, 'sz="2000"'), pPr), i % 2 ? 'EEF4FB' : 'FBF4EE')).join('');

const eastAsian: FixtureSlide[] = [
  {
    // `eaLnBrk` on and off at two widths: on breaks between any two CJK characters, off keeps
    // runs of CJK text together where it can.
    shapes:
      row(2, 0.3, 2.9, JA_LONG, [brk(1, 0, 1), brk(0, 0, 1), brk(1, 1, 1), brk(0, 1, 1)]) +
      row(6, 3.8, 3.05, JA_LONG, [brk(1, 0, 1), brk(0, 0, 1), brk(1, 1, 1), brk(0, 1, 1)]),
  },
  {
    // `hangingPunct` on and off: punctuation that may hang at the line end, and characters
    // that may not start a line, at two widths.
    shapes:
      row(2, 0.3, 2.9, KINSOKU, [brk(1, 0, 1), brk(1, 0, 0), brk(0, 0, 1), brk(0, 0, 0)]) +
      row(6, 3.8, 3.05, KINSOKU, [brk(1, 0, 1), brk(1, 0, 0), brk(0, 0, 1), brk(0, 0, 0)]),
  },
  {
    // `latinLnBrk` on and off: on may break a Latin word, off breaks Latin text at spaces.
    shapes:
      row(2, 0.3, 2.9, LATIN_MIX, [brk(1, 1, 1), brk(1, 0, 1), brk(0, 1, 1), brk(0, 0, 1)]) +
      [
        [1, 'Internationalization and localization in a narrow box, with latinLnBrk on.'],
        [0, 'Internationalization and localization in a narrow box, with latinLnBrk off.'],
      ]
        .map(([latin, text], i) =>
          box(6 + i, 0.4 + i * 6.3, 3.8, 2.4, 3.2, p(r(String(text), 'en-US', '<a:latin typeface="Liberation Sans"/>', 'sz="2400"'), brk(1, latin as 0 | 1, 1)))
        )
        .join(''),
  },
  {
    // The run fonts: `a:ea` for CJK text, `a:latin` for Latin text in the same run, and `a:cs`
    // for Arabic and Hebrew. Each slot names another open face, so the face shows the slot.
    shapes:
      box(
        2,
        0.4,
        0.3,
        6.2,
        6.9,
        p(ja('日本語の文 Liberation Serif', '<a:latin typeface="Liberation Serif"/><a:ea typeface="Noto Sans CJK JP"/>')) +
          p(ja('日本語の文 Liberation Sans', '<a:latin typeface="Liberation Sans"/><a:ea typeface="Noto Sans CJK JP"/>')) +
          p(r('简体中文的文本 Liberation Serif', 'zh-CN', '<a:latin typeface="Liberation Serif"/><a:ea typeface="Noto Sans CJK JP"/>')) +
          p(r('한국어 텍스트 Liberation Serif', 'ko-KR', '<a:latin typeface="Liberation Serif"/><a:ea typeface="Noto Sans CJK JP"/>')) +
          p(r('Latin text only, with an ea font that it does not use', 'en-US', '<a:latin typeface="Liberation Serif"/><a:ea typeface="Noto Sans CJK JP"/>'))
      ) +
      box(
        3,
        6.8,
        0.3,
        6.1,
        6.9,
        p(r('نص عربي بخط Noto Sans Arabic', 'ar-SA', '<a:latin typeface="Liberation Sans"/><a:cs typeface="Noto Sans Arabic"/>'), RTL_R) +
          p(r('טקסט בעברית בגופן Liberation Serif', 'he-IL', '<a:latin typeface="Liberation Sans"/><a:cs typeface="Liberation Serif"/>'), RTL_R) +
          p(
            ja('混在：') +
              r('Liberation Sans ', 'en-US', '<a:latin typeface="Liberation Sans"/>') +
              r('عربي', 'ar-SA', '<a:latin typeface="Liberation Sans"/><a:cs typeface="Noto Sans Arabic"/>') +
              ja(' 日本語')
          )
      ),
  },
];

// --- east-asian-theme-fonts ---------------------------------------------------

/** The theme script fonts of `east-asian-theme-fonts`, in both theme font collections. */
const THEME_SCRIPT_FONTS: readonly (readonly [string, string])[] = [
  ['Jpan', 'Noto Sans CJK JP'],
  ['Hans', 'Noto Sans CJK JP'],
  ['Hant', 'Noto Sans CJK JP'],
  ['Hang', 'Noto Sans CJK JP'],
  ['Arab', 'Noto Sans Arabic'],
];

/** A run that names no font: the theme decides each slot. */
const lr = (text: string, lang: string, altLang?: string): string =>
  r(text, lang, '', `sz="2400"${altLang ? ` altLang="${altLang}"` : ''}`);

/** The `east-asian-theme-fonts` deck: the theme names the East Asian faces by script. */
export function eastAsianThemeFontsDeck(): FixtureOptions {
  return {
    majorFont: 'Liberation Serif',
    minorFont: 'Liberation Serif',
    scriptFonts: THEME_SCRIPT_FONTS,
    slides: [
      {
        shapes:
          // Each language takes its script font; Latin text in the same run stays Liberation Serif.
          box(
            2,
            0.4,
            0.3,
            6.2,
            4.2,
            p(lr('日本語のテーマ Latin', 'ja-JP')) +
              p(lr('简体中文的主题 Latin', 'zh-CN')) +
              p(lr('繁體中文的主題 Latin', 'zh-TW')) +
              p(lr('한국어 테마 글꼴 Latin', 'ko-KR')) +
              p(lr('代替言語の日本語 Latin', 'en-US', 'ja-JP')) +
              p(lr('نص عربي Latin', 'ar-SA'), RTL_R)
          ) +
          // Shared symbols: curly quotes, the ellipsis, and arrows in a Japanese and a Latin run.
          // Only symbols that the open faces hold: any other symbol takes a fallback face.
          box(
            3,
            6.8,
            0.3,
            6.1,
            2,
            p(lr('“引用”…→ ①②③', 'ja-JP')) + p(lr('“Quote”…→', 'en-US'))
          ) +
          // Korean and Chinese text in narrow boxes: where each line breaks.
          box(
            4,
            6.8,
            2.6,
            2.9,
            4.6,
            p(lr('한국어 문장은 띄어쓰기로 단어를 나눕니다. 좁은 상자에서 줄이 바뀝니다.', 'ko-KR'))
          ) +
          box(
            5,
            10,
            2.6,
            2.9,
            4.6,
            p(lr('中文句子没有空格，每个字之间都可以换行。“引号”和（括号）不在行首。', 'zh-CN'))
          ),
      },
    ],
  };
}

// --- east-asian-typography ---------------------------------------------------

/**
 * The deck's `p:kinsoku` (ECMA-376 Part 1, 19.2.1.17): custom lists for Japanese. The first
 * kana of each kana row may not start a line, and the last may not end one. `。` and `」` are in
 * neither list.
 */
const KINSOKU_LISTS = '<p:kinsoku lang="ja-JP" invalStChars="かさたなはまやらわ" invalEndChars="おこそとのほもよろ"/>';
const KANA = 'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん';
const STOPS = 'テストです。「次」の文。テストです。「次」の文。テストです。「次」の文。';
const JUSTIFY =
  '両端揃えの段落では、行の文字の間隔が広がります。英語のwordsと数字123、（括弧）と「かぎ括弧」が混ざっています。';

/** The width in inches of a box that holds `n` 20 pt full-width characters per line. */
const fits = (n: number): number => Math.round(((n * 20 + 14.4 + 4) / 72) * 100) / 100;

/** A row of boxes with the same Japanese text, from `from` to `from + count - 1` characters wide. */
function widths(id: number, y: number, from: number, count: number, text: string, height: number): string {
  let x = 0.3;
  let out = '';
  for (let i = 0; i < count; i += 1) {
    const cx = fits(from + i);
    out += box(id + i, x, y, cx, height, p(ja(text)), i % 2 ? 'EEF4FB' : 'FBF4EE');
    x += cx + 0.35;
  }
  return out;
}

export function eastAsianTypographyDeck(): FixtureOptions {
  return {
    presentationLists: KINSOKU_LISTS,
    slides: [
      {
        // The custom kinsoku lists: the kana rows at five widths, then the stops, which the
        // custom lists do not name, at five widths.
        shapes: widths(2, 0.2, 5, 5, KANA, 3.4) + widths(7, 3.8, 5, 5, STOPS, 3.4),
      },
      {
        // `algn="just"` and `algn="dist"` with Japanese text at two widths, and Korean text.
        shapes:
          [
            ['just', 3.4],
            ['dist', 3.4],
            ['just', 4.6],
            ['dist', 4.6],
          ]
            .map(([algn, cx], i) =>
              box(2 + i, 0.3 + (i % 2) * (Number(cx) + 0.5), i < 2 ? 0.2 : 3.0, Number(cx), 2.6, p(ja(JUSTIFY), `<a:pPr algn="${algn}"/>`), i % 2 ? 'EEF4FB' : 'FBF4EE')
            )
            .join('') +
          box(
            6,
            9.8,
            0.2,
            3.2,
            5.4,
            p(r('한국어 문단을 양쪽 맞춤으로 정렬하면 단어 사이와 글자 사이의 간격이 넓어집니다.', 'ko-KR', '<a:latin typeface="Liberation Sans"/><a:ea typeface="Noto Sans CJK JP"/>'), '<a:pPr algn="just"/>') +
              p(r('균등 분할 정렬은 마지막 줄도 채웁니다.', 'ko-KR', '<a:latin typeface="Liberation Sans"/><a:ea typeface="Noto Sans CJK JP"/>'), '<a:pPr algn="dist"/>')
          ),
      },
    ],
  };
}

export const textScriptFixtures: Record<string, FixtureSlide[]> = {
  'rtl-text': rtlText,
  'east-asian-line-breaking': eastAsian,
};
