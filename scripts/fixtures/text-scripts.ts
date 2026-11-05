// The `rtl-text` and `east-asian-line-breaking` fixtures. `rtl-text` holds right-to-left
// paragraphs (`a:pPr/@rtl`, ECMA-376 Part 1, 21.1.2.2.7 pPr) with Arabic and Hebrew runs, mixed
// directions in one line, numbers in right-to-left text, right-to-left bullets and numbering,
// and alignment and indents. `east-asian-line-breaking` holds CJK text with `eaLnBrk`,
// `latinLnBrk`, and `hangingPunct` on and off, punctuation at line ends, and the `a:ea`, `a:cs`,
// and `a:sym` run fonts (21.1.2.3.3 ea, 21.1.2.3.1 cs, 21.1.2.3.10 sym). Every run names its
// fonts, so no theme fallback decides the face.

import { escapeText, shape, solid, type FixtureSlide } from './builder.ts';

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

const ARIAL = '<a:latin typeface="Arial"/><a:cs typeface="Arial"/>';
const TAHOMA = '<a:latin typeface="Tahoma"/><a:cs typeface="Tahoma"/>';
const DAVID = '<a:latin typeface="Arial"/><a:cs typeface="David"/>';
const ar = (text: string, fonts = ARIAL, attrs?: string): string => r(text, 'ar-SA', fonts, attrs);
const he = (text: string, fonts = ARIAL, attrs?: string): string => r(text, 'he-IL', fonts, attrs);
const en = (text: string, fonts = ARIAL, attrs?: string): string => r(text, 'en-US', fonts, attrs);
const RTL_R = '<a:pPr algn="r" rtl="1"/>';

const rtlText: FixtureSlide[] = [
  {
    // Right-to-left paragraphs in Arabic and Hebrew, right aligned, in three faces.
    shapes:
      box(
        2,
        0.5,
        0.5,
        6,
        3,
        p(ar('هذا نص عربي يبدأ من اليمين إلى اليسار.'), RTL_R) +
          p(ar('تتدفق الفقرة من اليمين، وتلتف الأسطر داخل المربع حسب عرضه دون أن تتغير الكلمات.'), RTL_R) +
          p(ar('خط آخر للنص نفسه في فقرة ثالثة.', TAHOMA), RTL_R)
      ) +
      box(
        3,
        6.9,
        0.5,
        5.9,
        3,
        p(he('זהו טקסט בעברית שנכתב מימין לשמאל.'), RTL_R) +
          p(he('השורות נשברות בתוך התיבה לפי הרוחב שלה, והמילים נשארות שלמות.'), RTL_R) +
          p(he('גופן אחר לאותו טקסט בפסקה שלישית.', DAVID), RTL_R)
      ) +
      box(
        4,
        0.5,
        3.9,
        12.3,
        3,
        p(ar('عنوان كبير بخط عريض', ARIAL, 'sz="3600" b="1"'), RTL_R) +
          p(he('כותרת משנה בגופן נטוי', ARIAL, 'sz="2800" i="1"'), RTL_R) +
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
            p(ar(t), '<a:pPr marL="342900" indent="-342900" algn="r" rtl="1"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>')
          )
          .join('') +
          p(ar('بند فرعي في المستوى الثاني'), '<a:pPr marL="800100" lvl="1" indent="-342900" algn="r" rtl="1"><a:buFont typeface="Arial"/><a:buChar char="–"/></a:pPr>')
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
            p(he(t), '<a:pPr marL="342900" indent="-342900" algn="r" rtl="1"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>')
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

const YU_GOTHIC = '<a:latin typeface="Yu Gothic"/><a:ea typeface="Yu Gothic"/>';
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
          box(6 + i, 0.4 + i * 6.3, 3.8, 2.4, 3.2, p(r(String(text), 'en-US', '<a:latin typeface="Arial"/>', 'sz="2400"'), brk(1, latin as 0 | 1, 1)))
        )
        .join(''),
  },
  {
    // The run fonts: `a:ea` for CJK text, `a:latin` for Latin text in the same run, `a:cs` for
    // Arabic and Hebrew, and `a:sym` for symbol characters (U+F020 to U+F0FF).
    shapes:
      box(
        2,
        0.4,
        0.3,
        6.2,
        6.9,
        p(ja('游ゴシックの日本語 Arial', '<a:latin typeface="Arial"/><a:ea typeface="Yu Gothic"/>')) +
          p(ja('游明朝の日本語 Arial', '<a:latin typeface="Arial"/><a:ea typeface="Yu Mincho"/>')) +
          p(ja('メイリオの日本語 Arial', '<a:latin typeface="Arial"/><a:ea typeface="Meiryo"/>')) +
          p(r('宋体的中文文本 Arial', 'zh-CN', '<a:latin typeface="Arial"/><a:ea typeface="SimSun"/>')) +
          p(r('等线的中文文本 Arial', 'zh-CN', '<a:latin typeface="Arial"/><a:ea typeface="DengXian"/>')) +
          p(r('맑은 고딕 한국어 텍스트 Arial', 'ko-KR', '<a:latin typeface="Arial"/><a:ea typeface="Malgun Gothic"/>')) +
          p(r('Latin text only, with an ea font that it does not use', 'en-US', '<a:latin typeface="Arial"/><a:ea typeface="Yu Mincho"/>'))
      ) +
      box(
        3,
        6.8,
        0.3,
        6.1,
        6.9,
        p(r('نص عربي بخط Tahoma', 'ar-SA', '<a:latin typeface="Arial"/><a:cs typeface="Tahoma"/>'), RTL_R) +
          p(r('טקסט בעברית בגופן David', 'he-IL', '<a:latin typeface="Arial"/><a:cs typeface="David"/>'), RTL_R) +
          p(
            r('Wingdings: ', 'en-US', '<a:latin typeface="Arial"/>') +
              r('', 'en-US', '<a:latin typeface="Arial"/><a:sym typeface="Wingdings"/>')
          ) +
          p(
            r('Symbol: ', 'en-US', '<a:latin typeface="Arial"/>') +
              r('', 'en-US', '<a:latin typeface="Arial"/><a:sym typeface="Symbol"/>')
          ) +
          p(
            ja('混在：', YU_GOTHIC) +
              r('Arial ', 'en-US', '<a:latin typeface="Arial"/>') +
              r('عربي', 'ar-SA', '<a:latin typeface="Arial"/><a:cs typeface="Tahoma"/>') +
              ja(' 日本語', YU_GOTHIC)
          )
      ),
  },
];

export const textScriptFixtures: Record<string, FixtureSlide[]> = {
  'rtl-text': rtlText,
  'east-asian-line-breaking': eastAsian,
};
