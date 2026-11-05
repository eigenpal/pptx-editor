// The showcase deck's slides 7 to 11 (scripts/sample-deck/showcase.ts builds the deck).

import { table, type FixtureSlide } from '../fixtures/builder.ts';
import {
  IMAGE,
  NO_LINE,
  box,
  color,
  cxn,
  emu,
  fill,
  hardShadow,
  header,
  label,
  line,
  newId,
  p,
  pic,
  r,
  rounded,
  slide,
  stops,
} from './showcase-kit.ts';

// 7. Fills -------------------------------------------------------------------------------------

export function fillsSlide(): FixtureSlide {
  const tiles: [string, string, (x: number, y: number) => string][] = [
    [
      'Solid fill',
      'A theme color: a new theme recolors it',
      (x, y) =>
        box({
          name: 'Solid',
          x,
          y,
          w: 3.63,
          h: 1.55,
          ...rounded(8),
          fill: fill('accent1'),
          line: NO_LINE,
        }),
    ],
    [
      'Linear gradient',
      'Two theme colors, left to right',
      (x, y) =>
        box({
          name: 'Linear',
          x,
          y,
          w: 3.63,
          h: 1.55,
          ...rounded(8),
          fill: `<a:gradFill rotWithShape="1">${stops([0, 'accent1'], [100, 'accent3'])}<a:lin ang="0" scaled="0"/></a:gradFill>`,
          line: NO_LINE,
        }),
    ],
    [
      'Radial gradient',
      'From white at the center to green',
      (x, y) =>
        box({
          name: 'Radial',
          x,
          y,
          w: 3.63,
          h: 1.55,
          ...rounded(8),
          fill: `<a:gradFill rotWithShape="1">${stops([0, 'bg1'], [100, 'accent2'])}<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`,
          line: NO_LINE,
        }),
    ],
    [
      'Three-stop gradient',
      'Blue, violet, and orange at 45°',
      (x, y) =>
        box({
          name: 'Three stops',
          x,
          y,
          w: 3.63,
          h: 1.55,
          ...rounded(8),
          fill: `<a:gradFill rotWithShape="1">${stops([0, 'accent4'], [50, 'accent5'], [100, 'accent1'])}<a:lin ang="2700000" scaled="0"/></a:gradFill>`,
          line: NO_LINE,
        }),
    ],
    [
      'Picture fill',
      'A picture inside any shape, cropped to fit',
      (x, y) =>
        box({
          name: 'Picture fill',
          x,
          y,
          w: 3.63,
          h: 1.55,
          ...rounded(8),
          fill: '<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rIdImg"/><a:srcRect t="20000" b="11700"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>',
          line: NO_LINE,
        }),
    ],
    [
      'Transparency',
      'Theme colors at 55% opacity mix where they overlap',
      (x, y) =>
        ['accent1', 'accent2', 'accent4']
          .map((c, i) =>
            box({
              name: `Circle ${i + 1}`,
              x: x + 0.32 + i * 0.83,
              y: y + 0.08,
              w: 1.4,
              h: 1.4,
              prst: 'ellipse',
              fill: fill(c, 55),
              line: NO_LINE,
            })
          )
          .join(''),
    ],
  ];
  return slide(
    () => ({
      rels: [IMAGE],
      shapes:
        header('Fills', 'Solid, gradient, picture, and transparent fills') +
        tiles
          .map(([name, desc, draw], i) => {
            const x = 0.7 + (i % 3) * 4.15;
            const y = 1.85 + Math.floor(i / 3) * 2.45;
            return (
              draw(x, y) +
              box({
                name: `Caption ${name}`,
                x,
                y: y + 1.6,
                w: 3.63,
                h: 0.62,
                ins: [0, 0.03, 0, 0],
                text: p(r(name, { sz: 14, b: true })) + p(r(desc, { sz: 12, color: 'tx2' })),
              })
            );
          })
          .join(''),
    }),
    [
      'Solid and gradient fills use theme colors, so they follow the theme. The picture fill reuses the picture from the Pictures slide.',
      'Try: select a tile and pick another fill color in the toolbar.',
      'Try: open the fill menu and choose a gradient.',
    ]
  );
}

// 8. Lines -------------------------------------------------------------------------------------

export function linesSlide(): FixtureSlide {
  const weights = [0.75, 1.5, 3, 4.5, 6];
  const dashes: [string, string, string][] = [
    ['Solid', 'solid', ''],
    ['Square dot', 'sysDot', ''],
    ['Round dot', 'sysDot', ' cap="rnd"'],
    ['Dash', 'dash', ''],
    ['Long dash', 'lgDash', ''],
    ['Dash dot', 'dashDot', ''],
  ];
  const rowY = (i: number): number => 2.45 + i * 0.68;
  const caption = (x: number, y: number, w: number, t: string): string =>
    box({
      name: `Caption ${t}`,
      x,
      y: y - 0.2,
      w,
      h: 0.4,
      anchor: 'ctr',
      ins: [0, 0, 0, 0],
      text: p(r(t, { sz: 12, color: 'tx2' })),
    });
  const node = (x: number, y: number, c: string): string =>
    box({ name: 'Node', x, y, w: 0.9, h: 0.45, ...rounded(30), fill: fill(c), line: NO_LINE });
  const connectors: [string, string, number, string][] = [
    [
      'Straight',
      'straightConnector1',
      0,
      '<a:headEnd type="oval" w="lg" len="lg"/><a:tailEnd type="oval" w="lg" len="lg"/>',
    ],
    ['Elbow', 'bentConnector3', 0.55, '<a:tailEnd type="diamond" w="lg" len="lg"/>'],
    [
      'Curved',
      'curvedConnector3',
      0.55,
      '<a:headEnd type="oval" w="lg" len="lg"/><a:tailEnd type="oval" w="lg" len="lg"/>',
    ],
  ];
  return slide(
    () => ({
      shapes:
        header('Lines', 'Line weights, dashes, and connectors') +
        label(0.7, 1.75, 3.63, 'Weight') +
        weights
          .map(
            (w, i) =>
              caption(0.7, rowY(i), 0.9, `${w} pt`) +
              cxn({ name: `Weight ${w}`, x: 1.7, y: rowY(i), w: 2.63, h: 0, line: line(w, 'tx1') })
          )
          .join('') +
        caption(0.7, rowY(5), 0.9, 'Round cap') +
        cxn({
          name: 'Round cap',
          x: 1.85,
          y: rowY(5),
          w: 2.33,
          h: 0,
          line: line(9, 'tx1', '', ' cap="rnd"'),
        }) +
        label(4.85, 1.75, 3.63, 'Dash') +
        dashes
          .map(
            ([name, dash, cap], i) =>
              caption(4.85, rowY(i), 1.3, name) +
              cxn({
                name: `Dash ${name}`,
                x: 6.25,
                y: rowY(i),
                w: 2.23,
                h: 0,
                line: line(3, 'accent1', `<a:prstDash val="${dash}"/>`, cap),
              })
          )
          .join('') +
        label(9.0, 1.75, 3.63, 'Connectors') +
        connectors
          .map(([name, prst, dy, ends], i) => {
            const y = 2.2 + i * 1.5;
            const a = node(9.0, y + 0.35, 'accent2');
            const b = node(11.73, y + 0.35 + dy, 'accent4');
            return (
              caption(9.0, y + 0.1, 3.63, name) +
              a +
              b +
              cxn({
                name: `${name} connector`,
                prst,
                x: 9.9,
                y: y + 0.575,
                w: 1.83,
                h: dy,
                line: line(2, 'tx1', ends),
              })
            );
          })
          .join(''),
    }),
    [
      'Lines take their width, dash, cap, and ends from the file. Connectors join shapes with straight, elbow, and curved paths.',
      'Try: select a line and change its width or dash in the toolbar.',
      'Try: Insert > Line, then drag on the slide to draw one.',
    ]
  );
}

// 9. Effects -----------------------------------------------------------------------------------

export function effectsSlide(): FixtureSlide {
  const cards: [string, string, (x: number) => string][] = [
    [
      'Outer shadow',
      'A hard shadow, 8 pt at 45°, in a theme color',
      (x) =>
        box({
          name: 'Outer shadow',
          x,
          y: 2.15,
          w: 2.6,
          h: 2.0,
          ...rounded(8),
          fill: fill('bg1'),
          line: line(1.5, 'tx1'),
          effects: hardShadow(8, 45, 'accent1'),
        }),
    ],
    [
      'Inner shadow',
      'Cast inside the shape from the top left',
      (x) =>
        box({
          name: 'Inner shadow',
          x,
          y: 2.15,
          w: 2.6,
          h: 2.0,
          ...rounded(8),
          fill: fill('bg2'),
          line: NO_LINE,
          effects: `<a:effectLst><a:innerShdw dist="${8 * 12700}" dir="2700000">${color('tx1', 22)}</a:innerShdw></a:effectLst>`,
        }),
    ],
    [
      'Reflection',
      'A fading mirror image under the shape',
      (x) =>
        box({
          name: 'Reflection',
          x,
          y: 2.15,
          w: 2.6,
          h: 2.0,
          ...rounded(8),
          fill: fill('accent2'),
          line: NO_LINE,
          effects:
            '<a:effectLst><a:reflection stA="50000" endA="0" endPos="45000" dist="50800" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/></a:effectLst>',
        }),
    ],
    [
      'Text shadow',
      'A text box with no fill casts the shadow of its text',
      (x) =>
        box({
          name: 'Text shadow',
          x,
          y: 2.15,
          w: 2.6,
          h: 2.0,
          anchor: 'ctr',
          ins: [0, 0, 0, 0],
          effects: hardShadow(4, 45, 'accent1'),
          text: p(r('Shadow', { sz: 48, b: true }), { algn: 'ctr' }),
        }),
    ],
  ];
  return slide(
    () => ({
      shapes:
        header(
          'Effects',
          'Shadows and reflections',
          'Effects come from the shape’s effect list and draw with the shape and its text.'
        ) +
        cards
          .map(([name, desc, draw], i) => {
            const x = 0.7 + i * 3.11;
            return (
              draw(x) +
              box({
                name: `Caption ${name}`,
                x,
                y: 5.35,
                w: 2.6,
                h: 0.9,
                ins: [0, 0, 0, 0],
                text:
                  p(r(name, { sz: 16, b: true })) +
                  p(r(desc, { sz: 12, color: 'tx2' }), { before: 2 }),
              })
            );
          })
          .join(''),
    }),
    [
      'Shadows and reflections are part of each shape. They move, resize, and rotate with it.',
      'Try: drag the white card. Its orange shadow moves with it.',
      'Try: double-click the word Shadow and type. The shadow follows the new text.',
    ]
  );
}

// 10. Pictures ---------------------------------------------------------------------------------

export function picturesSlide(): FixtureSlide {
  const variants: [string, Partial<Parameters<typeof pic>[0]>][] = [
    ['Crop to a circle', { prst: 'ellipse', crop: [34.75, 0, 2.75, 0], w: 1.5 }],
    ['Grayscale', { blip: '<a:grayscl/>' }],
    [
      'Duotone',
      { blip: '<a:duotone><a:srgbClr val="1B1A19"/><a:srgbClr val="F2A07C"/></a:duotone>' },
    ],
    ['50% transparent', { blip: '<a:alphaModFix amt="50000"/>' }],
  ];
  return slide(
    () => ({
      rels: [IMAGE],
      shapes:
        header('Pictures', 'Pictures with crop, borders, and recoloring') +
        pic({
          name: 'Landscape',
          descr: 'A generated landscape: a dusk sky, a low sun, mountains, and a lake',
          x: 0.7,
          y: 1.85,
          w: 6.4,
          h: 4.0,
          line: line(1.5, 'tx1'),
          effects: hardShadow(8, 45, 'accent1'),
        }) +
        box({
          name: 'Picture caption',
          x: 0.7,
          y: 6.05,
          w: 6.4,
          h: 0.4,
          ins: [0, 0, 0, 0],
          text: p(
            r('A generated 1000 × 625 PNG with a 1.5 pt border and an offset shadow.', {
              sz: 12,
              color: 'tx2',
            })
          ),
        }) +
        variants
          .map(([name, o], i) => {
            const x = 7.65 + (i % 2) * 2.58;
            const y = 1.85 + Math.floor(i / 2) * 2.3;
            const w = o.w ?? 2.4;
            return (
              pic({
                name,
                descr: `The landscape: ${name.toLowerCase()}`,
                x: x + (2.4 - w) / 2,
                y,
                w,
                h: 1.5,
                ...o,
              }) +
              box({
                name: `Caption ${name}`,
                x,
                y: y + 1.58,
                w: 2.4,
                h: 0.35,
                ins: [0, 0, 0, 0],
                text: p(r(name, { sz: 13, b: true }), { algn: 'ctr' }),
              })
            );
          })
          .join(''),
    }),
    [
      'All five pictures use one image part. Crop, shape, and color changes are stored on each picture, so the original stays intact.',
      'Try: drag a corner handle of the large picture. Its aspect ratio stays locked.',
      'Try: select a picture and press Ctrl+D (Cmd+D on a Mac) to duplicate it.',
    ]
  );
}

// 11. Tables -----------------------------------------------------------------------------------

const TABLE_STYLE = '{8A1F3C52-6B0E-4D7A-9C3E-5F2B1A0D4E61}';
const TABLE_STYLE_XML =
  `<a:tblStyle styleId="${TABLE_STYLE}" styleName="Showcase - Accent 1">` +
  `<a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="tx1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/><a:fill>${fill('bg1')}</a:fill></a:tcStyle></a:wholeTbl>` +
  `<a:band1H><a:tcStyle><a:tcBdr/><a:fill>${fill('bg2')}</a:fill></a:tcStyle></a:band1H>` +
  `<a:lastRow><a:tcTxStyle b="on"/><a:tcStyle><a:tcBdr/><a:fill>${fill('accent1', 14)}</a:fill></a:tcStyle></a:lastRow>` +
  `<a:firstRow><a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="bg1"/></a:tcTxStyle><a:tcStyle><a:tcBdr/><a:fill>${fill('tx1')}</a:fill></a:tcStyle></a:firstRow>` +
  '</a:tblStyle>';

interface CellOptions {
  readonly algn?: 'l' | 'ctr' | 'r';
  readonly span?: number;
  readonly rowSpan?: number;
  readonly hMerge?: boolean;
  readonly vMerge?: boolean;
  readonly fill?: string;
  readonly b?: boolean;
  readonly color?: string;
}

function tc(text: string, o: CellOptions = {}): string {
  const attrs =
    (o.span ? ` gridSpan="${o.span}"` : '') +
    (o.rowSpan ? ` rowSpan="${o.rowSpan}"` : '') +
    (o.hMerge ? ' hMerge="1"' : '') +
    (o.vMerge ? ' vMerge="1"' : '');
  const body = text
    ? p(r(text, { sz: 15, b: o.b, color: o.color }), { algn: o.algn })
    : '<a:p><a:endParaRPr lang="en-US" sz="1500" dirty="0"/></a:p>';
  return `<a:tc${attrs}><a:txBody><a:bodyPr/><a:lstStyle/>${body}</a:txBody><a:tcPr marL="${emu(0.15)}" marR="${emu(0.15)}" anchor="ctr">${o.fill ?? ''}</a:tcPr></a:tc>`;
}

export function tablesSlide(): FixtureSlide {
  const rowH = 0.52;
  const body: [string, ...string[]][] = [
    ['Design', '12.0', '14.5', '13.0', '15.5', '55.0'],
    ['Engineering', '48.0', '52.0', '55.5', '58.0', '213.5'],
    ['Marketing', '9.5', '11.0', '16.0', '21.5', '58.0'],
    ['Support', '7.0', '7.5', '8.0', '8.5', '31.0'],
    ['Total', '76.5', '85.0', '92.5', '103.5', '357.5'],
  ];
  const q = (t: string): string =>
    tc(t, { algn: 'r', fill: fill('accent1'), b: true, color: 'bg1' });
  const rows =
    `<a:tr h="${emu(rowH)}">` +
    tc('Category', { rowSpan: 2 }) +
    tc('First half', { span: 2, algn: 'ctr' }) +
    tc('', { hMerge: true }) +
    tc('Second half', { span: 2, algn: 'ctr' }) +
    tc('', { hMerge: true }) +
    tc('Total', { rowSpan: 2, algn: 'r' }) +
    '</a:tr>' +
    `<a:tr h="${emu(rowH)}">${tc('', { vMerge: true })}${['Q1', 'Q2', 'Q3', 'Q4'].map(q).join('')}${tc('', { vMerge: true })}</a:tr>` +
    body
      .map(
        ([name, ...values]) =>
          `<a:tr h="${emu(rowH)}">${tc(name)}${values.map((v) => tc(v, { algn: 'r' })).join('')}</a:tr>`
      )
      .join('');
  const points = [
    'The header row, banded rows, and total row come from the table style.',
    'First half and Second half span two columns. Category and Total span two rows.',
    'Numbers align right; headings align center.',
    'Rows grow when their text needs more room.',
  ];
  return slide(
    () => ({
      shapes:
        header('Tables', 'Table styles, merged cells, and alignment') +
        table({
          id: newId(),
          x: emu(0.7),
          y: emu(1.85),
          cx: emu(8.2),
          cy: emu(rowH * 7),
          columns: [2.2, 1.2, 1.2, 1.2, 1.2, 1.2].map(emu),
          tblPr: `<a:tblPr firstRow="1" lastRow="1" bandRow="1"><a:tableStyleId>${TABLE_STYLE}</a:tableStyleId></a:tblPr>`,
          rows,
        }) +
        box({
          name: 'Table note',
          x: 0.7,
          y: 5.62,
          w: 8.2,
          h: 0.35,
          ins: [0, 0, 0, 0],
          text: p(r('Sample data, in thousands.', { sz: 12, color: 'tx2' })),
        }) +
        label(9.4, 1.85, 3.23, 'What to look for') +
        box({
          name: 'Table points',
          x: 9.4,
          y: 2.25,
          w: 3.23,
          h: 4.2,
          ins: [0, 0, 0, 0],
          text: points
            .map((t) =>
              p(r(t, { sz: 14 }), {
                marL: 0.25,
                indent: -0.25,
                after: 10,
                lnSpc: 100,
                bullet:
                  '<a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="•"/>',
              })
            )
            .join(''),
        }),
    }),
    [
      'The table uses a table style from the file, with the header row, banded rows, and total row turned on.',
      'Try: double-click a number and type a new value. The column keeps its right alignment.',
      'Try: type a long label in a Category cell. The row grows to fit the text.',
    ]
  );
}

export const TABLE_STYLES = TABLE_STYLE_XML;
