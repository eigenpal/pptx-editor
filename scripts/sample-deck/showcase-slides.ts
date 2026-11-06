// The showcase deck's slides 1 to 6 (scripts/sample-deck/showcase.ts builds the deck).

import type { FixtureSlide } from '../fixtures/builder.ts';
import {
  DARK_SLIDE,
  FAINT_TEXT,
  LIGHT_TEXT,
  NO_LINE,
  box,
  color,
  cxn,
  fill,
  group,
  header,
  label,
  line,
  listLevel,
  p,
  r,
  rounded,
  slide,
} from './showcase-kit.ts';

// 1. Title -------------------------------------------------------------------------------------

export function titleSlide(): FixtureSlide {
  return slide(
    () => ({
      ...DARK_SLIDE,
      shapes:
        box({
          name: 'Kicker',
          x: 0.8,
          y: 2.0,
          w: 6.5,
          h: 0.36,
          ins: [0, 0, 0, 0],
          anchor: 'b',
          text: p(
            r('EigenPal editor family', { sz: 14, b: true, caps: true, spc: 3, color: 'accent2' })
          ),
        }) +
        box({
          name: 'Title',
          ph: '<p:ph type="title"/>',
          x: 0.8,
          y: 2.38,
          w: 7.2,
          h: 1.25,
          ins: [0, 0.05, 0, 0.05],
          anchor: 'b',
          text: p(r('PPTX Editor', { sz: 72, b: true, color: 'tx1' })),
        }) +
        box({
          name: 'Subtitle',
          x: 0.8,
          y: 3.7,
          w: 7.2,
          h: 1.0,
          ins: [0, 0.05, 0, 0.05],
          text:
            p(r('Faithful PPTX editing', { sz: 28, color: LIGHT_TEXT })) +
            p(r('in the browser', { sz: 28, color: LIGHT_TEXT })),
        }) +
        box({
          name: 'Accent bar',
          x: 0.8,
          y: 5.0,
          w: 1.1,
          h: 0.07,
          fill: fill('accent1'),
          line: NO_LINE,
        }) +
        box({
          name: 'Site',
          x: 0.8,
          y: 5.25,
          w: 6,
          h: 0.4,
          ins: [0, 0.05, 0, 0.05],
          text: p(r('pptx-editor.dev', { sz: 16, color: FAINT_TEXT })),
        }) +
        // Three slide cards: two rotated behind, and a group in front that draws a tiny slide.
        box({
          name: 'Back card',
          x: 8.15,
          y: 2.45,
          w: 3.9,
          h: 2.19,
          rot: -14,
          ...rounded(6),
          fill: fill('accent2'),
          line: NO_LINE,
        }) +
        box({
          name: 'Middle card',
          x: 8.3,
          y: 2.58,
          w: 3.9,
          h: 2.19,
          rot: -6,
          ...rounded(6),
          fill: fill('accent1'),
          line: NO_LINE,
        }) +
        group({ name: 'Front card', x: 8.45, y: 2.72, w: 3.9, h: 2.19, rot: 3 }, () =>
          [
            box({
              name: 'Card',
              x: 8.45,
              y: 2.72,
              w: 3.9,
              h: 2.19,
              ...rounded(6),
              fill: fill('tx1'),
              line: NO_LINE,
            }),
            box({
              name: 'Card title',
              x: 8.75,
              y: 3.0,
              w: 1.9,
              h: 0.18,
              fill: fill('bg1'),
              line: NO_LINE,
            }),
            box({
              name: 'Card text 1',
              x: 8.75,
              y: 3.34,
              w: 1.5,
              h: 0.08,
              fill: fill(LIGHT_TEXT),
              line: NO_LINE,
            }),
            box({
              name: 'Card text 2',
              x: 8.75,
              y: 3.52,
              w: 1.15,
              h: 0.08,
              fill: fill(LIGHT_TEXT),
              line: NO_LINE,
            }),
            ...[0.45, 0.75, 0.6].map((h, i) =>
              box({
                name: `Bar ${i + 1}`,
                x: 8.75 + i * 0.42,
                y: 4.62 - h,
                w: 0.3,
                h,
                fill: fill('accent2'),
                line: NO_LINE,
              })
            ),
            box({
              name: 'Pie back',
              x: 11.0,
              y: 3.42,
              w: 1.0,
              h: 1.0,
              prst: 'ellipse',
              fill: fill('accent1', 22),
              line: NO_LINE,
            }),
            box({
              name: 'Pie',
              x: 11.0,
              y: 3.42,
              w: 1.0,
              h: 1.0,
              prst: 'pie',
              fill: fill('accent1'),
              line: NO_LINE,
            }),
          ].join('')
        ),
    }),
    [
      'Welcome to the PPTX Editor showcase. Every slide is an ordinary PPTX slide, and you can edit all of it.',
      'Try: double-click the title and type.',
      'Try: click the small slide on the right. It is a group: double-click it to select one part, and press Esc to leave the group.',
      'Try: press F5 to present from the start, and Esc to end the show.',
    ]
  );
}

// 2. Pipeline ----------------------------------------------------------------------------------

export function pipelineSlide(): FixtureSlide {
  const steps: [string, string, string, string][] = [
    ['Open', 'bounded zip and XML', 'tx1', 'bg1'],
    ['Store', 'the only write path', 'accent1', 'bg1'],
    ['Cascade', 'theme and inheritance', 'bg2', 'tx1'],
    ['Layout', 'HarfBuzz text', 'bg2', 'tx1'],
    ['Paint', 'SVG you edit on', 'bg2', 'tx1'],
    ['Save', 'lossless .pptx', 'tx1', 'bg1'],
  ];
  const w = 2.2;
  const h = 1.3;
  const step = 1.942;
  const top = 2.1;
  const cx = (i: number): number => 0.7 + i * step + w / 2;
  return slide(() => {
    const chevrons = steps
      .map(([name, sub, bg, fg], i) =>
        box({
          name: `Step ${name}`,
          x: 0.7 + i * step,
          y: top,
          w,
          h,
          prst: i === 0 ? 'homePlate' : 'chevron',
          avLst: '<a:gd name="adj" fmla="val 26000"/>',
          fill: fill(bg),
          line: NO_LINE,
          anchor: 'ctr',
          ins: [0.05, 0.05, 0.05, 0.05],
          text:
            p(r(name, { sz: 16, b: true, color: fg }), { algn: 'ctr' }) +
            p(r(sub, { sz: 11, color: fg === 'tx1' ? 'tx2' : LIGHT_TEXT }), { algn: 'ctr' }),
        })
      )
      .join('');
    const editId = 30;
    const loop = line(2, 'accent1', '<a:prstDash val="sysDash"/>');
    const edit = box({
      id: editId,
      name: 'Edit',
      x: cx(4) - 1.0,
      y: 4.0,
      w: 2.0,
      h: 0.55,
      ...rounded(50),
      fill: fill('accent1', 12),
      line: line(1, 'accent1'),
      anchor: 'ctr',
      text: p(r('Edit on the slide', { sz: 13, b: true, color: 'accent1' }), { algn: 'ctr' }),
    });
    const cards: [string, string][] = [
      ['187', 'preset shapes, each drawn from the formulas in the ECMA-376 standard'],
      ['0', 'bytes uploaded: your deck opens, renders, and saves in the browser'],
      ['100%', 'of the content the editor does not know stays in the file when you save'],
    ];
    return {
      shapes:
        header(
          'How it works',
          'One pipeline from file to screen and back',
          'The editor and the fidelity tests run the same steps. There is no second renderer.'
        ) +
        chevrons +
        edit +
        cxn({
          name: 'Paint to edit',
          x: cx(4),
          y: top + h,
          w: 0,
          h: 4.0 - top - h,
          line: line(2, 'accent1', '<a:prstDash val="sysDash"/><a:headEnd type="oval"/>'),
          end: [editId, 0],
        }) +
        cxn({
          name: 'Edit to store',
          prst: 'bentConnector2',
          x: cx(1),
          y: top + h,
          w: cx(4) - 1.0 - cx(1),
          h: 4.275 - top - h,
          flipH: true,
          flipV: true,
          line: loop.replace('</a:ln>', '<a:tailEnd type="oval"/></a:ln>'),
          start: [editId, 1],
        }) +
        box({
          name: 'Loop label',
          x: cx(1) + 0.15,
          y: 3.55,
          w: cx(4) - cx(1) - 0.3,
          h: 0.4,
          anchor: 'ctr',
          ins: [0, 0, 0, 0],
          text: p(
            r('Edits are store operations, so undo, redo, and save stay exact.', {
              sz: 13,
              color: 'tx2',
            }),
            { algn: 'ctr' }
          ),
        }) +
        cards
          .map(([big, text], i) =>
            box({
              name: `Fact ${i + 1}`,
              x: 0.7 + i * 4.09,
              y: 5.0,
              w: 3.75,
              h: 1.65,
              ...rounded(8),
              fill: fill('bg2'),
              line: NO_LINE,
              anchor: 'ctr',
              ins: [0.3, 0.15, 0.3, 0.15],
              text: p(r(big, { sz: 32, b: true, color: 'accent1' })) + p(r(text, { sz: 14 })),
            })
          )
          .join(''),
    };
  }, [
    'PPTX Editor reads the file into a document store. The cascade resolves themes and inheritance, layout places every line, and paint draws SVG.',
    'Try: click a chevron and drag it. Smart guides appear when its edges or center line up with another shape.',
    'Try: select the dashed connector, then press Delete, then Ctrl+Z (Cmd+Z on a Mac) to bring it back.',
  ]);
}

// 3. Character formatting ----------------------------------------------------------------------

export function charactersSlide(): FixtureSlide {
  const gap = (sz: number): string => r('   ', { sz });
  const rows: [string, string][] = [
    [
      'Style',
      [
        r('Bold', { sz: 22, b: true }),
        r('Italic', { sz: 22, i: true }),
        r('Underline', { sz: 22, u: 'sng' }),
        r('Double underline', { sz: 22, u: 'dbl' }),
        r('Strikethrough', { sz: 22, strike: 'sngStrike' }),
      ].join(gap(22)),
    ],
    [
      'Font',
      ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New']
        .map((f) => r(f, { sz: 22, font: f }))
        .join(gap(22)),
    ],
    ['Size', [12, 18, 24, 32, 40].map((s) => r(`${s} pt`, { sz: s })).join(gap(18))],
    [
      'Theme color',
      (
        [
          ['Orange', 'accent1'],
          ['Green', 'accent2'],
          ['Amber', 'accent3'],
          ['Blue', 'accent4'],
          ['Violet', 'accent5'],
          ['Teal', 'accent6'],
        ] as const
      )
        .map(([n, c]) => r(n, { sz: 22, b: true, color: c }))
        .join(gap(22)),
    ],
    [
      'Highlight',
      r('Mark ', { sz: 22 }) +
        r('the words that matter', { sz: 22, highlight: 'FFE58A' }) +
        r(', in ', { sz: 22 }) +
        r('any color', { sz: 22, highlight: 'C6EFC6' }) +
        r('.', { sz: 22 }),
    ],
    [
      'Spacing',
      [
        r('Expanded by 3 pt', { sz: 22, spc: 3 }),
        r('Condensed by 1 pt', { sz: 22, spc: -1 }),
        r('All caps', { sz: 22, caps: true, spc: 1 }),
      ].join(gap(22)),
    ],
  ];
  const top = 1.95;
  const rowH = 0.76;
  return slide(
    () => ({
      shapes:
        header(
          'Text',
          'Character formatting, measured with real font metrics',
          'Every run is shaped with HarfBuzz and the real font, so each line breaks at the measured width.'
        ) +
        Array.from({ length: rows.length + 1 }, (_, i) =>
          cxn({
            name: `Rule ${i + 1}`,
            x: 0.7,
            y: top + i * rowH,
            w: 11.93,
            h: 0,
            line: line(0.75, LIGHT_TEXT),
          })
        ).join('') +
        rows
          .map(
            ([name, runs], i) =>
              box({
                name: `Label ${name}`,
                x: 0.7,
                y: top + i * rowH,
                w: 2.0,
                h: rowH,
                anchor: 'ctr',
                ins: [0, 0, 0, 0],
                text: p(r(name, { sz: 12, b: true, caps: true, spc: 1.5, color: 'tx2' })),
              }) +
              box({
                name: `Sample ${name}`,
                x: 2.8,
                y: top + i * rowH,
                w: 9.83,
                h: rowH,
                anchor: 'ctr',
                ins: [0, 0, 0, 0],
                text: p(runs, { lnSpc: 100 }),
              })
          )
          .join(''),
    }),
    [
      'Each row shows one kind of character formatting. Fonts that cannot ship with the editor use metric-compatible substitutes, so line breaks stay the same.',
      'Try: double-click a word and press Ctrl+B, Ctrl+I, or Ctrl+U (Cmd on a Mac).',
      'Try: select text and pick a font, a size, or a theme color in the toolbar.',
    ]
  );
}

// 4. Paragraphs --------------------------------------------------------------------------------

export function paragraphsSlide(): FixtureSlide {
  const text =
    'Text stays aligned while you type. Each line breaks at the measured width, and justified lines stretch to the full width.';
  const aligns: [string, 'l' | 'ctr' | 'r' | 'just'][] = [
    ['Left', 'l'],
    ['Center', 'ctr'],
    ['Right', 'r'],
    ['Justify', 'just'],
  ];
  const spacing =
    'Line spacing sets the distance between the lines of a paragraph: single, 1.5 lines, double, or an exact size in points.';
  const card = (x: number, y: number, w: number, h: number, name: string, body: string): string =>
    box({
      name,
      x,
      y,
      w,
      h,
      ...rounded(5),
      fill: fill('bg2'),
      line: NO_LINE,
      ins: [0.2, 0.15, 0.2, 0.15],
      text: body,
    });
  const bottom: [number, string, string, string][] = [
    [0.7, 'Single spacing', 'Single', p(r(spacing, { sz: 14 }), { lnSpc: 100 })],
    [4.75, '1.5 line spacing', 'One and a half', p(r(spacing, { sz: 14 }), { lnSpc: 150 })],
    [
      8.8,
      'Space after: 12 pt',
      'Space after',
      ['First paragraph, then 12 pt of space.', 'Second paragraph.', 'Third paragraph.']
        .map((t) => p(r(t, { sz: 14 }), { lnSpc: 100, after: 12 }))
        .join(''),
    ],
  ];
  return slide(
    () => ({
      shapes:
        header('Paragraphs', 'Alignment, line spacing, and paragraph spacing') +
        aligns
          .map(
            ([name, algn], i) =>
              label(0.7 + i * 3.03, 1.75, 2.83, name) +
              card(
                0.7 + i * 3.03,
                2.1,
                2.83,
                1.65,
                `Align ${name}`,
                p(r(text, { sz: 14 }), { algn, lnSpc: 100 })
              )
          )
          .join('') +
        bottom
          .map(
            ([x, name, shapeName, body]) =>
              label(x, 4.15, 3.84, name) + card(x, 4.5, 3.83, 1.65, shapeName, body)
          )
          .join(''),
    }),
    [
      'The four cards above hold the same paragraph with four alignments. The cards below compare line spacing and paragraph spacing.',
      'Try: click in a card and change the alignment or the line spacing in the toolbar.',
      'Try: type at the end of the justified card and watch every line stay justified.',
    ]
  );
}

// 5. Lists -------------------------------------------------------------------------------------

export function listsSlide(): FixtureSlide {
  const char = (c: string, ch: string): string =>
    `<a:buClr>${color(c)}</a:buClr><a:buFont typeface="Arial"/><a:buChar char="${ch}"/>`;
  const num = (type: string): string =>
    `<a:buClr>${color('accent1')}</a:buClr><a:buFont typeface="+mj-lt"/><a:buAutoNum type="${type}"/>`;
  const levels = (one: string, two: string, three: string): string =>
    listLevel(1, 0.375, 22, one, 14) +
    listLevel(2, 0.875, 20, two, 8) +
    listLevel(3, 1.375, 18, three, 8);
  const bullets = levels(char('accent1', '•'), char('accent1', '–'), char('tx2', '•'));
  const numbers = levels(num('arabicPeriod'), num('alphaLcParenR'), char('tx2', '•'));
  const items = (list: [number, string][]): string =>
    list.map(([lvl, t]) => p(r(t), { lvl })).join('');
  const card = (x: number, name: string, lst: string, text: string): string =>
    box({
      name,
      x,
      y: 2.15,
      w: 5.86,
      h: 3.95,
      ...rounded(4),
      fill: fill('bg2'),
      line: NO_LINE,
      ins: [0.35, 0.2, 0.35, 0.2],
      lstStyle: lst,
      text,
    });
  return slide(
    () => ({
      shapes:
        header('Lists', 'Bullets and numbering that keep their levels') +
        label(0.7, 1.75, 5.86, 'Bullets, three levels') +
        card(
          0.7,
          'Bulleted list',
          bullets,
          items([
            [0, 'Markers come from the list style'],
            [1, 'Each level has its own marker'],
            [2, 'Up to nine levels in each list style'],
            [0, 'Press Tab to go one level down'],
            [1, 'Shift+Tab goes back up'],
            [0, 'Backspace removes the bullet'],
          ])
        ) +
        label(6.77, 1.75, 5.86, 'Numbering, mixed with bullets') +
        card(
          6.77,
          'Numbered list',
          numbers,
          items([
            [0, 'Open a deck'],
            [1, 'Choose File > Open'],
            [1, 'Or drop a .pptx file on the editor'],
            [0, 'Edit the slides'],
            [1, 'Letters restart under each number'],
            [2, 'Bullets and numbers mix freely'],
            [0, 'Export a .pptx file'],
          ])
        ),
    }),
    [
      'Both lists take their markers, sizes, and indents from the shape’s list style.',
      'Try: click at the end of a list item and press Enter. A ghost bullet shows where the new item starts.',
      'Try: press Tab to move an item one level down, and Shift+Tab to move it back up.',
      'Try: press Backspace at the start of an item to remove its bullet. Press it again to join the item to the one above.',
    ]
  );
}

// 6. Shapes ------------------------------------------------------------------------------------

export function shapesSlide(): FixtureSlide {
  const columns: [string, string, string[]][] = [
    ['Basic', 'accent1', ['roundRect', 'hexagon', 'heart']],
    ['Arrows', 'accent2', ['rightArrow', 'circularArrow', 'quadArrow']],
    ['Callouts', 'accent3', ['wedgeRectCallout', 'wedgeEllipseCallout', 'cloudCallout']],
    ['Equation', 'accent4', ['mathPlus', 'mathDivide', 'mathNotEqual']],
    ['Stars', 'accent5', ['star5', 'star12', 'irregularSeal1']],
    ['Flowchart', 'accent6', ['flowChartDecision', 'flowChartDocument', 'flowChartTerminator']],
  ];
  return slide(
    () => ({
      shapes:
        header('Shapes', 'Every preset shape, drawn from the standard') +
        columns
          .map(([name, c, presets], col) => {
            const x = 0.7 + col * 2.02;
            return (
              label(x, 1.75, 1.83, name, c) +
              presets
                .map((prst, row) =>
                  box({
                    name: `${name} ${prst}`,
                    x: x + 0.19,
                    y: 2.35 + row * 1.5,
                    w: 1.45,
                    h: 1.1,
                    prst,
                    fill: fill(c),
                    line: NO_LINE,
                  })
                )
                .join('')
            );
          })
          .join(''),
    }),
    [
      'Every shape here comes from the preset geometry in the ECMA-376 standard, filled with a theme color.',
      'Try: Insert > Shape opens one submenu for each category. Pick a shape: it goes on the slide, selected, ready to move.',
      'Try: drag a shape near another one. Smart guides show when edges and centers line up.',
      'Try: select a shape and change its fill or outline color in the toolbar.',
    ]
  );
}
