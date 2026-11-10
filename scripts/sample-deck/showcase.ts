// Build the demo's showcase deck: one slide per feature that PPTX Editor renders and edits.
// Output: examples/vite/public/showcase.pptx.
//
//   bun scripts/sample-deck/showcase.ts      (also run by `bun run sample:build`)
//
// The showcase test records known approximation diagnostics by slide. These checks do not
// replace PDF fidelity comparisons. Every slide's speaker notes say what to try.
// All content is written here; the only picture is generated (showcase-image.ts).

import { mkdirSync, writeFileSync } from 'node:fs';
import { buildPptx, type FixtureSlide } from '../fixtures/builder.ts';
import { landscapePng } from './showcase-image.ts';
import { specializedShowcase } from './showcase-specialized.ts';
import {
  NO_FILL,
  NO_LINE,
  R_LINK,
  SHOWCASE_COLORS,
  box,
  cxn,
  fill,
  group,
  header,
  label,
  line,
  p,
  r,
  rounded,
  DARK_SLIDE,
  FAINT_TEXT,
  LIGHT_TEXT,
  slide,
} from './showcase-kit.ts';
import {
  charactersSlide,
  listsSlide,
  paragraphsSlide,
  pipelineSlide,
  shapesSlide,
  titleSlide,
} from './showcase-slides.ts';
import {
  TABLE_STYLES,
  effectsSlide,
  fillsSlide,
  linesSlide,
  picturesSlide,
  tablesSlide,
} from './showcase-visuals.ts';

const OUT = 'examples/vite/public';
const SITE = 'https://pptx-editor.dev';

// 12. Groups and transforms --------------------------------------------------------------------

function arrangeSlide(): FixtureSlide {
  const caption = (x: number, text: string): string =>
    box({
      name: 'Caption',
      x,
      y: 5.75,
      w: 3.63,
      h: 0.8,
      ins: [0, 0, 0, 0],
      text: p(r(text, { sz: 13, color: 'tx2' })),
    });
  const flips: [string, boolean, boolean][] = [
    ['Original', false, false],
    ['Flip horizontal', true, false],
    ['Flip vertical', false, true],
  ];
  return slide(
    () => ({
      shapes:
        header('Arrange', 'Groups, rotation, and flips') +
        label(0.7, 1.85, 3.63, 'Group') +
        group({ name: 'Badge group', x: 1.3, y: 2.55, w: 2.45, h: 2.75 }, () =>
          [
            box({
              name: 'Badge',
              x: 1.6,
              y: 2.75,
              w: 1.85,
              h: 1.85,
              prst: 'ellipse',
              fill: fill('accent2'),
              line: NO_LINE,
            }),
            box({
              name: 'Badge star',
              x: 2.95,
              y: 2.55,
              w: 0.8,
              h: 0.8,
              prst: 'star5',
              fill: fill('accent3'),
              line: NO_LINE,
            }),
            box({
              name: 'Badge label',
              x: 1.3,
              y: 4.5,
              w: 2.45,
              h: 0.8,
              ...rounded(50),
              fill: fill('tx1'),
              line: NO_LINE,
              anchor: 'ctr',
              text: p(r('Three shapes', { sz: 16, b: true, color: 'bg1' }), { algn: 'ctr' }),
            }),
          ].join('')
        ) +
        caption(
          0.7,
          'Double-click the group to select one shape in it. Ctrl+G groups; Ctrl+Shift+G ungroups.'
        ) +
        label(4.85, 1.85, 3.63, 'Rotation') +
        [15, 30, 45]
          .map((deg, i) =>
            box({
              name: `Rotated ${deg}`,
              x: 5.05 + i * 1.2,
              y: 3.2,
              w: 0.8,
              h: 1.1,
              rot: deg,
              ...rounded(14),
              fill: fill('accent1'),
              line: NO_LINE,
              anchor: 'ctr',
              ins: [0, 0, 0, 0],
              text: p(r(`${deg}°`, { sz: 16, b: true, color: 'bg1' }), { algn: 'ctr' }),
            })
          )
          .join('') +
        caption(
          4.85,
          'Text turns with its shape. Drag the round handle above a shape to rotate it.'
        ) +
        label(9.0, 1.85, 3.63, 'Flip') +
        flips
          .map(([name, flipH, flipV], i) => {
            const x = 9.0 + i * 1.21;
            return (
              box({
                name,
                x: x + 0.13,
                y: 3.2,
                w: 0.95,
                h: 0.95,
                prst: 'rtTriangle',
                flipH,
                flipV,
                fill: fill('accent4'),
                line: NO_LINE,
              }) +
              box({
                name: `Label ${name}`,
                x: x - 0.05,
                y: 4.3,
                w: 1.31,
                h: 0.5,
                ins: [0, 0, 0, 0],
                text: p(r(name, { sz: 12, b: true }), { algn: 'ctr' }),
              })
            );
          })
          .join('') +
        caption(
          9.0,
          'A flip is part of the shape’s transform, and it is kept exactly when you save.'
        ),
    }),
    [
      'Groups, rotation, and flips are transforms in the file. The editor keeps them exact when you move and save.',
      'Try: click the badge to select the group, then double-click to select one shape inside it.',
      'Try: hold Shift and click two shapes to select both, then press Ctrl+G (Cmd+G on a Mac) to group them.',
      'Try: drag the round handle above an orange card to rotate it.',
    ]
  );
}

// 13. Text boxes -------------------------------------------------------------------------------

function textBoxesSlide(): FixtureSlide {
  const columns = [
    'Text flows from the bottom of the first column to the top of the second.',
    'The number of columns and the gap between them belong to the text box, so the words move from one column to the next as you type.',
    'Columns suit notes, quotes, and short lists that would make a wide box hard to read.',
  ];
  const vertical = (x: number, vert: 'vert' | 'vert270', text: string, c: string, fg: string) =>
    box({
      name: `Vertical ${vert}`,
      x,
      y: 2.25,
      w: 0.95,
      h: 3.85,
      ...rounded(10),
      fill: fill(c),
      line: NO_LINE,
      vert,
      anchor: 'ctr',
      text: p(r(text, { sz: 20, b: true, color: fg }), { algn: 'ctr' }),
    });
  return slide(
    () => ({
      shapes:
        header('Text boxes', 'Vertical text, columns, and autofit') +
        label(0.7, 1.85, 2.1, 'Vertical') +
        vertical(0.7, 'vert270', 'Rotated 270°', 'accent1', 'bg1') +
        vertical(1.85, 'vert', 'Rotated 90°', 'bg2', 'tx1') +
        label(3.3, 1.85, 9.33, 'Two columns') +
        box({
          name: 'Two columns',
          x: 3.3,
          y: 2.25,
          w: 9.33,
          h: 1.55,
          ...rounded(5),
          fill: fill('bg2'),
          line: NO_LINE,
          cols: [2, 0.4],
          ins: [0.3, 0.2, 0.3, 0.2],
          text: columns.map((t) => p(r(t, { sz: 15 }), { lnSpc: 100, after: 8 })).join(''),
        }) +
        label(3.3, 4.15, 4.5, 'Shape grows to fit') +
        box({
          name: 'Shape autofit',
          x: 3.3,
          y: 4.5,
          w: 4.5,
          h: 0.8333,
          ...rounded(12),
          fill: fill('accent2', 16),
          line: line(1, 'accent2'),
          fit: 'shape',
          ins: [0.2, 0.15, 0.2, 0.15],
          text:
            p(r('This box grows as you type.', { sz: 16 }), { lnSpc: 100 }) +
            p(r('Press Enter to add a line.', { sz: 16 }), { lnSpc: 100 }),
        }) +
        label(8.13, 4.15, 4.5, 'Text shrinks to fit') +
        box({
          name: 'Shrink on overflow',
          x: 8.13,
          y: 4.5,
          w: 4.5,
          h: 1.1,
          ...rounded(8),
          fill: fill('accent4', 12),
          line: line(1, 'accent4'),
          fit: 'shrink 85',
          ins: [0.2, 0.12, 0.2, 0.12],
          text: p(
            r(
              'This box keeps its size. When the text is too long, the editor draws it smaller to fit.',
              { sz: 18 }
            ),
            { lnSpc: 100 }
          ),
        }) +
        box({
          name: 'Autofit note',
          x: 3.3,
          y: 5.8,
          w: 9.33,
          h: 0.3,
          ins: [0, 0, 0, 0],
          anchor: 'b',
          text: p(
            r('Both settings are stored in the file: shape autofit and shrink text on overflow.', {
              sz: 12,
              color: 'tx2',
            })
          ),
        }),
    }),
    [
      'Text boxes can turn their text, flow it in columns, and fit text and shape to each other.',
      'Try: click at the end of the green box and press Enter. The box grows with the new line.',
      'Try: type more text in the blue box. The text gets smaller so that it stays inside.',
      'Try: type in the two-column box and watch the words move to the second column.',
    ]
  );
}

// 14. Links ------------------------------------------------------------------------------------

function linksSlide(): FixtureSlide {
  const points = [
    'A link opens only when you click it, never when the deck opens.',
    'Only web, email, phone, and FTP links are allowed.',
    'Ctrl+K adds, changes, or removes a link.',
  ];
  return slide(
    () => ({
      rels: [['rIdLink', R_LINK, SITE, 'External']],
      shapes:
        header('Links', 'Hyperlinks you can follow and edit') +
        box({
          name: 'Link sentence',
          x: 0.7,
          y: 2.1,
          w: 7.4,
          h: 1.45,
          ins: [0, 0.05, 0, 0.05],
          anchor: 't',
          text:
            p(r('Read the documentation at', { sz: 32, b: true }), { lnSpc: 100 }) +
            p(r('pptx-editor.dev', { sz: 32, b: true, link: 'rIdLink' }), { lnSpc: 100 }),
        }) +
        box({
          name: 'Link points',
          x: 0.7,
          y: 3.95,
          w: 7.4,
          h: 2.4,
          ins: [0, 0, 0, 0],
          text: points
            .map((t) =>
              p(r(t, { sz: 18 }), {
                marL: 0.3,
                indent: -0.3,
                after: 12,
                lnSpc: 100,
                bullet:
                  '<a:buClr><a:schemeClr val="accent1"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="•"/>',
              })
            )
            .join(''),
        }) +
        box({
          name: 'Link card',
          x: 8.85,
          y: 2.1,
          w: 3.78,
          h: 3.6,
          ...rounded(6),
          fill: fill('bg2'),
          line: NO_LINE,
        }) +
        // A chain link drawn from two rounded outlines, turned 45 degrees.
        box({
          name: 'Chain 1',
          x: 9.55,
          y: 2.85,
          w: 1.6,
          h: 0.75,
          rot: -45,
          ...rounded(50),
          fill: NO_FILL,
          line: line(9, 'accent4'),
        }) +
        box({
          name: 'Chain 2',
          x: 10.35,
          y: 2.85,
          w: 1.6,
          h: 0.75,
          rot: -45,
          ...rounded(50),
          fill: NO_FILL,
          line: line(9, 'accent1'),
        }) +
        box({
          name: 'Link address',
          x: 8.85,
          y: 4.6,
          w: 3.78,
          h: 0.5,
          ins: [0, 0, 0, 0],
          anchor: 'ctr',
          text: p(r(SITE, { sz: 14, color: 'tx2' }), { algn: 'ctr' }),
        }),
    }),
    [
      'The link points to https://pptx-editor.dev. It is stored as an external relationship and never fetched when the deck opens.',
      'Try: present the deck (F5) and click the link to open it in a new tab.',
      'Try: select a word, press Ctrl+K (Cmd+K on a Mac), paste a link, and press Enter.',
    ]
  );
}

// 15. Editing features -------------------------------------------------------------------------

function editingSlide(): FixtureSlide {
  const tiles: [string, string, string][] = [
    ['Select slides', 'Pick several slides in the rail', 'Shift+click, Ctrl+click'],
    ['Right-click menus', 'Actions for shapes, text, and slides', 'Right-click'],
    ['Format painter', 'Copy formatting to other shapes', 'Ctrl+Shift+C, Ctrl+Shift+V'],
    ['Find and replace', 'Search the slides and the notes', 'Ctrl+F, Ctrl+H'],
    ['Undo and redo', 'Every edit can be undone', 'Ctrl+Z, Ctrl+Y'],
    ['Group', 'Move several shapes as one', 'Ctrl+G, Ctrl+Shift+G'],
    ['Slide numbers', 'Number every slide in one step', 'Insert > Slide numbers'],
    ['Export pictures', 'A slide as PNG or SVG, or all slides', 'File > Current slide as PNG'],
    ['Present', 'Show the deck full screen', 'F5, Shift+F5'],
  ];
  return slide(
    () => ({
      shapes:
        header('Editing', 'Edit slides directly on the canvas') +
        tiles
          .map(([title, desc, keys], i) =>
            box({
              name: `Feature ${title}`,
              x: 0.7 + (i % 3) * 4.05,
              y: 1.85 + Math.floor(i / 3) * 1.55,
              w: 3.83,
              h: 1.38,
              ...rounded(7),
              fill: fill('bg2'),
              line: NO_LINE,
              anchor: 'ctr',
              ins: [0.25, 0.12, 0.25, 0.12],
              text:
                p(r(title, { sz: 17, b: true })) +
                p(r(desc, { sz: 13, color: 'tx2' }), { before: 2 }) +
                p(r(keys, { sz: 13, b: true, color: 'accent1' }), { before: 6 }),
            })
          )
          .join('') +
        box({
          name: 'Mac note',
          x: 0.7,
          y: 6.5,
          w: 8,
          h: 0.32,
          ins: [0, 0, 0, 0],
          text: p(r('On a Mac, press Cmd where this slide says Ctrl.', { sz: 12, color: 'tx2' })),
        }),
    }),
    [
      'Each tile names a feature and how to reach it.',
      'Try: Shift+click slides 2 to 4 in the rail, then right-click them to duplicate or delete all three.',
      'Try: select a tile, press Ctrl+Shift+C, select another shape, and press Ctrl+Shift+V.',
      'Try: press Ctrl+F and search for "slide". Matches in the speaker notes count too.',
      'Try: Insert > Slide numbers, then File > Current slide as PNG.',
    ]
  );
}

// 16. Hidden slide -----------------------------------------------------------------------------

function hiddenSlide(): FixtureSlide {
  return slide(
    () => ({
      hidden: true,
      shapes:
        header(
          'Hidden slide',
          'The slide show skips this slide',
          'It is hidden: the rail marks it, and the show goes from slide 15 straight to slide 17.'
        ) +
        box({
          name: 'Eye',
          x: 5.17,
          y: 2.75,
          w: 3.0,
          h: 1.7,
          prst: 'ellipse',
          fill: NO_FILL,
          line: line(7, 'tx1'),
        }) +
        box({
          name: 'Pupil',
          x: 6.22,
          y: 3.15,
          w: 0.9,
          h: 0.9,
          prst: 'ellipse',
          fill: fill('accent1'),
          line: NO_LINE,
        }) +
        cxn({
          name: 'Slash',
          x: 5.37,
          y: 2.35,
          w: 2.6,
          h: 2.5,
          flipH: true,
          line: line(7, 'tx1', '', ' cap="rnd"'),
        }) +
        box({
          name: 'Hidden note',
          x: 2.67,
          y: 5.25,
          w: 8.0,
          h: 0.5,
          ins: [0, 0, 0, 0],
          anchor: 'ctr',
          text: p(
            r('During the show, type 16 and press Enter to see it anyway.', {
              sz: 18,
              color: 'tx2',
            }),
            { algn: 'ctr' }
          ),
        }),
    }),
    [
      'This slide is hidden (show="0" in the file). The slide show skips it.',
      'Try: press F5 and step through the deck. The show goes from slide 15 to slide 17.',
      'Try: during the show, type 16 and press Enter to show this slide anyway.',
    ]
  );
}

// 17. Closing ----------------------------------------------------------------------------------

function closingSlide(): FixtureSlide {
  return slide(
    () => ({
      ...DARK_SLIDE,
      shapes:
        box({
          name: 'Circle orange',
          x: 9.15,
          y: 1.75,
          w: 2.9,
          h: 2.9,
          prst: 'ellipse',
          fill: fill('accent1', 90),
          line: NO_LINE,
        }) +
        box({
          name: 'Circle green',
          x: 10.45,
          y: 3.35,
          w: 2.2,
          h: 2.2,
          prst: 'ellipse',
          fill: fill('accent2', 85),
          line: NO_LINE,
        }) +
        box({
          name: 'Kicker',
          x: 0.8,
          y: 1.95,
          w: 7.5,
          h: 0.36,
          ins: [0, 0, 0, 0],
          anchor: 'b',
          text: p(
            r('Try it with your own deck', {
              sz: 14,
              b: true,
              caps: true,
              spc: 3,
              color: 'accent2',
            })
          ),
        }) +
        box({
          name: 'Title',
          ph: '<p:ph type="title"/>',
          x: 0.8,
          y: 2.35,
          w: 8.2,
          h: 1.6,
          ins: [0, 0.05, 0, 0.05],
          anchor: 't',
          text:
            p(r('Open a .pptx file', { sz: 48, b: true, color: 'tx1' })) +
            p(r('and start editing', { sz: 48, b: true, color: 'tx1' })),
        }) +
        box({
          name: 'Body',
          x: 0.8,
          y: 4.0,
          w: 8.0,
          h: 1.1,
          ins: [0, 0.05, 0, 0.05],
          text: p(
            r(
              'Drop a file on the editor or choose File > Open. Your deck stays in your browser, and Export saves a standard .pptx file.',
              { sz: 20, color: LIGHT_TEXT }
            ),
            { lnSpc: 100 }
          ),
        }) +
        box({
          name: 'Accent bar',
          x: 0.8,
          y: 5.35,
          w: 1.1,
          h: 0.07,
          fill: fill('accent1'),
          line: NO_LINE,
        }) +
        box({
          name: 'Site',
          x: 0.8,
          y: 5.55,
          w: 7,
          h: 0.45,
          ins: [0, 0.05, 0, 0.05],
          text: p(r('pptx-editor.dev', { sz: 20, b: true, color: 'tx1' })),
        }) +
        box({
          name: 'Family',
          x: 0.8,
          y: 6.0,
          w: 8,
          h: 0.4,
          ins: [0, 0.05, 0, 0.05],
          text: p(
            r('An EigenPal editor family product, with DOCX Editor', { sz: 14, color: FAINT_TEXT })
          ),
        }),
    }),
    [
      'Thank you. Open your own deck to see how it looks and to edit it.',
      'Try: File > Open, or drag a .pptx file onto the editor. Nothing is uploaded.',
      'Try: Export, then the .pptx item, to download this deck with your edits.',
    ]
  );
}

/** The showcase deck's slides, in order. */
export function showcaseSlides(): FixtureSlide[] {
  return [
    titleSlide(),
    pipelineSlide(),
    charactersSlide(),
    paragraphsSlide(),
    listsSlide(),
    shapesSlide(),
    fillsSlide(),
    linesSlide(),
    effectsSlide(),
    picturesSlide(),
    tablesSlide(),
    arrangeSlide(),
    textBoxesSlide(),
    linksSlide(),
    editingSlide(),
    hiddenSlide(),
    closingSlide(),
    ...specializedShowcase().slides,
  ];
}

/** The showcase deck as `.pptx` bytes. */
export function buildShowcase(): Uint8Array {
  return buildPptx({
    width: 12_192_000,
    height: 6_858_000,
    slides: showcaseSlides(),
    extraParts: specializedShowcase().extraParts,
    binaryParts: specializedShowcase().binaryParts,
    media: [{ name: 'image1.png', bytes: landscapePng(), contentType: 'image/png' }],
    colors: SHOWCASE_COLORS,
    majorFont: 'Calibri',
    minorFont: 'Calibri',
    tableStyles: TABLE_STYLES,
    footers: true,
  });
}

if (import.meta.main) {
  mkdirSync(OUT, { recursive: true });
  writeFileSync(`${OUT}/showcase.pptx`, buildShowcase());
  console.log(`Wrote ${OUT}/showcase.pptx`);
}
