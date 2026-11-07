// Synthetic decks of the interaction hang bench (`e2e/interaction-hang.bench.ts`): a deck
// open in the editor that the user then pastes into, types in, selects, and exports.

import { buildPptx, para, run, shape, table } from '../fixtures/builder.ts';
import { IN, LOREM, range, tableRows } from './deck-parts.ts';

/** `n` text slides, each with a title, a body, and speaker notes. */
export function manySlides(n: number): Uint8Array {
  return buildPptx({
    slides: range(n).map((i) => ({
      shapes:
        shape({
          id: 2,
          name: 'Title',
          x: IN,
          y: 0.5 * IN,
          cx: 11 * IN,
          cy: IN,
          text: para(run(`Slide ${i + 1}`, 'sz="3600"')),
        }) +
        shape({
          id: 3,
          name: 'Body',
          x: IN,
          y: 2 * IN,
          cx: 10 * IN,
          cy: 4 * IN,
          text: para(run(`${i}. ${LOREM}`, 'sz="1800"')),
        }),
      notes: `Notes of slide ${i + 1}. ${LOREM}`,
    })),
  });
}

/** One slide with `n` small text shapes in a grid. */
export function manyShapes(n: number): Uint8Array {
  const cols = Math.ceil(Math.sqrt(n));
  const w = Math.floor((12 * IN) / cols);
  const h = Math.floor((6.5 * IN) / cols);
  const shapes = range(n)
    .map((i) =>
      shape({
        id: i + 2,
        name: `Cell ${i}`,
        x: (i % cols) * w,
        y: Math.floor(i / cols) * h,
        cx: w,
        cy: h,
        text: para(run(`${i}`, 'sz="600"')),
        bodyPr: '<a:bodyPr wrap="square"/>',
      })
    )
    .join('');
  return buildPptx({ slides: [{ shapes }, { shapes: '' }] });
}

/** One slide with one text box of about `chars` characters in paragraphs of 500. */
export function bigText(chars: number, autofit = false): Uint8Array {
  const paras = range(Math.ceil(chars / 500))
    .map((i) => para(run(`${i} ${LOREM.repeat(6)}`.slice(0, 500), 'sz="1200"')))
    .join('');
  const bodyPr = autofit
    ? '<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>'
    : '<a:bodyPr wrap="square"/>';
  return buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          name: 'Body',
          x: IN,
          y: IN,
          cx: 10 * IN,
          cy: 5 * IN,
          text: paras,
          bodyPr,
        }),
      },
    ],
  });
}

/** One slide with a `rows`×`cols` table. */
export function bigTable(rows: number, cols: number): Uint8Array {
  return buildPptx({
    slides: [
      {
        shapes: table({
          id: 2,
          x: IN / 2,
          y: IN / 2,
          cx: 11 * IN,
          cy: 6 * IN,
          columns: range(cols).map(() => Math.floor((11 * IN) / cols)),
          tblPr: '<a:tblPr firstRow="1"/>',
          rows: tableRows(rows, cols, 'cell'),
        }),
      },
    ],
  });
}

/** One slide whose notes hold about `chars` characters. */
export function longNotes(chars: number): Uint8Array {
  const text = range(Math.ceil(chars / LOREM.length))
    .map(() => LOREM)
    .join('');
  return buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          name: 'Title',
          x: IN,
          y: IN,
          cx: 10 * IN,
          cy: IN,
          text: para(run('Notes')),
        }),
        notes: text,
      },
    ],
  });
}
