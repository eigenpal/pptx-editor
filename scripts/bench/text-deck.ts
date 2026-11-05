// The synthetic decks of the editing benches (`bun run bench:edit`, `bench:edit:browser`).

import { buildPptx, para, run, shape } from '../fixtures/builder.ts';

const IN = 914_400;
const LOREM =
  'Quarterly results show steady growth across every region, with margins ahead of plan.';

/**
 * A deck of `n` text slides. Each has a title, a three-paragraph body (`Body`), and a
 * shrink-on-overflow box (`Fit`, `a:normAutofit`). The middle slide's `Fit` box overflows,
 * so a keystroke there runs the whole autofit search.
 */
export function textDeck(n: number): Uint8Array {
  const slides = Array.from({ length: n }, (_, i) => ({
    shapes:
      shape({
        id: 2,
        name: 'Title',
        x: IN,
        y: 0.5 * IN,
        cx: 11 * IN,
        cy: IN,
        text: para(run(`Slide ${i + 1} title`, 'sz="3600"')),
      }) +
      shape({
        id: 3,
        name: 'Body',
        x: IN,
        y: 2 * IN,
        cx: 6 * IN,
        cy: 3 * IN,
        text: [1, 2, 3].map((p) => para(run(`${p}. ${LOREM}`, 'sz="1800"'))).join(''),
      }) +
      shape({
        id: 4,
        name: 'Fit',
        x: 7.5 * IN,
        y: 2 * IN,
        cx: 4 * IN,
        cy: 1.5 * IN,
        bodyPr: '<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>',
        text: para(run(LOREM.repeat(i === Math.floor(n / 2) ? 6 : 1), 'sz="2000"')),
      }),
  }));
  return buildPptx({ slides });
}
