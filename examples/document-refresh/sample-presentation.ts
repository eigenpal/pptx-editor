// Controlled fixtures for the mock processor. This is not a general PPTX processor.
//
// `samplePptx(round, sequence)` returns the presentation after `sequence` of request `round`.
// Round 0 is the presentation that the page opens. Each request returns two cumulative files:
// sequence 1 changes the revenue figure and adds a summary slide, and sequence 2 also rewrites
// the notes of slide 3 and removes the draft label from slide 1.

import { buildPptx, para, run, shape, solid } from '../../scripts/fixtures/builder.ts';
import { DOCUMENT_ID, type ServerUpdate } from './src/refresh-job.ts';

const REVENUE_PREFIX = 'Revenue reached ';

/** The revenue figure after a request. Two consecutive rounds never return the same figure. */
function revenue(round: number): string {
  return round > 0 ? `$4.${(round % 9) + 1}M` : '$4.0M';
}

function notes(round: number): string {
  return round > 0
    ? `Mention the launch on March ${(round % 28) + 1}.`
    : 'Mention the launch plan.';
}

/** The draft label on the title slide, until the processor removes it. */
const DRAFT_LABEL = shape({
  id: 5,
  name: 'Draft label',
  x: 9_900_000,
  y: 5_600_000,
  cx: 1_600_000,
  cy: 600_000,
  prst: 'roundRect',
  fill: solid('FDE68A'),
  text: para(run('DRAFT', 'sz="2000" b="1"'), '<a:pPr algn="ctr"/>'),
});

function slide(title: string, body: string, extra = '', slideNotes?: string) {
  return {
    layout: 'blank' as const,
    shapes:
      shape({ id: 2, x: 0, y: 0, cx: 12_192_000, cy: 180_000, fill: solid('4F46E5') }) +
      shape({
        id: 3,
        name: 'Title',
        x: 700_000,
        y: 600_000,
        cx: 10_800_000,
        cy: 1_000_000,
        text: para(run(title, 'sz="4400" b="1"')),
      }) +
      shape({
        id: 4,
        name: 'Body',
        x: 700_000,
        y: 2_000_000,
        cx: 10_800_000,
        cy: 2_400_000,
        text: para(run(body, 'sz="2800"')),
      }) +
      extra,
    ...(slideNotes !== undefined ? { notes: slideNotes } : {}),
  };
}

/** The presentation after `sequence` (1 or 2) of request `round`. */
export function samplePptx(round = 0, sequence = 0): Uint8Array {
  const reviewed = round > 1 || sequence === 2;
  const notesRound = sequence === 2 ? round : round - 1;
  return buildPptx({
    slides: [
      slide('Quarterly review', 'Finance and product update', reviewed ? '' : DRAFT_LABEL),
      slide('Revenue', `${REVENUE_PREFIX}${revenue(round)} this quarter.`),
      slide('Roadmap', 'Launch the mobile editor.', '', notes(notesRound)),
      ...(round > 0 ? [slide('Summary', 'Revenue grew. The launch stays on schedule.')] : []),
    ],
  });
}

/** One processor result: the cumulative file, its change locations, and failed operations. */
export function sampleUpdate(submissionId: string, round: number, sequence: number): ServerUpdate {
  const reviewed = round > 1 || sequence === 2;
  const figure = revenue(round);
  return {
    documentId: DOCUMENT_ID,
    submissionId,
    sequence,
    bytes: Buffer.from(samplePptx(round, sequence)).toString('base64'),
    changes: [
      {
        id: 'revenue',
        description: `Changed the revenue figure to ${figure}`,
        location: {
          slideIndex: 1,
          shapeId: 4,
          paragraphIndex: 0,
          start: REVENUE_PREFIX.length,
          end: REVENUE_PREFIX.length + figure.length,
          text: figure,
        },
      },
      {
        id: 'summary-slide',
        description: 'Added a summary slide',
        location: { slideIndex: 3 },
      },
      ...(reviewed
        ? [
            {
              id: 'launch-notes',
              description: 'Rewrote the speaker notes of slide 3',
              location: { slideIndex: 2, notes: true },
            },
            {
              id: 'draft-label',
              description: 'Removed the draft label from slide 1',
              unavailableReason: 'deleted' as const,
            },
          ]
        : []),
    ],
    // Simulate a failed processor operation without repeating successful work.
    failures: sequence === 2 ? ['regional-chart'] : [],
  };
}
