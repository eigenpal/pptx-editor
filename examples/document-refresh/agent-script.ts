// A scripted "agent" for the replay page. Each step is one edit that an AI agent could make on a
// server. `agentPptx(step)` returns the complete presentation after that many steps, so every
// file is cumulative, as a real processor returns it. This is not a general PPTX processor.

import {
  buildPptx,
  para,
  run,
  shape,
  solid,
  type FixtureSlide,
} from '../../scripts/fixtures/builder.ts';
import { AGENT_STEPS } from './src/agent-steps.ts';

export { AGENT_STEPS };

const TITLE_BAR = shape({ id: 2, x: 0, y: 0, cx: 12_192_000, cy: 180_000, fill: solid('4F46E5') });

function text(id: number, name: string, y: number, cy: number, words: string, size: number) {
  return shape({
    id,
    name,
    x: 700_000,
    y,
    cx: 10_800_000,
    cy,
    text: para(run(words, `sz="${size}"${size >= 4000 ? ' b="1"' : ''}`)),
  });
}

function slide(title: string, body: string, extra = '', notes?: string): FixtureSlide {
  return {
    layout: 'blank',
    shapes:
      TITLE_BAR +
      text(3, 'Title', 600_000, 1_000_000, title, 4400) +
      text(4, 'Body', 2_000_000, 1_600_000, body, 2800) +
      extra,
    ...(notes !== undefined ? { notes } : {}),
  };
}

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

function growthBox(highlighted: boolean) {
  return shape({
    id: 5,
    name: 'Growth',
    x: 700_000,
    y: 4_000_000,
    cx: 3_600_000,
    cy: 1_200_000,
    prst: 'roundRect',
    fill: solid(highlighted ? '22C55E' : 'CBD5E1'),
    text: para(
      run(highlighted ? 'Growth: 15%' : 'Growth: 8%', 'sz="3200" b="1"'),
      '<a:pPr algn="ctr"/>'
    ),
  });
}

/** The presentation after `step` agent steps (0 is the presentation the page opens). */
export function agentPptx(step: number): Uint8Array {
  const done = (n: number) => step >= n;
  const title = slide(
    done(3) ? 'Q3 business review' : 'Quarterly review',
    'Finance and product update',
    done(6) ? '' : DRAFT_LABEL
  );
  const revenue = slide(
    'Revenue',
    `Revenue reached ${done(1) ? '$4.6M' : '$4.0M'} this quarter.`,
    growthBox(done(2))
  );
  const roadmap = slide(
    'Roadmap',
    'Launch the mobile editor.',
    '',
    done(5)
      ? 'Open with the beta numbers. Then announce the March launch date.'
      : 'Mention the launch plan.'
  );
  const team = slide('Team', 'Hire two engineers and one designer.');
  const risks = slide('Risks', 'Supplier delays could move the launch by two weeks.');
  const slides = done(7)
    ? [title, revenue, team, roadmap, risks]
    : [title, revenue, roadmap, ...(done(4) ? [risks] : []), team];
  return buildPptx({ slides });
}
