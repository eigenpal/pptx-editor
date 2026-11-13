// The group effects fixture: effects on `p:grpSpPr` (ECMA-376 Part 1, 19.3.1.23) apply to the
// group's composed picture, not to each child. Each group holds an overlapping rectangle and
// ellipse, with a child space at half the group size. The first slide has outer shadows on a
// plain group, a rotated and flipped group, and a flipped group with a blurred shadow that
// keeps its direction. The second slide has a glow, a reflection, a soft edge, and an inner
// shadow on a group.

import { group, para, run, shape, solid, type FixtureSlide } from './builder.ts';

const IN = 914_400;
const BLACK_50 = '<a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr>';
const fx = (inside: string): string => `<a:effectLst>${inside}</a:effectLst>`;

/** A group at (x, y) of 3.6 × 2.4 in, with a child space of 1.8 × 1.2 in. */
function effectGroup(o: {
  id: number;
  x: number;
  y: number;
  label: string;
  effects: string;
  rot?: number;
  flipH?: boolean;
  flipV?: boolean;
}): string {
  const label = (text: string): string => para(run(text, 'sz="1400"'), '<a:pPr algn="ctr"/>');
  const bodyPr = '<a:bodyPr wrap="square" anchor="ctr"><a:noAutofit/></a:bodyPr>';
  return group(
    {
      id: o.id,
      name: `${o.label} group`,
      x: o.x,
      y: o.y,
      cx: 3.6 * IN,
      cy: 2.4 * IN,
      chX: 0,
      chY: 0,
      chCx: 1.8 * IN,
      chCy: 1.2 * IN,
      rot: o.rot,
      flipH: o.flipH,
      flipV: o.flipV,
      effects: o.effects,
    },
    shape({ id: o.id + 1, x: 0, y: 0, cx: 1.1 * IN, cy: 0.75 * IN, fill: solid('4472C4'), text: label(o.label), bodyPr }) +
      shape({
        id: o.id + 2,
        x: 0.7 * IN,
        y: 0.45 * IN,
        cx: 1.1 * IN,
        cy: 0.75 * IN,
        prst: 'ellipse',
        fill: solid('70AD47'),
        text: label(`${o.label} oval`),
        bodyPr,
      })
  );
}

const shadow = (blurRad: number, extra = ''): string =>
  fx(`<a:outerShdw blurRad="${blurRad}" dist="152400" dir="2700000"${extra}>${BLACK_50}</a:outerShdw>`);

export const groupEffectsFixtures: Record<string, FixtureSlide[]> = {
  'group-effects': [
    {
      layout: 'blank',
      shapes:
        effectGroup({ id: 2, x: 0.5 * IN, y: 1.4 * IN, label: 'Plain', effects: shadow(0) }) +
        effectGroup({ id: 5, x: 4.85 * IN, y: 1.4 * IN, label: 'Turned', effects: shadow(0), rot: 1_800_000, flipH: true }) +
        effectGroup({
          id: 8,
          x: 9.2 * IN,
          y: 1.4 * IN,
          label: 'Blurred',
          effects: shadow(101_600, ' rotWithShape="0"'),
          flipV: true,
        }),
    },
    {
      layout: 'blank',
      shapes:
        effectGroup({
          id: 2,
          x: 0.9 * IN,
          y: 0.5 * IN,
          label: 'Glow',
          effects: fx('<a:glow rad="101600"><a:srgbClr val="ED7D31"><a:alpha val="60000"/></a:srgbClr></a:glow>'),
        }) +
        effectGroup({ id: 5, x: 7.6 * IN, y: 0.5 * IN, label: 'Soft', effects: fx('<a:softEdge rad="63500"/>') }) +
        effectGroup({
          id: 8,
          x: 0.9 * IN,
          y: 3.7 * IN,
          label: 'Mirror',
          effects: fx('<a:reflection stA="50000" endPos="55000" sy="-100000" algn="bl" rotWithShape="0"/>'),
        }) +
        effectGroup({
          id: 11,
          x: 7.6 * IN,
          y: 3.7 * IN,
          label: 'Inset',
          effects: fx(`<a:innerShdw blurRad="63500" dist="50800" dir="5400000">${BLACK_50}</a:innerShdw>`),
        }),
    },
  ],
};
