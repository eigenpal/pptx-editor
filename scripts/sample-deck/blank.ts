// The demo's blank deck for "New": one Title Slide with empty placeholders, and the layouts a
// new deck needs, in this picker order: Title Slide, Title and Content, Section Header, Two
// Content, Title Only, Blank. Every layout shows the slide number. Titles and body text use Calibri, which the editor draws with a metric-compatible
// substitute, so a new deck shows no font notice.

import {
  buildPptx,
  LAYOUT_SLIDE_NUMBER,
  para,
  run,
  type FixtureLayout,
} from '../fixtures/builder.ts';
import { LAYOUTS } from '../fixtures/slide-layouts.ts';

const IN = 914_400;
/** The master's body box (`builder.ts`, `master`), so a layout box lines up with it. */
const MARGIN = 838_200;
const WIDTH = 10_515_600;
const BODY_Y = 1_825_625;
const BODY_HEIGHT = 4_351_338;
const GAP = 0.3 * IN;
const BODY_TEXT = 'Click to edit Master text styles';

/** A layout placeholder with its own box and the layout's sample text. */
function place(o: {
  id: number;
  name: string;
  ph: string;
  x?: number;
  cx?: number;
  y: number;
  cy: number;
  text: string;
  bodyPr: string;
  lstStyle: string;
}): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${o.id}" name="${o.name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${o.ph}</p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${Math.round(o.x ?? MARGIN)}" y="${Math.round(o.y)}"/><a:ext cx="${Math.round(o.cx ?? WIDTH)}" cy="${Math.round(o.cy)}"/></a:xfrm></p:spPr>` +
    `<p:txBody>${o.bodyPr}<a:lstStyle>${o.lstStyle}</a:lstStyle>${para(run(o.text))}</p:txBody></p:sp>`
  );
}

const SECTION_HEADER: FixtureLayout = {
  name: 'Section Header',
  type: 'secHead',
  shapes:
    place({
      id: 2,
      name: 'Title 1',
      ph: '<p:ph type="title"/>',
      y: 1.71 * IN,
      cy: 2.85 * IN,
      text: 'Click to edit Master title style',
      bodyPr: '<a:bodyPr anchor="b"/>',
      lstStyle: '<a:lvl1pPr><a:defRPr sz="6000"/></a:lvl1pPr>',
    }) +
    place({
      id: 3,
      name: 'Text Placeholder 2',
      ph: '<p:ph type="body" idx="1"/>',
      y: 4.59 * IN,
      cy: 1.5 * IN,
      text: BODY_TEXT,
      bodyPr: '<a:bodyPr/>',
      lstStyle:
        '<a:lvl1pPr marL="0" indent="0"><a:buNone/><a:defRPr sz="2400"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr>',
    }),
};

/** Two equal columns inside the master's body box. */
const half = (WIDTH - GAP) / 2;
const column = (id: number, idx: number, x: number): string =>
  place({
    id,
    name: `Content Placeholder ${id - 1}`,
    ph: `<p:ph sz="half" idx="${idx}"/>`,
    x,
    cx: half,
    y: BODY_Y,
    cy: BODY_HEIGHT,
    text: BODY_TEXT,
    bodyPr: '<a:bodyPr/>',
    lstStyle: '',
  });

const TWO_CONTENT: FixtureLayout = {
  name: 'Two Content',
  type: 'twoObj',
  shapes: layout('Title Only').shapes + column(3, 1, MARGIN) + column(4, 2, MARGIN + half + GAP),
};

function layout(name: string): FixtureLayout {
  const found = LAYOUTS.find((l) => l.name === name);
  if (!found) throw new Error(`no layout ${name}`);
  return found;
}

/** An empty slide placeholder: the editor shows the prompt for its type. */
const emptyPh = (id: number, name: string, ph: string): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr/>` +
  '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>';

/** A layout with the master's slide number, as the builder's own layouts have with `footers`. */
const withSlideNumber = (l: FixtureLayout): FixtureLayout => ({
  ...l,
  shapes: l.shapes + LAYOUT_SLIDE_NUMBER,
});

export function buildBlank(): Uint8Array {
  return buildPptx({
    width: 12_192_000,
    height: 6_858_000,
    themeName: 'Simple',
    majorFont: 'Calibri',
    minorFont: 'Calibri',
    footers: true,
    layouts: [layout('Title Slide'), SECTION_HEADER, TWO_CONTENT, layout('Title Only')].map(
      withSlideNumber
    ),
    layoutOrder: [
      'Title Slide',
      'Title and Content',
      'Section Header',
      'Two Content',
      'Title Only',
      'Blank',
    ],
    slides: [
      {
        layoutName: 'Title Slide',
        shapes:
          emptyPh(2, 'Title 1', '<p:ph type="ctrTitle"/>') +
          emptyPh(3, 'Subtitle 2', '<p:ph type="subTitle" idx="1"/>'),
      },
    ],
  });
}
