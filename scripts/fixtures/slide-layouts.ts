// The `slide-layouts` fixture: six layouts on one master (Title and Content, Blank, Title Slide,
// Two Content, Title Only, and a custom Quote layout with its own prompt text), slides on each,
// a hidden slide, two sections, empty placeholders, and speaker notes with formatted runs.
// Change layout, reset, hide, sections, and placeholder prompts are checked against it.

import { buildPptx, para, run, shape, type FixtureLayout, type FixtureOptions } from './builder.ts';

const IN = 914_400;

/** The section ids, as GUIDs (`p14:section/@id`). */
export const LAYOUT_SECTION_IDS = ['{9A3C1F52-7B2E-4D61-8E0A-3C5B7D9E1F23}', '{2D4E6F80-1A3B-4C5D-9E7F-0A1B2C3D4E56}'];

/** A layout placeholder with its own box (`xfrm`) and the layout's sample text. */
function place(o: {
  id: number;
  name: string;
  ph: string;
  x: number;
  y: number;
  cx: number;
  cy: number;
  text: string;
  lstStyle?: string;
  bodyPr?: string;
}): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${o.id}" name="${o.name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${o.ph}</p:nvPr></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${Math.round(o.x)}" y="${Math.round(o.y)}"/><a:ext cx="${Math.round(o.cx)}" cy="${Math.round(o.cy)}"/></a:xfrm></p:spPr>` +
    `<p:txBody>${o.bodyPr ?? '<a:bodyPr/>'}<a:lstStyle>${o.lstStyle ?? ''}</a:lstStyle>${para(run(o.text))}</p:txBody></p:sp>`
  );
}

/** A layout placeholder that takes its box from the master. */
const inherited = (id: number, name: string, ph: string, text: string): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr/>` +
  `<p:txBody><a:bodyPr/><a:lstStyle/>${para(run(text))}</p:txBody></p:sp>`;

const TITLE_TEXT = 'Click to edit Master title style';
const BODY_TEXT = 'Click to edit Master text styles';

export const LAYOUTS: readonly FixtureLayout[] = [
  {
    name: 'Title Slide',
    type: 'title',
    shapes:
      place({
        id: 2,
        name: 'Title 1',
        ph: '<p:ph type="ctrTitle"/>',
        x: 1.67 * IN,
        y: 1.12 * IN,
        cx: 10 * IN,
        cy: 2.39 * IN,
        text: TITLE_TEXT,
        bodyPr: '<a:bodyPr anchor="b"/>',
        lstStyle: '<a:lvl1pPr algn="ctr"><a:defRPr sz="6000"/></a:lvl1pPr>',
      }) +
      place({
        id: 3,
        name: 'Subtitle 2',
        ph: '<p:ph type="subTitle" idx="1"/>',
        x: 1.67 * IN,
        y: 3.6 * IN,
        cx: 10 * IN,
        cy: 1.66 * IN,
        text: 'Click to edit Master subtitle style',
        lstStyle: '<a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr sz="2400"/></a:lvl1pPr>',
      }),
  },
  {
    name: 'Two Content',
    type: 'twoObj',
    shapes:
      inherited(2, 'Title 1', '<p:ph type="title"/>', TITLE_TEXT) +
      place({ id: 3, name: 'Content Placeholder 2', ph: '<p:ph sz="half" idx="1"/>', x: 0.92 * IN, y: 2 * IN, cx: 5.18 * IN, cy: 4.76 * IN, text: BODY_TEXT }) +
      place({ id: 4, name: 'Content Placeholder 3', ph: '<p:ph sz="half" idx="2"/>', x: 6.25 * IN, y: 2 * IN, cx: 5.18 * IN, cy: 4.76 * IN, text: BODY_TEXT }),
  },
  { name: 'Title Only', type: 'titleOnly', shapes: inherited(2, 'Title 1', '<p:ph type="title"/>', TITLE_TEXT) },
  {
    // A custom layout: its body placeholder shows its own prompt (`hasCustomPrompt`).
    name: 'Quote',
    shapes:
      inherited(2, 'Title 1', '<p:ph type="title"/>', TITLE_TEXT) +
      place({
        id: 3,
        name: 'Quote Placeholder 2',
        ph: '<p:ph type="body" sz="quarter" idx="13" hasCustomPrompt="1"/>',
        x: 1.5 * IN,
        y: 2.2 * IN,
        cx: 9.2 * IN,
        cy: 2.6 * IN,
        text: 'Type a quote here',
        lstStyle: '<a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr sz="3200" i="1"/></a:lvl1pPr>',
      }),
  },
];

/** A slide placeholder: `ph` names its layout place, `text` is its body (empty for none). */
function slidePh(id: number, name: string, ph: string, text: string, spPr = '<p:spPr/>'): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr>${spPr}` +
    `<p:txBody><a:bodyPr/><a:lstStyle/>${text || '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>'}</p:txBody></p:sp>`
  );
}

const sections =
  '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">' +
  `<p14:section name="Opening" id="${LAYOUT_SECTION_IDS[0]}"><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="257"/></p14:sldIdLst></p14:section>` +
  `<p14:section name="Details" id="${LAYOUT_SECTION_IDS[1]}"><p14:sldIdLst><p14:sldId id="258"/><p14:sldId id="259"/><p14:sldId id="260"/><p14:sldId id="261"/></p14:sldIdLst></p14:section>` +
  '</p14:sectionLst></p:ext></p:extLst>';

export function slideLayoutsDeck(): FixtureOptions {
  return {
    layouts: LAYOUTS,
    presentationExt: sections,
    slides: [
      {
        // 1. Title Slide.
        layoutName: 'Title Slide',
        shapes:
          slidePh(2, 'Title 1', '<p:ph type="ctrTitle"/>', para(run('Quarterly review'))) +
          slidePh(3, 'Subtitle 2', '<p:ph type="subTitle" idx="1"/>', para(run('Second quarter'))),
      },
      {
        // 2. Title and Content: the body was moved and has formatted runs; the notes are formatted.
        layout: 'title',
        shapes:
          slidePh(2, 'Title 1', '<p:ph type="title"/>', para(run('Agenda'))) +
          slidePh(
            3,
            'Content Placeholder 2',
            '<p:ph idx="1"/>',
            para(run('Results ') + run('by region', 'b="1"')) + para(run('Next steps', 'sz="3600"', '<a:solidFill><a:srgbClr val="C00000"/></a:solidFill>')),
            `<p:spPr><a:xfrm><a:off x="${2 * IN}" y="${2.5 * IN}"/><a:ext cx="${6 * IN}" cy="${3 * IN}"/></a:xfrm></p:spPr>`
          ) +
          shape({ id: 4, name: 'Note box', x: 9 * IN, y: 6 * IN, cx: 2.5 * IN, cy: 0.6 * IN, text: para(run('Draft')) }),
        notesXml:
          para(run('Open with the ') + run('headline number', 'b="1"') + run(' and ') + run('pause', 'i="1"') + run('.')) +
          para(run('Then show the chart.', 'u="sng"')),
      },
      {
        // 3. Two Content.
        layoutName: 'Two Content',
        shapes:
          slidePh(2, 'Title 1', '<p:ph type="title"/>', para(run('Before and after'))) +
          slidePh(3, 'Content Placeholder 2', '<p:ph sz="half" idx="1"/>', para(run('Manual steps'))) +
          slidePh(4, 'Content Placeholder 3', '<p:ph sz="half" idx="2"/>', para(run('One click'))),
      },
      {
        // 4. Title Only, hidden: a slide show skips it.
        layoutName: 'Title Only',
        hidden: true,
        shapes:
          slidePh(2, 'Title 1', '<p:ph type="title"/>', para(run('Backup slide'))) +
          shape({ id: 3, name: 'Text box', x: IN, y: 2 * IN, cx: 6 * IN, cy: IN, text: para(run('Only if asked')) }),
      },
      {
        // 5. Title and Content with empty placeholders: the editor shows their prompts.
        layout: 'title',
        shapes: slidePh(2, 'Title 1', '<p:ph type="title"/>', '') + slidePh(3, 'Content Placeholder 2', '<p:ph idx="1"/>', ''),
      },
      {
        // 6. Quote, empty: the body shows the layout's own prompt.
        layoutName: 'Quote',
        shapes:
          slidePh(2, 'Title 1', '<p:ph type="title"/>', para(run('In their words'))) +
          slidePh(3, 'Quote Placeholder 2', '<p:ph type="body" sz="quarter" idx="13"/>', ''),
      },
    ],
  };
}

export function buildSlideLayoutsDeck(): Uint8Array {
  return buildPptx(slideLayoutsDeck());
}
