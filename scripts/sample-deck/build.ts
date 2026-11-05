// Build the demo's own sample deck (authored here, like DOCX Editor's sample.docx), a blank
// deck for "New" (blank.ts), and the feature showcase (showcase.ts). Output:
// examples/vite/public/{sample,blank,showcase}.pptx.
//
//   bun scripts/sample-deck/build.ts

import { mkdirSync, writeFileSync } from 'node:fs';
import { buildPptx, para, run, shape, solid } from '../fixtures/builder.ts';
import { buildBlank } from './blank.ts';
import { buildShowcase } from './showcase.ts';

const IN = 914_400;
const OUT = 'examples/vite/public';
const INK = '1B1A19';
const ORANGE = 'C43E1C';
const GREEN = '2EB52E';
const MUTED = '75716B';

// A one-line 36 pt title at 90 % line spacing needs 0.64 in (1.2 × 36 pt × 0.9 plus the
// 0.05 in top and bottom insets). The box fits it, so the text sits evenly inside its frame.
const title = (text: string, y = 0.45 * IN): string =>
  shape({
    id: 2,
    x: 0.7 * IN,
    y: y + 0.26 * IN,
    cx: 11.9 * IN,
    cy: 0.64 * IN,
    bodyPr: '<a:bodyPr wrap="square" lIns="0" rIns="0" anchor="b"><a:normAutofit/></a:bodyPr>',
    text: para(run(text, `sz="3600" b="1"`, solid(INK))),
  });

const kicker = (text: string): string =>
  shape({
    id: 3,
    x: 0.7 * IN,
    y: 0.3 * IN,
    cx: 6 * IN,
    cy: 0.35 * IN,
    bodyPr: '<a:bodyPr wrap="square" lIns="0" rIns="0" tIns="0" bIns="0"><a:noAutofit/></a:bodyPr>',
    text: para(run(text, 'sz="1200" b="1" spc="200" cap="all"', solid(ORANGE))),
  });

const bullet = (text: string, level = 0): string =>
  para(
    run(text, 'sz="2000"', solid(INK)),
    `<a:pPr marL="${342900 + level * 342900}" indent="-342900"><a:spcBef><a:spcPts val="900"/></a:spcBef><a:buClr><a:srgbClr val="${ORANGE}"/></a:buClr><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>`
  );

const pipelineStep = (id: number, i: number, label: string, sub: string): string =>
  shape({
    id,
    x: (0.7 + i * 2.45) * IN,
    y: 2.6 * IN,
    cx: 2.05 * IN,
    cy: 1.5 * IN,
    prst: 'roundRect',
    avLst: '<a:gd name="adj" fmla="val 12000"/>',
    fill: solid(i === 2 ? 'FBE3D6' : 'F4F3F0'),
    line: `<a:ln w="12700"><a:solidFill><a:srgbClr val="${i === 2 ? ORANGE : 'D9D5CE'}"/></a:solidFill></a:ln>`,
    bodyPr: '<a:bodyPr wrap="square" anchor="ctr"><a:noAutofit/></a:bodyPr>',
    text:
      para(run(label, 'sz="1600" b="1"', solid(INK)), '<a:pPr algn="ctr"/>') +
      para(run(sub, 'sz="1100"', solid(MUTED)), '<a:pPr algn="ctr"/>'),
  });

const arrow = (id: number, i: number): string =>
  `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="Arrow ${i}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="${Math.round((2.75 + i * 2.45) * IN)}" y="${Math.round(3.35 * IN)}"/><a:ext cx="${Math.round(0.4 * IN)}" cy="0"/></a:xfrm>` +
  `<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln w="19050"><a:solidFill><a:srgbClr val="${MUTED}"/></a:solidFill><a:tailEnd type="triangle"/></a:ln></p:spPr></p:cxnSp>`;

const slides = [
  {
    background:
      '<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="F4F3F0"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>',
    shapes:
      shape({
        id: 4,
        x: 0,
        y: 6.9 * IN,
        cx: 13.333 * IN,
        cy: 0.6 * IN,
        fill: `<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="${ORANGE}"/></a:gs><a:gs pos="100000"><a:srgbClr val="F2A07C"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>`,
      }) +
      shape({
        id: 5,
        x: 0.7 * IN,
        y: 1.6 * IN,
        cx: 0.9 * IN,
        cy: 0.9 * IN,
        prst: 'ellipse',
        fill: solid(GREEN),
        effects:
          '<a:effectLst><a:outerShdw blurRad="76200" dist="25400" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="25000"/></a:srgbClr></a:outerShdw></a:effectLst>',
      }) +
      shape({
        id: 6,
        x: 0.7 * IN,
        y: 2.7 * IN,
        cx: 11 * IN,
        cy: 1.2 * IN,
        bodyPr: '<a:bodyPr wrap="square" lIns="0" rIns="0"><a:normAutofit/></a:bodyPr>',
        text: para(run('PPTX Editor', 'sz="6000" b="1"', solid(INK))),
      }) +
      shape({
        id: 7,
        x: 0.7 * IN,
        y: 3.95 * IN,
        cx: 10 * IN,
        cy: 0.8 * IN,
        bodyPr: '<a:bodyPr wrap="square" lIns="0" rIns="0"><a:noAutofit/></a:bodyPr>',
        text: para(
          run('Faithful PPTX slides that you can edit in the browser', 'sz="2400"', solid(MUTED))
        ),
      }) +
      shape({
        id: 8,
        x: 0.7 * IN,
        y: 5.6 * IN,
        cx: 6 * IN,
        cy: 0.5 * IN,
        bodyPr: '<a:bodyPr wrap="square" lIns="0" rIns="0"><a:noAutofit/></a:bodyPr>',
        text: para(run('An EigenPal editor family product', 'sz="1400" b="1"', solid(GREEN))),
      }),
    notes: 'Welcome. Everything in this deck is editable: double-click any text to type.',
  },
  {
    shapes:
      kicker('What it does') +
      title('One engine, faithful to the file') +
      shape({
        id: 4,
        x: 0.7 * IN,
        y: 1.7 * IN,
        cx: 7.2 * IN,
        cy: 4.6 * IN,
        bodyPr: '<a:bodyPr wrap="square" lIns="0"><a:normAutofit/></a:bodyPr>',
        text:
          bullet('Opens any .pptx in the browser; nothing is uploaded') +
          bullet('Renders themes, placeholders, geometry, and effects from the spec') +
          bullet('Measures text with HarfBuzz and the deck’s own fonts') +
          bullet('Embedded fonts are decoded in memory', 1) +
          bullet('Edits are store operations: undoable and lossless') +
          bullet('Saves a valid deck that opens without a repair prompt'),
      }) +
      shape({
        id: 5,
        x: 8.4 * IN,
        y: 1.9 * IN,
        cx: 4.2 * IN,
        cy: 4.2 * IN,
        prst: 'roundRect',
        avLst: '<a:gd name="adj" fmla="val 8000"/>',
        fill: solid('FBE3D6'),
        bodyPr:
          '<a:bodyPr wrap="square" lIns="274320" rIns="274320" anchor="ctr"><a:noAutofit/></a:bodyPr>',
        text:
          para(run('114', 'sz="6000" b="1"', solid(ORANGE)), '<a:pPr algn="ctr"/>') +
          para(
            run('real slides in the fidelity corpus', 'sz="1600"', solid(INK)),
            '<a:pPr algn="ctr"/>'
          ),
      }),
    notes: 'The fidelity harness compares every slide with a reference PDF of the same deck.',
  },
  {
    shapes:
      kicker('Architecture') +
      title('The same pipeline as DOCX Editor') +
      ['Bytes', 'Store', 'Cascade', 'Layout', 'Paint']
        .map((label, i) =>
          pipelineStep(
            10 + i,
            i,
            label,
            [
              'bounded OPC read',
              'the only truth',
              'master → slide',
              'HarfBuzz text',
              'SVG surface',
            ][i]
          )
        )
        .join('') +
      [0, 1, 2, 3].map((i) => arrow(20 + i, i)).join('') +
      shape({
        id: 30,
        x: 0.7 * IN,
        y: 4.7 * IN,
        cx: 11.9 * IN,
        cy: 1.4 * IN,
        bodyPr:
          '<a:bodyPr wrap="square" lIns="0" numCol="2" spcCol="457200"><a:noAutofit/></a:bodyPr>',
        text:
          para(
            run(
              'Every edit is a transaction of serializable operations on the document store. Layout and paint read resolved values only, so the editor and the fidelity harness share one renderer.',
              'sz="1400"',
              solid(MUTED)
            )
          ) +
          para(
            run(
              'Unknown content survives save untouched. A deck that opens, saves, and reopens keeps an identical canonical fingerprint.',
              'sz="1400"',
              solid(MUTED)
            )
          ),
      }),
    notes: '',
  },
  {
    shapes:
      kicker('Shapes') +
      title('Preset geometry from the spec') +
      ['rect', 'roundRect', 'ellipse', 'hexagon', 'rightArrow', 'chevron', 'star5', 'cloud']
        .map((prst, i) =>
          shape({
            id: 10 + i,
            x: (0.9 + (i % 4) * 3.05) * IN,
            y: (1.8 + Math.floor(i / 4) * 2.4) * IN,
            cx: 2.3 * IN,
            cy: 1.9 * IN,
            prst,
            rot: i === 4 ? 10 * 60000 : 0,
            fill: `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="${i % 2 ? 'F2A07C' : '9AD49A'}"/></a:gs><a:gs pos="100000"><a:srgbClr val="${i % 2 ? ORANGE : '1E7A34'}"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill>`,
            line: '<a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>',
          })
        )
        .join(''),
    notes: 'Each shape is drawn from the ECMA-376 guide formulas, never by hand.',
  },
  {
    shapes:
      kicker('Try it') +
      title('Edit this slide') +
      shape({
        id: 4,
        x: 0.7 * IN,
        y: 1.8 * IN,
        cx: 6 * IN,
        cy: 3.6 * IN,
        bodyPr: '<a:bodyPr wrap="square" lIns="0"><a:noAutofit/></a:bodyPr>',
        text:
          para(
            run('1. Click a shape and drag it, or pull a handle.', 'sz="1800"', solid(INK)),
            '<a:pPr><a:spcAft><a:spcPts val="1200"/></a:spcAft></a:pPr>'
          ) +
          para(
            run('2. Double-click text and type.', 'sz="1800"', solid(INK)),
            '<a:pPr><a:spcAft><a:spcPts val="1200"/></a:spcAft></a:pPr>'
          ) +
          para(
            run('3. Use the toolbar for ', 'sz="1800"', solid(INK)) +
              run('bold', 'sz="1800" b="1"', solid(INK)) +
              run(', ', 'sz="1800"', solid(INK)) +
              run('color', 'sz="1800"', solid(ORANGE)) +
              run(', and size.', 'sz="1800"', solid(INK)),
            '<a:pPr><a:spcAft><a:spcPts val="1200"/></a:spcAft></a:pPr>'
          ) +
          para(
            run('4. Export the deck and open it in your presentation app.', 'sz="1800"', solid(INK))
          ),
      }) +
      shape({
        id: 5,
        x: 7.6 * IN,
        y: 2.2 * IN,
        cx: 4.9 * IN,
        cy: 1.15 * IN,
        prst: 'wedgeRoundRectCallout',
        fill: solid('F4F3F0'),
        line: '<a:ln w="12700"><a:solidFill><a:srgbClr val="D9D5CE"/></a:solidFill></a:ln>',
        bodyPr:
          '<a:bodyPr wrap="square" lIns="182880" rIns="182880" anchor="ctr"><a:spAutoFit/></a:bodyPr>',
        text: para(
          run(
            'This callout grows as you type: it uses shape autofit.',
            'sz="2000" i="1"',
            solid(INK)
          ),
          '<a:pPr algn="ctr"/>'
        ),
      }),
    notes: 'Speaker notes are editable too.',
  },
];

mkdirSync(OUT, { recursive: true });
writeFileSync(
  `${OUT}/sample.pptx`,
  buildPptx({
    width: 12_192_000,
    height: 6_858_000,
    slides,
    majorFont: 'Calibri',
    minorFont: 'Calibri',
    footers: true,
  })
);
writeFileSync(`${OUT}/blank.pptx`, buildBlank());
writeFileSync(`${OUT}/showcase.pptx`, buildShowcase());
console.log(`Wrote ${OUT}/sample.pptx, ${OUT}/blank.pptx, and ${OUT}/showcase.pptx`);
