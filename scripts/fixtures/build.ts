// Build the synthetic fixtures: one deck per feature, committed under fixtures/. CI's fidelity
// gate runs on these, never on the private corpus.
//
//   bun scripts/fixtures/build.ts
//
// Each fixture also needs a reference PDF next to it (`CONTRIBUTING.md`, "Fidelity harness").

import { mkdirSync, writeFileSync } from 'node:fs';
import { PNG } from 'pngjs';
import { buildPptx, cell, para, run, shape, solid, table, type FixtureSlide } from './builder.ts';
import { slideLayoutsDeck } from './slide-layouts.ts';
import { fillFidelityDeck } from './fill-fidelity.ts';
import { slideStructureDeck } from './slide-structure.ts';
import { shapeFormatFixtures } from './shape-format.ts';
import { eastAsianThemeFontsDeck, eastAsianTypographyDeck, textScriptFixtures } from './text-scripts.ts';
import { followUpFixtures } from './follow-ups.ts';
import { lineEndsDashesFixtures } from './line-ends-dashes.ts';
import { chartSizesDeck, chartsDeck, diagramsDeck } from './drawingml.ts';
import { chartDataDeck } from './chart-data.ts';
import { diagramEditingDeck } from './diagram-editing.ts';
import { buildRoundTripDecks } from './round-trip.ts';
import { pictureBulletsDeck } from './picture-bullets.ts';
import { textWarpFixtures } from './text-warp.ts';
import { equationFixtures } from './equations.ts';
import { THREE_D_EFFECT_STYLES, threeDFixtures } from './three-d.ts';

const OUT = 'fixtures';
const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

/** A deterministic 128×96 test image: a hue ramp with a dark border and a diagonal. */
function testPng(): Uint8Array {
  const png = new PNG({ width: 128, height: 96 });
  for (let y = 0; y < 96; y += 1)
    for (let x = 0; x < 128; x += 1) {
      const i = (y * 128 + x) * 4;
      const border = x < 4 || y < 4 || x > 123 || y > 91;
      const diag = Math.abs(x * 0.75 - y) < 2;
      png.data[i] = border || diag ? 30 : Math.round((x / 127) * 255);
      png.data[i + 1] = border || diag ? 30 : Math.round((y / 95) * 255);
      png.data[i + 2] = border || diag ? 30 : 160;
      png.data[i + 3] = 255;
    }
  return new Uint8Array(PNG.sync.write(png));
}

/** A second deterministic image: a 64×48 solid teal block, so two pictures are distinct parts. */
function solidPng(): Uint8Array {
  const png = new PNG({ width: 64, height: 48 });
  for (let i = 0; i < 64 * 48 * 4; i += 4) png.data.set([0, 128, 128, 255], i);
  return new Uint8Array(PNG.sync.write(png));
}

const box = (id: number, x: number, y: number, cx: number, cy: number, text: string, extra = {}): string =>
  shape({ id, x, y, cx, cy, text, bodyPr: '<a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"><a:noAutofit/></a:bodyPr>', ...extra });

const LOREM =
  'The quick brown fox jumps over the lazy dog while seven wizards quietly judge boxing matches in the park.';

const DEFAULT_STYLE = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';
/** A style defined only by this file, so a reader must use the file's definition. */
const CUSTOM_STYLE = '{0C3A3B8E-5A47-4F0B-9C1D-6E2F7A8B9C01}';
const ln = (w: number, color: string, dash?: string): string =>
  `<a:ln w="${w}" cmpd="sng"><a:solidFill>${color}</a:solidFill>${dash ? `<a:prstDash val="${dash}"/>` : ''}</a:ln>`;
const scheme = (val: string, mods = ''): string => `<a:schemeClr val="${val}">${mods}</a:schemeClr>`;
const CUSTOM_STYLE_XML =
  `<a:tblStyle styleId="${CUSTOM_STYLE}" styleName="Fixture Style - Accent 2">` +
  `<a:tblBg><a:fill><a:solidFill>${scheme('bg1', '<a:lumMod val="95000"/>')}</a:solidFill></a:fill></a:tblBg>` +
  `<a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"/>${scheme('tx1')}</a:tcTxStyle><a:tcStyle><a:tcBdr>` +
  `<a:left><a:lnRef idx="1">${scheme('accent2')}</a:lnRef></a:left><a:right><a:lnRef idx="1">${scheme('accent2')}</a:lnRef></a:right>` +
  `<a:top>${ln(25400, scheme('accent2'))}</a:top><a:bottom>${ln(25400, scheme('accent2'))}</a:bottom>` +
  `<a:insideH>${ln(6350, scheme('accent2', '<a:tint val="50000"/>'))}</a:insideH><a:insideV><a:ln><a:noFill/></a:ln></a:insideV>` +
  '</a:tcBdr></a:tcStyle></a:wholeTbl>' +
  `<a:band1V><a:tcStyle><a:tcBdr/><a:fill><a:solidFill>${scheme('accent2', '<a:tint val="20000"/>')}</a:solidFill></a:fill></a:tcStyle></a:band1V>` +
  `<a:firstCol><a:tcTxStyle b="on" i="on"><a:fontRef idx="major"/>${scheme('lt1')}</a:tcTxStyle><a:tcStyle><a:tcBdr/><a:fill><a:solidFill>${scheme('accent2')}</a:solidFill></a:fill></a:tcStyle></a:firstCol>` +
  `<a:lastRow><a:tcTxStyle b="on"/><a:tcStyle><a:tcBdr><a:top>${ln(38100, scheme('accent2'))}</a:top></a:tcBdr></a:tcStyle></a:lastRow>` +
  `<a:swCell><a:tcStyle><a:tcBdr/><a:fill><a:solidFill>${scheme('accent6')}</a:solidFill></a:fill></a:tcStyle></a:swCell>` +
  '</a:tblStyle>';
const grid = (color = '404040'): string =>
  ['lnL', 'lnR', 'lnT', 'lnB'].map((n) => `<a:${n} w="12700"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:${n}>`).join('');
const rows = (h: number, cells: readonly string[][]): string =>
  cells.map((r) => `<a:tr h="${h}">${r.join('')}</a:tr>`).join('');
const tableStylesFor: Record<string, string> = { tables: CUSTOM_STYLE_XML };

/** The ISO A4 paper size as a custom slide size, 842 × 595.5 pt. Its PDF page is 842 × 596 pt. */
const A4 = { width: 10_693_400, height: 7_562_850 };
/**
 * 400.25 × 300.5 pt. The reference PDF rounds the width down to a 400 pt page, which ends 0.25 pt
 * (1 px at 1920 px wide) before the slide, and the height up to 301 pt.
 */
const ROUNDING = { width: 5_083_175, height: 3_816_350 };
/** Slide sizes other than the 16:9 default. */
const slideSizeFor: Record<string, { width: number; height: number }> = {
  'slide-size': A4,
  'slide-size-rounding': ROUNDING,
};
/** A blue bar on the bottom edge, an orange bar on the right edge, and a green square at the top-left. */
const edgeShapes = (size: { width: number; height: number }): string =>
  shape({ id: 2, x: 0, y: size.height - 0.35 * IN, cx: size.width, cy: 0.35 * IN, fill: solid('1F4E79') }) +
  shape({ id: 3, x: size.width - 0.35 * IN, y: 0, cx: 0.35 * IN, cy: size.height - 0.35 * IN, fill: solid('ED7D31') }) +
  shape({ id: 4, x: 0, y: 0, cx: 0.35 * IN, cy: 0.35 * IN, fill: solid('70AD47') });

const fixtures: Record<string, FixtureSlide[]> = {
  ...shapeFormatFixtures,
  ...textWarpFixtures,
  ...equationFixtures,
  ...threeDFixtures,
  ...textScriptFixtures,
  ...followUpFixtures,
  ...lineEndsDashesFixtures,
  'tables': [
    {
      // The default style, referenced but not written in tableStyles.xml.
      shapes: table({
        id: 2,
        x: IN,
        y: IN,
        cx: 7.5 * IN,
        cy: 5 * 370_840,
        columns: [2.5 * IN, 2.5 * IN, 2.5 * IN],
        tblPr: `<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${DEFAULT_STYLE}</a:tableStyleId></a:tblPr>`,
        rows: rows(370_840, [
          [cell('Region'), cell('Q1'), cell('Q2')],
          [cell('North'), cell('1,240'), cell('1,310')],
          [cell('South'), cell('980'), cell('1,045')],
          [cell('East'), cell('1,105'), cell('1,190')],
          [cell('West'), cell('870'), cell('925')],
        ]),
      }),
    },
    {
      // A style from tableStyles.xml with first column, last row, and banded columns.
      shapes: table({
        id: 2,
        x: IN,
        y: IN,
        cx: 8 * IN,
        cy: 4 * 457_200,
        columns: [2 * IN, 2 * IN, 2 * IN, 2 * IN],
        tblPr: `<a:tblPr firstCol="1" lastRow="1" bandCol="1"><a:tableStyleId>${CUSTOM_STYLE}</a:tableStyleId></a:tblPr>`,
        rows: rows(457_200, [
          [cell('Item'), cell('Cost'), cell('Units'), cell('Total')],
          [cell('Paper'), cell('4.00'), cell('12'), cell('48.00')],
          [cell('Ink'), cell('18.50'), cell('3'), cell('55.50')],
          [cell('Sum'), cell(''), cell('15'), cell('103.50')],
        ]),
      }),
    },
    {
      // No style: merges, direct borders, diagonals, fills, margins, anchors, vertical text,
      // and a row that grows to fit its text.
      shapes: table({
        id: 2,
        x: 0.5 * IN,
        y: 0.5 * IN,
        cx: 9 * IN,
        cy: 3 * 548_640,
        columns: [2.5 * IN, 2 * IN, 2 * IN, 2.5 * IN],
        tblPr: '<a:tblPr/>',
        rows: rows(548_640, [
          [
            cell('Merged across two columns', { attrs: ' gridSpan="2"', tcPr: `${grid()}<a:solidFill><a:srgbClr val="D9E2F3"/></a:solidFill>` }),
            cell('', { attrs: ' hMerge="1"', tcPr: grid() }),
            cell('Anchored bottom', { tcPrAttrs: ' anchor="b"', tcPr: grid() }),
            cell('Merged down, centered', {
              attrs: ' rowSpan="2"',
              tcPrAttrs: ' anchor="ctr"',
              tcPr: `${grid()}<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFF2CC"/></a:gs><a:gs pos="100000"><a:srgbClr val="F4B183"/></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>`,
            }),
          ],
          [
            cell('Diagonal', { tcPr: `${grid()}<a:lnTlToBr w="12700"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:lnTlToBr>` }),
            cell('Dashed', {
              tcPr: ['lnL', 'lnR', 'lnT', 'lnB'].map((n) => `<a:${n} w="19050"><a:solidFill><a:srgbClr val="2E75B6"/></a:solidFill><a:prstDash val="dash"/></a:${n}>`).join(''),
            }),
            cell('Vertical', { tcPrAttrs: ' vert="vert270" anchor="ctr"', tcPr: grid() }),
            cell('', { attrs: ' vMerge="1"', tcPr: grid() }),
          ],
          [
            cell(LOREM, { tcPr: grid() }),
            cell('No margins', { tcPrAttrs: ' marL="0" marR="0" marT="0" marB="0"', tcPr: grid() }),
            cell('Bold', { rPr: ' b="1"', tcPr: grid() }),
            cell('Two\nparagraphs', { tcPr: grid() }),
          ],
        ]),
      }),
    },
  ],

  // Tables in the exact form the table commands write them (editor/table-*.ts), so a
  // reference PDF of this deck checks what the editor saves: Insert > Table, a merged block,
  // shading, borders, margins, anchors, and split cells.
  'table-editing': [
    {
      // Insert > Table (3 × 4, two thirds of the slide wide), then Merge Cells on a 2 × 2 block,
      // Shading, Borders > Outside, Cell Margins > Wide, and Align Center.
      shapes: table({
        id: 2,
        x: 2_032_000,
        y: IN,
        cx: 3 * 2_709_333,
        cy: 4 * 370_840,
        columns: [2_709_333, 2_709_333, 2_709_333],
        tblPr: `<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${DEFAULT_STYLE}</a:tableStyleId></a:tblPr>`,
        rows: rows(370_840, [
          [cell('Plan'), cell('Owner'), cell('Due')],
          [
            cell('Merged\nblock', { attrs: ' gridSpan="2" rowSpan="2"', tcPrAttrs: ' anchor="ctr"' }),
            cell('', { attrs: ' hMerge="1" rowSpan="2"' }),
            cell('Shaded', { tcPr: '<a:solidFill><a:srgbClr val="FFC000"/></a:solidFill>' }),
          ],
          [
            cell('', { attrs: ' gridSpan="2" vMerge="1"' }),
            cell('', { attrs: ' hMerge="1" vMerge="1"' }),
            cell('Outside border', {
              tcPr: ['lnL', 'lnR', 'lnT', 'lnB']
                .map(
                  (n) =>
                    `<a:${n} w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:${n}>`
                )
                .join(''),
            }),
          ],
          [
            cell('Wide margins', { tcPrAttrs: ' marL="137160" marR="137160" marT="137160" marB="137160"' }),
            cell('No fill', { tcPr: '<a:noFill/>' }),
            cell('Bottom', { tcPrAttrs: ' anchor="b"' }),
          ],
        ]),
      }),
    },
    {
      // Split Cells (2 columns, 2 rows) on the middle cell of a 3 × 3 table: the column splits
      // into two equal grid columns and the row into two rows; the cells beside them span.
      shapes: table({
        id: 2,
        x: IN,
        y: IN,
        cx: 6 * IN,
        cy: 4 * 370_840,
        columns: [2 * IN, IN, IN, 2 * IN],
        tblPr: `<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${DEFAULT_STYLE}</a:tableStyleId></a:tblPr>`,
        // The split row's height is shared by its two rows.
        rows:
          rows(370_840, [[cell('A'), cell('B', { attrs: ' gridSpan="2"' }), cell('', { attrs: ' hMerge="1"' }), cell('C')]]) +
          rows(185_420, [
            [cell('D', { attrs: ' rowSpan="2"' }), cell('E1'), cell('E2'), cell('F', { attrs: ' rowSpan="2"' })],
            [cell('', { attrs: ' vMerge="1"' }), cell('E3'), cell('E4'), cell('', { attrs: ' vMerge="1"' })],
          ]) +
          rows(370_840, [[cell('G'), cell('H', { attrs: ' gridSpan="2"' }), cell('', { attrs: ' hMerge="1"' }), cell('I')]]),
      }),
    },
  ],

  'text-runs': [
    {
      shapes:
        box(2, IN / 2, IN / 2, 6 * IN, 4 * IN,
          para(run('Regular 18 ') + run('Bold ', 'b="1"') + run('Italic ', 'i="1"') + run('Underline ', 'u="sng"') + run('Strike', 'strike="sngStrike"')) +
          para(run('Size 32 bold red', 'sz="3200" b="1"', solid('C00000'))) +
          para(run('All caps text', 'cap="all"') + run(' and small caps', 'cap="small"')) +
          para(run('Spaced out', 'spc="600"') + run(' tight', 'spc="-100"')) +
          para(run('H') + run('2', 'baseline="-25000"') + run('O and x') + run('2', 'baseline="30000"')) +
          para(run('Highlighted text', '', '<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>'))
        ) +
        box(3, 7 * IN, IN / 2, 5 * IN, 4 * IN, para(run(LOREM)) + para(run(LOREM, 'sz="1400"'))),
    },
  ],
  'bullets-numbering': [
    {
      layout: 'title',
      shapes:
        shape({ id: 2, x: 0, y: 0, cx: 0, cy: 0, ph: '<p:ph type="title"/>', text: para(run('Bullets and numbering')) }).replace('<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst></a:avLst></a:prstGeom>', '') +
        shape({
          id: 3, x: 0, y: 0, cx: 0, cy: 0, ph: '<p:ph idx="1"/>',
          text:
            para(run('First level bullet from the master body style')) +
            para(run('Second level'), '<a:pPr lvl="1"/>') +
            para(run('Third level'), '<a:pPr lvl="2"/>') +
            para(run('Numbered one'), '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/></a:pPr>') +
            para(run('Numbered two'), '<a:pPr marL="342900" indent="-342900"><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/></a:pPr>') +
            para(run('Roman with color'), '<a:pPr marL="342900" indent="-342900"><a:buClr><a:srgbClr val="2E7D32"/></a:buClr><a:buAutoNum type="romanUcParenR"/></a:pPr>') +
            para(run('Custom character, 150%'), '<a:pPr marL="342900" indent="-342900"><a:buSzPct val="150000"/><a:buFont typeface="Liberation Sans"/><a:buChar char="–"/></a:pPr>') +
            para(run('No bullet'), '<a:pPr marL="0" indent="0"><a:buNone/></a:pPr>'),
        }).replace('<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst></a:avLst></a:prstGeom>', ''),
    },
    {
      // Numbering restarts under each parent and on a scheme or startAt change; bullets move
      // with centered and right-aligned text; a trailing break adds an empty line.
      shapes:
        box(2, 0.5 * IN, 0.4 * IN, 5.6 * IN, 4.2 * IN,
          [
            [0, 'Parent one', 'arabicPeriod', ''],
            [1, 'Child a', 'alphaLcPeriod', ''],
            [1, 'Child b', 'alphaLcPeriod', ''],
            [0, 'Parent two', 'arabicPeriod', ''],
            [1, 'Child restarts at a', 'alphaLcPeriod', ''],
            [0, 'Scheme change restarts', 'romanUcPeriod', ''],
            [0, 'Then continues', 'romanUcPeriod', ''],
            [0, 'startAt 5', 'arabicPeriod', ' startAt="5"'],
          ]
            .map(([lvl, t, scheme, startAt]) =>
              para(run(String(t), 'sz="1800"'), `<a:pPr lvl="${lvl}" marL="${(Number(lvl) + 1) * 457200}" indent="-342900"><a:buFont typeface="+mj-lt"/><a:buAutoNum type="${scheme}"${startAt}/></a:pPr>`))
            .join(''), { fill: solid('F2F2F2') }) +
        ['l', 'ctr', 'r']
          .map((algn, i) =>
            box(3 + i, 6.6 * IN, (0.4 + i * 1.1) * IN, 5 * IN, 0.9 * IN,
              para(run(`Bullet with algn="${algn}"`, 'sz="1800"'), `<a:pPr algn="${algn}" marL="342900" indent="-342900"><a:buFont typeface="Liberation Sans"/><a:buChar char="•"/></a:pPr>`), { fill: solid('DEEBF7') }))
          .join('') +
        box(6, 6.6 * IN, 3.8 * IN, 5 * IN, 2.6 * IN,
          '<a:p><a:r><a:rPr lang="en-US" sz="2400" dirty="0"/><a:t>Trailing break</a:t></a:r><a:br><a:rPr lang="en-US" sz="2400" dirty="0"/></a:br><a:endParaRPr lang="en-US" sz="2400" dirty="0"/></a:p>',
          { fill: solid('FFF2CC'), bodyPr: '<a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="b"><a:noAutofit/></a:bodyPr>' }),
    },
  ],
  'spacing': [
    {
      shapes:
        box(2, IN / 2, IN / 2, 3.6 * IN, 5 * IN,
          para(run('Line spacing 100% ' + LOREM.slice(0, 60))) +
          para(run('Line spacing 150% ' + LOREM.slice(0, 60)), '<a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:pPr>') +
          para(run('Line spacing exactly 30 pt ' + LOREM.slice(0, 50)), '<a:pPr><a:lnSpc><a:spcPts val="3000"/></a:lnSpc></a:pPr>')) +
        box(3, 4.4 * IN, IN / 2, 3.6 * IN, 5 * IN,
          para(run('Space before 24 pt'), '<a:pPr><a:spcBef><a:spcPts val="2400"/></a:spcBef></a:pPr>') +
          para(run('Space after 18 pt'), '<a:pPr><a:spcAft><a:spcPts val="1800"/></a:spcAft></a:pPr>') +
          para(run('Indented paragraph with a hanging first line ' + LOREM.slice(0, 40)), '<a:pPr marL="457200" indent="-228600"/>')) +
        box(4, 8.3 * IN, IN / 2, 3.6 * IN, 5 * IN,
          para(run('Mixed sizes '), '') + para(run('small ', 'sz="1000"') + run('LARGE', 'sz="4000"') + run(' line height follows the largest')),
          { bodyPr: '<a:bodyPr wrap="square" spcFirstLastPara="1"><a:noAutofit/></a:bodyPr>' }),
    },
  ],
  // What the Bullets and Numbering galleries write: the standard bullet characters with their
  // fonts (an open symbol face here), a custom character in the text font, the numbering schemes,
  // a size, a color, and startAt.
  'bullets-gallery': [
    {
      shapes:
        box(2, 0.4 * IN, 0.4 * IN, 6 * IN, 6.6 * IN,
          [
            ['Filled round', '<a:buFont typeface="Liberation Sans"/><a:buChar char="•"/>'],
            ['Hollow round', '<a:buFont typeface="Liberation Mono"/><a:buChar char="o"/>'],
            ['Filled square', '<a:buFont typeface="Noto Sans Symbols 2"/><a:buChar char="▪"/>'],
            ['Hollow square', '<a:buFont typeface="Noto Sans Symbols 2"/><a:buChar char="❑"/>'],
            ['Star', '<a:buFont typeface="Noto Sans Symbols 2"/><a:buChar char="❖"/>'],
            ['Arrow', '<a:buFont typeface="Noto Sans Symbols 2"/><a:buChar char="➢"/>'],
            ['Checkmark', '<a:buFont typeface="Noto Sans Symbols 2"/><a:buChar char="✓"/>'],
            ['Size 150 %, accent color', '<a:buClr><a:schemeClr val="accent2"/></a:buClr><a:buSzPct val="150000"/><a:buFont typeface="Liberation Sans"/><a:buChar char="•"/>'],
            ['Custom en dash, text font', '<a:buChar char="–"/>'],
          ]
            .map(([text, bullet]) => para(run(text, 'sz="2000"'), `<a:pPr marL="342900" indent="-342900">${bullet}</a:pPr>`))
            .join(''),
          { fill: solid('EEF3FA') }) +
        box(3, 6.9 * IN, 0.4 * IN, 6 * IN, 6.6 * IN,
          [
            ['arabicPeriod', ''],
            ['arabicParenR', ''],
            ['romanUcPeriod', ''],
            ['alphaUcPeriod', ''],
            ['alphaLcParenR', ''],
            ['alphaLcPeriod', ''],
            ['romanLcPeriod', ''],
            ['arabicPeriod', ' startAt="3"'],
          ]
            .map(([scheme, startAt]) =>
              para(run(`${scheme}${startAt}`, 'sz="2000"'), `<a:pPr marL="457200" indent="-457200"><a:buFont typeface="+mj-lt"/><a:buAutoNum type="${scheme}"${startAt}/></a:pPr>`))
            .join(''),
          { fill: solid('FDF2E9') }),
    },
  ],
  // What the Text Direction and Columns menus write, and the two vertical types the editor
  // refuses to set (drawn approximately: "Other vertical text types" on the public fidelity page).
  'text-direction': [
    {
      shapes:
        [
          ['vert', 'Rotate all text 90°'],
          ['vert270', 'Rotate all text 270°'],
          ['wordArtVert', 'Stacked'],
          ['mongolianVert', 'Mongolian vert, two lines'],
          ['wordArtVertRtl', 'Stacked RTL'],
        ]
          .map(([vert, text], i) =>
            shape({ id: 2 + i, x: (0.3 + i * 1.3) * IN, y: 0.4 * IN, cx: 1.1 * IN, cy: 4.2 * IN, fill: solid('F3E5F5'),
              bodyPr: `<a:bodyPr vert="${vert}" wrap="square"><a:noAutofit/></a:bodyPr>`, text: para(run(text, 'sz="1600"')) }))
          .join('') +
        shape({ id: 10, x: 7 * IN, y: 0.4 * IN, cx: 6 * IN, cy: 3.2 * IN, fill: solid('E0F7FA'),
          bodyPr: '<a:bodyPr numCol="3" spcCol="274320" wrap="square"><a:noAutofit/></a:bodyPr>',
          text: [1, 2, 3, 4, 5].map(() => para(run(LOREM, 'sz="1400"'))).join('') }) +
        shape({ id: 11, x: 7 * IN, y: 3.9 * IN, cx: 6 * IN, cy: 3.2 * IN, fill: solid('FFF8E1'),
          bodyPr: '<a:bodyPr numCol="2" wrap="square"><a:noAutofit/></a:bodyPr>',
          text: [1, 2, 3, 4, 5, 6].map(() => para(run(LOREM, 'sz="1400"'))).join('') }) +
        shape({ id: 12, x: 0.3 * IN, y: 4.9 * IN, cx: 6.4 * IN, cy: 2.2 * IN, fill: solid('E8F5E9'),
          bodyPr: '<a:bodyPr vert="horz" wrap="square"><a:noAutofit/></a:bodyPr>', text: para(run('Horizontal: ' + LOREM, 'sz="1600"')) }),
    },
  ],
  // What the Font dialog and the Font group write: offsets, strikethrough kinds, character
  // spacing presets, kerning, caps, underline styles, and underline colors (`a:uFill`).
  'font-effects': [
    {
      shapes:
        box(2, 0.4 * IN, 0.4 * IN, 6.2 * IN, 6.6 * IN,
          para(run('Superscript x') + run('2', 'baseline="30000"') + run(', offset 50% x') + run('3', 'baseline="50000"')) +
          para(run('Subscript H') + run('2', 'baseline="-25000"') + run('O, offset -40% C') + run('6', 'baseline="-40000"')) +
          para(run('Single strike', 'strike="sngStrike"') + run(' and ') + run('double strike', 'strike="dblStrike"')) +
          para(run('Very tight', 'spc="-300"') + run(' tight', 'spc="-150"') + run(' normal', 'spc="0"')) +
          para(run('Loose', 'spc="300"') + run(' very loose', 'spc="600"')) +
          para(run('Small caps', 'cap="small"') + run(' and ') + run('all caps', 'cap="all"')),
          { fill: solid('EEF3FA') }) +
        box(3, 6.9 * IN, 0.4 * IN, 6 * IN, 6.6 * IN,
          para(run('AVATAR WAVE kerned', 'sz="2800" kern="1200"')) +
          para(run('AVATAR WAVE not kerned', 'sz="2800" kern="0"')) +
          para(run('Double', 'u="dbl"') + run(' heavy', 'u="heavy"') + run(' dotted', 'u="dotted"') + run(' wavy', 'u="wavy"')) +
          para(run('Red underline', 'u="sng"', '<a:uFill><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:uFill>') +
            run(' blue text, accent line', 'u="dbl"', '<a:solidFill><a:srgbClr val="0070C0"/></a:solidFill><a:uFill><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:uFill>')) +
          para(run('Change Case Writes Letters, Not Formatting')),
          { fill: solid('FDF2E9') }),
    },
  ],
  // What the Paragraph dialog writes: spacing before and after in lines (`a:spcPct`) and points,
  // first-line and hanging indents, and line spacing as a multiple and exactly.
  'paragraph-format': [
    {
      shapes:
        box(2, 0.4 * IN, 0.4 * IN, 4 * IN, 6.6 * IN,
          para(run('Space before 0.5 lines'), '<a:pPr><a:spcBef><a:spcPct val="50000"/></a:spcBef></a:pPr>') +
          para(run('Space after 1 line'), '<a:pPr><a:spcAft><a:spcPct val="100000"/></a:spcAft></a:pPr>') +
          para(run('Before 12 pt, after 6 pt'), '<a:pPr><a:spcBef><a:spcPts val="1200"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft></a:pPr>') +
          para(run('Before 10.5 pt'), '<a:pPr><a:spcBef><a:spcPts val="1050"/></a:spcBef></a:pPr>') +
          para(run('No own spacing')),
          { fill: solid('EEF3FA') }) +
        box(3, 4.6 * IN, 0.4 * IN, 4 * IN, 6.6 * IN,
          para(run('First line 0.5 in: ' + LOREM.slice(0, 70)), '<a:pPr marL="0" indent="457200"/>') +
          para(run('Hanging 0.5 in: ' + LOREM.slice(0, 70)), '<a:pPr marL="457200" indent="-457200"/>') +
          para(run('Before text 1 in: ' + LOREM.slice(0, 50)), '<a:pPr marL="914400" indent="0"/>') +
          para(run('Right aligned, before text 0.25 in'), '<a:pPr marL="228600" algn="r"/>'),
          { fill: solid('FDF2E9') }) +
        box(4, 8.8 * IN, 0.4 * IN, 4.1 * IN, 6.6 * IN,
          para(run('Multiple 0.8: ' + LOREM.slice(0, 60)), '<a:pPr><a:lnSpc><a:spcPct val="80000"/></a:lnSpc></a:pPr>') +
          para(run('Double: ' + LOREM.slice(0, 60)), '<a:pPr><a:lnSpc><a:spcPct val="200000"/></a:lnSpc></a:pPr>') +
          para(run('Exactly 20 pt: ' + LOREM.slice(0, 60)), '<a:pPr><a:lnSpc><a:spcPts val="2000"/></a:lnSpc></a:pPr>') +
          para(run('Justified: ' + LOREM), '<a:pPr algn="just"/>'),
          { fill: solid('E8F5E9') }),
    },
  ],
  'alignment-anchor': [
    {
      shapes: (['l', 'ctr', 'r', 'just'] as const)
        .map((algn, i) =>
          shape({ id: 2 + i, x: (0.3 + i * 3.1) * IN, y: 0.4 * IN, cx: 2.9 * IN, cy: 2.8 * IN, fill: solid('EEF3FA'), line: '<a:ln w="9525"><a:solidFill><a:srgbClr val="7F7F7F"/></a:solidFill></a:ln>',
            bodyPr: '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>', text: para(run(`${algn}: ${LOREM}`, 'sz="1400"'), `<a:pPr algn="${algn}"/>`) }))
        .join('') +
        (['t', 'ctr', 'b'] as const)
          .map((anchor, i) =>
            shape({ id: 10 + i, x: (0.3 + i * 4.1) * IN, y: 3.6 * IN, cx: 3.9 * IN, cy: 3.4 * IN, fill: solid('FDF2E9'), line: '<a:ln w="9525"><a:solidFill><a:srgbClr val="7F7F7F"/></a:solidFill></a:ln>',
              bodyPr: `<a:bodyPr wrap="square" anchor="${anchor}"><a:noAutofit/></a:bodyPr>`, text: para(run(`anchor=${anchor}`, 'sz="2000"'), '<a:pPr algn="ctr"/>') }))
          .join(''),
    },
  ],
  'autofit': [
    {
      shapes:
        shape({ id: 2, x: 0.5 * IN, y: 0.5 * IN, cx: 5 * IN, cy: 2.2 * IN, fill: solid('E8F5E9'),
          bodyPr: '<a:bodyPr wrap="square"><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>',
          text: para(run(LOREM + ' ' + LOREM, 'sz="2800"')) }) +
        shape({ id: 3, x: 6.5 * IN, y: 0.5 * IN, cx: 5 * IN, cy: 0.5 * IN, fill: solid('E3F2FD'),
          bodyPr: '<a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr>', text: para(run('Shape grows to fit: ' + LOREM)) }) +
        shape({ id: 4, x: 0.5 * IN, y: 3.5 * IN, cx: 5 * IN, cy: 1 * IN, fill: solid('FFF3E0'),
          bodyPr: '<a:bodyPr wrap="none"><a:noAutofit/></a:bodyPr>', text: para(run('wrap="none" keeps this on one line even past the edge')) }),
    },
  ],
  // The line model's thinly measured cases:
  // exact spacing below and above the single line, 80–200 %, fractional paragraph spacing,
  // overflowing anchored text, and stored autofit pairs.
  'line-model': [
    {
      shapes:
        box(2, 0.3 * IN, 0.3 * IN, 3 * IN, 6.9 * IN,
          para(run(LOREM, 'sz="1400"', '<a:latin typeface="Liberation Sans"/>'), '<a:pPr><a:lnSpc><a:spcPts val="1560"/></a:lnSpc></a:pPr>') +
          para(run(LOREM.slice(0, 60), 'sz="1800"', '<a:latin typeface="Liberation Sans"/>'), '<a:pPr><a:lnSpc><a:spcPts val="2100"/></a:lnSpc></a:pPr>') +
          para(run(LOREM.slice(0, 60), 'sz="1300"', '<a:latin typeface="Liberation Sans"/>'), '<a:pPr><a:lnSpc><a:spcPts val="1680"/></a:lnSpc></a:pPr>') +
          para(run(LOREM.slice(0, 60), 'sz="1600"'), '<a:pPr><a:lnSpc><a:spcPts val="1200"/></a:lnSpc></a:pPr>')) +
        box(3, 3.5 * IN, 0.3 * IN, 3 * IN, 6.9 * IN,
          ['80000', '115000', '150000', '200000']
            .map((pct) => para(run(`${Number(pct) / 1000} %: ${LOREM.slice(0, 44)}`, 'sz="2000"'), `<a:pPr><a:lnSpc><a:spcPct val="${pct}"/></a:lnSpc></a:pPr>`))
            .join('')) +
        box(4, 6.7 * IN, 0.3 * IN, 3 * IN, 6.9 * IN,
          para(run('First, space before 6.4 pt'), '<a:pPr><a:spcBef><a:spcPts val="640"/></a:spcBef></a:pPr>') +
          para(run('Space after 7.5 pt'), '<a:pPr><a:spcAft><a:spcPts val="750"/></a:spcAft></a:pPr>') +
          para(run('Space before 20 % of the line'), '<a:pPr><a:spcBef><a:spcPct val="20000"/></a:spcBef></a:pPr>') +
          para(run('Space before 3.2 pt'), '<a:pPr><a:spcBef><a:spcPts val="320"/></a:spcBef></a:pPr>'),
          { bodyPr: '<a:bodyPr wrap="square" spcFirstLastPara="1"><a:noAutofit/></a:bodyPr>' }) +
        box(5, 9.9 * IN, 0.3 * IN, 3.1 * IN, 6.9 * IN,
          para(run('Mixed ', 'sz="1200"') + run('sizes', 'sz="2800"') + run(' on one line, then a second line', 'sz="1200"')) +
          para(run(LOREM.slice(0, 50), 'sz="2400"', '<a:latin typeface="Liberation Serif"/>'), '<a:pPr><a:lnSpc><a:spcPct val="120000"/></a:lnSpc></a:pPr>')),
    },
    {
      shapes:
        ([['ctr', 0.3], ['b', 3.5]] as const)
          .map(([anchor, x], i) =>
            box(2 + i, x * IN, 0.3 * IN, 3 * IN, 1 * IN, para(run(`anchor=${anchor}, taller than the box: ${LOREM.slice(0, 60)}`, 'sz="2000"')),
              { bodyPr: `<a:bodyPr wrap="square" anchor="${anchor}"><a:noAutofit/></a:bodyPr>`, fill: solid('FDF2E9') }))
          .join('') +
        box(4, 6.7 * IN, 0.3 * IN, 3 * IN, 1.13 * IN, para(run('anchor=ctr, fractional offset', 'sz="2000"')),
          { bodyPr: '<a:bodyPr wrap="square" anchor="ctr"><a:noAutofit/></a:bodyPr>', fill: solid('FDF2E9') }) +
        ([
          ['92500', '10000', '2000', '110000'],
          ['85000', '20000', '2000', '90000'],
          ['77500', '0', '2400', '100000'],
        ] as const)
          .map(([scale, reduction, sz, pct], i) =>
            box(5 + i, (0.3 + i * 3.2) * IN, 3.2 * IN, 3 * IN, 3.8 * IN,
              para(run(`fontScale ${Number(scale) / 1000} %, lnSpcReduction ${Number(reduction) / 1000} %: ${LOREM}`, `sz="${sz}"`), `<a:pPr><a:lnSpc><a:spcPct val="${pct}"/></a:lnSpc></a:pPr>`),
              { bodyPr: `<a:bodyPr wrap="square"><a:normAutofit fontScale="${scale}"${reduction === '0' ? '' : ` lnSpcReduction="${reduction}"`}/></a:bodyPr>`, fill: solid('E8F5E9') }))
          .join(''),
    },
  ],
  'vertical-columns': [
    {
      shapes:
        shape({ id: 2, x: 0.5 * IN, y: 0.5 * IN, cx: 1.2 * IN, cy: 5 * IN, fill: solid('F3E5F5'), bodyPr: '<a:bodyPr vert="vert" wrap="square"><a:noAutofit/></a:bodyPr>', text: para(run('Vertical text vert')) }) +
        shape({ id: 3, x: 2 * IN, y: 0.5 * IN, cx: 1.2 * IN, cy: 5 * IN, fill: solid('E0F7FA'), bodyPr: '<a:bodyPr vert="vert270" wrap="square"><a:noAutofit/></a:bodyPr>', text: para(run('Vertical text vert270')) }) +
        shape({ id: 4, x: 3.6 * IN, y: 0.5 * IN, cx: 8 * IN, cy: 5 * IN, fill: solid('FAFAFA'), bodyPr: '<a:bodyPr numCol="2" spcCol="365760" wrap="square"><a:noAutofit/></a:bodyPr>', text: [1, 2, 3, 4].map(() => para(run(LOREM))).join('') }),
    },
  ],
  'colors': [
    {
      shapes: [
        ['accent1', ''], ['accent1', '<a:lumMod val="75000"/>'], ['accent1', '<a:lumMod val="60000"/><a:lumOff val="40000"/>'],
        ['accent2', '<a:shade val="50000"/>'], ['accent2', '<a:tint val="40000"/>'], ['accent6', '<a:satMod val="200000"/>'],
        ['tx1', '<a:lumMod val="50000"/><a:lumOff val="50000"/>'], ['accent4', '<a:alpha val="50000"/>'], ['accent5', '<a:hueOff val="5400000"/>'],
        ['accent3', '<a:comp/>'], ['accent1', '<a:inv/>'], ['accent2', '<a:gray/>'],
      ]
        .map(([c, t], i) =>
          shape({ id: 2 + i, x: (0.4 + (i % 6) * 2) * IN, y: (0.6 + Math.floor(i / 6) * 3) * IN, cx: 1.8 * IN, cy: 2.4 * IN,
            fill: `<a:solidFill><a:schemeClr val="${c}">${t}</a:schemeClr></a:solidFill>` }))
        .join(''),
    },
  ],
  'preset-shapes': [
    {
      shapes: ['rect', 'roundRect', 'ellipse', 'hexagon', 'rightArrow', 'flowChartOffpageConnector', 'triangle', 'star5', 'chevron', 'cloud', 'heart', 'wedgeRectCallout', 'donut', 'can', 'cube', 'leftRightArrow', 'pentagon', 'plus']
        .map((prst, i) =>
          shape({ id: 2 + i, x: (0.3 + (i % 6) * 2.05) * IN, y: (0.3 + Math.floor(i / 6) * 2.4) * IN, cx: 1.8 * IN, cy: 1.8 * IN, prst,
            fill: solid('5B9BD5'), line: '<a:ln w="19050"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:ln>' }))
        .join(''),
    },
    {
      shapes:
        shape({ id: 2, x: IN, y: IN, cx: 4 * IN, cy: 2 * IN, prst: 'roundRect', avLst: '<a:gd name="adj" fmla="val 40000"/>', fill: solid('70AD47') }) +
        shape({ id: 3, x: 6 * IN, y: IN, cx: 4 * IN, cy: 2 * IN, prst: 'rightArrow', avLst: '<a:gd name="adj1" fmla="val 30000"/><a:gd name="adj2" fmla="val 80000"/>', fill: solid('ED7D31') }),
    },
  ],
  'custom-geometry': [
    {
      shapes:
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Freeform"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="3200400"/></a:xfrm>' +
        '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100" h="70">' +
        '<a:moveTo><a:pt x="0" y="70"/></a:moveTo><a:lnTo><a:pt x="30" y="0"/></a:lnTo><a:cubicBezTo><a:pt x="50" y="20"/><a:pt x="70" y="-10"/><a:pt x="100" y="35"/></a:cubicBezTo>' +
        '<a:quadBezTo><a:pt x="80" y="70"/><a:pt x="50" y="60"/></a:quadBezTo><a:arcTo wR="25" hR="10" stAng="0" swAng="10800000"/><a:close/></a:path></a:pathLst></a:custGeom>' +
        `${solid('A9D18E')}<a:ln w="25400"><a:solidFill><a:srgbClr val="385723"/></a:solidFill></a:ln></p:spPr></p:sp>`,
    },
    {
      // Arcs evaluate in path units: a 4:1 cloud is the square cloud stretched. A path may
      // start with lnTo (schema-valid); it has no current point, so it draws nothing.
      shapes:
        shape({ id: 2, x: 0.5 * IN, y: 0.5 * IN, cx: 8 * IN, cy: 2 * IN, prst: 'cloud', fill: solid('DEEBF7'), line: '<a:ln w="25400"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:ln>' }) +
        shape({ id: 3, x: 9 * IN, y: 0.5 * IN, cx: 2 * IN, cy: 2 * IN, prst: 'cloud', fill: solid('DEEBF7'), line: '<a:ln w="25400"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:ln>' }) +
        '<p:sp><p:nvSpPr><p:cNvPr id="4" name="Starts with lnTo"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="3657600"/><a:ext cx="3657600" cy="2743200"/></a:xfrm>' +
        '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100" h="100">' +
        '<a:lnTo><a:pt x="100" y="0"/></a:lnTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:arcTo wR="50" hR="50" stAng="5400000" swAng="5400000"/><a:close/></a:path></a:pathLst></a:custGeom>' +
        `${solid('F4B183')}<a:ln w="25400"><a:solidFill><a:srgbClr val="843C0C"/></a:solidFill></a:ln></p:spPr></p:sp>`,
    },
  ],
  'gradients': [
    {
      shapes: [0, 45, 90, 135, 270]
        .map((ang, i) =>
          shape({ id: 2 + i, x: (0.4 + i * 2.4) * IN, y: 0.5 * IN, cx: 2.2 * IN, cy: 2.8 * IN,
            fill: `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="1F4E79"/></a:gs><a:gs pos="50000"><a:srgbClr val="9DC3E6"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="${ang * 60000}" scaled="0"/></a:gradFill>` }))
        .join('') +
        shape({ id: 10, x: 0.4 * IN, y: 3.8 * IN, cx: 5 * IN, cy: 3 * IN,
          fill: '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="C00000"/></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>' }) +
        shape({ id: 11, x: 6 * IN, y: 3.8 * IN, cx: 5 * IN, cy: 3 * IN, rot: 30 * 60000,
          fill: '<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="accent6"><a:lumMod val="50000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="accent6"><a:alpha val="20000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>' }),
    },
  ],
  'pictures': [
    {
      rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
      shapes: [
        ['', ''],
        ['<a:srcRect l="25000" t="10000" r="10000" b="20000"/>', ''],
        ['', '<a:alphaModFix amt="40000"/>'],
        ['', '<a:grayscl/>'],
        ['', '<a:duotone><a:prstClr val="black"/><a:srgbClr val="FFC000"/></a:duotone>'],
        ['', '<a:lum bright="20000" contrast="-30000"/>'],
      ]
        .map(([crop, fx], i) =>
          `<p:pic><p:nvPicPr><p:cNvPr id="${2 + i}" name="Picture ${i}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
          `<p:blipFill><a:blip r:embed="rIdImg">${fx}</a:blip>${crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
          `<p:spPr><a:xfrm><a:off x="${Math.round((0.4 + (i % 3) * 4) * IN)}" y="${Math.round((0.4 + Math.floor(i / 3) * 3.4) * IN)}"/><a:ext cx="${3.6 * IN}" cy="${2.7 * IN}"/></a:xfrm>` +
          `<a:prstGeom prst="${i === 1 ? 'ellipse' : 'rect'}"><a:avLst/></a:prstGeom></p:spPr></p:pic>`)
        .join(''),
    },
  ],
  // Picture editing: the XML the Picture Format tab writes.
  // Crop (`a:srcRect`, also negative for Fit), crop to shape, picture styles, transparency,
  // recolor, corrections, and alt text (`descr`, and the `adec:decorative` extension).
  'picture-editing': [
    [
      // [crop, blip effects, geometry, outline and effects, alt text, rotation]
      ['<a:srcRect l="20000" t="10000"/>', '', 'rect', '', ' descr="A gradient test image"', 0],
      ['', '', 'ellipse', '<a:ln w="88900" cap="sq"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:miter lim="800000"/></a:ln><a:effectLst><a:outerShdw blurRad="55000" dist="18000" dir="5400000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>', '', 0],
      ['', '', 'roundRect', '<a:effectLst><a:reflection blurRad="12700" stA="38000" endPos="28000" dist="5000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/></a:effectLst>', '', 0],
      ['', '<a:alphaModFix amt="50000"/>', 'rect', '', '', 0],
      ['', '<a:duotone><a:prstClr val="black"/><a:srgbClr val="D9C3A5"><a:tint val="50000"/><a:satMod val="180000"/></a:srgbClr></a:duotone>', 'rect', '', '', 0],
      ['', '<a:duotone><a:schemeClr val="accent1"><a:shade val="45000"/><a:satMod val="135000"/></a:schemeClr><a:prstClr val="white"/></a:duotone>', 'rect', '', '', 0],
    ],
    [
      ['', '<a:lum bright="20000" contrast="20000"/>', 'rect', '', '', 0],
      ['', '<a:biLevel thresh="50000"/>', 'rect', '', '', 0],
      ['', '', 'ellipse', '<a:effectLst><a:softEdge rad="112500"/></a:effectLst>', '', 0],
      ['<a:srcRect l="-16667" r="-16667"/>', '', 'rect', '', '><a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst></p:cNvPr', 30],
      ['<a:srcRect l="10000" t="10000" r="10000" b="10000"/>', '<a:grayscl/>', 'hexagon', '', '', 30],
    ],
  ].map((variants) => ({
    rels: [['rIdImg', R_IMAGE, '../media/image1.png']] as [string, string, string][],
    shapes: variants
      .map(([crop, fx, prst, style, alt, rot], i) => {
        // A decorative mark is a child of `p:cNvPr`; a description is an attribute.
        const cNvPr = String(alt).startsWith('>')
          ? `<p:cNvPr id="${2 + i}" name="Picture ${i}"${alt}>`
          : `<p:cNvPr id="${2 + i}" name="Picture ${i}"${alt}/>`;
        return (
          `<p:pic><p:nvPicPr>${cNvPr}<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
          `<p:blipFill><a:blip r:embed="rIdImg">${fx}</a:blip>${crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
          `<p:spPr><a:xfrm${rot ? ` rot="${Number(rot) * 60000}"` : ''}><a:off x="${Math.round((0.6 + (i % 3) * 4.2) * IN)}" y="${Math.round((0.5 + Math.floor(i / 3) * 3.4) * IN)}"/><a:ext cx="${3.2 * IN}" cy="${2.4 * IN}"/></a:xfrm>` +
          `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${style}</p:spPr></p:pic>`
        );
      })
      .join(''),
  })),
  'lines-arrows': [
    {
      shapes: ['solid', 'dot', 'dash', 'lgDash', 'dashDot', 'sysDot', 'sysDash']
        .map((dash, i) =>
          `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${2 + i}" name="Line ${i}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="${0.5 * IN}" y="${Math.round((0.5 + i * 0.8) * IN)}"/><a:ext cx="${5 * IN}" cy="0"/></a:xfrm>` +
          `<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln w="${19050 + i * 6350}" cap="${i % 2 ? 'rnd' : 'flat'}"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill><a:prstDash val="${dash}"/>` +
          `<a:headEnd type="${['none', 'oval', 'triangle', 'stealth', 'diamond', 'arrow', 'none'][i]}"/><a:tailEnd type="${['triangle', 'oval', 'arrow', 'none', 'stealth', 'diamond', 'triangle'][i]}" w="lg" len="lg"/></a:ln></p:spPr></p:cxnSp>`)
        .join('') +
        ['bentConnector3', 'curvedConnector3']
          .map((prst, i) =>
            `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${20 + i}" name="Connector ${i}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm${i ? ' flipV="1"' : ''}><a:off x="${7 * IN}" y="${Math.round((0.6 + i * 3) * IN)}"/><a:ext cx="${4 * IN}" cy="${2 * IN}"/></a:xfrm>` +
            `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom><a:ln w="28575"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:tailEnd type="triangle"/></a:ln></p:spPr></p:cxnSp>`)
          .join(''),
    },
    {
      // Line ends belong to open paths: the closed rectangle and callout box get none, the
      // callout's leader line and the open arc get both.
      shapes: ['rect', 'borderCallout1', 'arc', 'roundRect']
        .map((prst, i) =>
          shape({ id: 2 + i, x: (0.6 + (i % 2) * 6) * IN, y: (0.8 + Math.floor(i / 2) * 3.3) * IN, cx: 4 * IN, cy: 2 * IN, prst,
            fill: prst === 'arc' ? '<a:noFill/>' : solid('FFFFFF'),
            line: '<a:ln w="28575"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill><a:headEnd type="oval"/><a:tailEnd type="triangle" w="lg" len="lg"/></a:ln>' }))
        .join(''),
    },
  ],
  'effects': [
    {
      shapes:
        shape({ id: 2, x: IN, y: IN, cx: 3 * IN, cy: 2 * IN, fill: solid('FFFFFF'), effects: '<a:effectLst><a:outerShdw blurRad="152400" dist="76200" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>' }) +
        shape({ id: 3, x: 5 * IN, y: IN, cx: 3 * IN, cy: 2 * IN, prst: 'ellipse', fill: solid('FFC000'), effects: '<a:effectLst><a:glow rad="139700"><a:srgbClr val="4472C4"><a:alpha val="60000"/></a:srgbClr></a:glow></a:effectLst>' }) +
        shape({ id: 4, x: 9 * IN, y: IN, cx: 2.5 * IN, cy: 2 * IN, fill: solid('70AD47'), effects: '<a:effectLst><a:softEdge rad="127000"/></a:effectLst>' }) +
        shape({ id: 5, x: IN, y: 4 * IN, cx: 3 * IN, cy: 2 * IN, fill: solid('DEEBF7'), effects: '<a:effectLst><a:innerShdw blurRad="63500" dist="50800" dir="13500000"><a:prstClr val="black"><a:alpha val="50000"/></a:prstClr></a:innerShdw></a:effectLst>' }) +
        shape({ id: 6, x: 5 * IN, y: 4 * IN, cx: 3 * IN, cy: 2 * IN, rot: 20 * 60000, fill: solid('ED7D31'), effects: '<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="5400000" algn="t" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst>' }) +
        shape({ id: 7, x: 9 * IN, y: 4 * IN, cx: 2.5 * IN, cy: 2 * IN, fill: solid('5B9BD5'), style: '<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="3"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>', text: para(run('Style refs'), '<a:pPr algn="ctr"/>'), bodyPr: '<a:bodyPr anchor="ctr"/>' }),
    },
    {
      // Shadow placement (sx, sy, kx, algn, rotWithShape), reflections, and a text box shadow.
      shapes:
        shape({ id: 2, x: IN, y: 0.8 * IN, cx: 2.5 * IN, cy: 1.6 * IN, fill: solid('4472C4'), effects: '<a:effectLst><a:outerShdw blurRad="76200" dist="12700" dir="8100000" sy="23000" kx="800400" algn="br" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/></a:prstClr></a:outerShdw></a:effectLst>' }) +
        shape({ id: 3, x: 5 * IN, y: 0.8 * IN, cx: 2.5 * IN, cy: 1.6 * IN, prst: 'roundRect', fill: solid('A9D18E'), effects: '<a:effectLst><a:outerShdw blurRad="838200" dist="190500" dir="5400000" sx="90000" sy="90000" algn="t" rotWithShape="0"><a:prstClr val="black"><a:alpha val="35000"/></a:prstClr></a:outerShdw></a:effectLst>' }) +
        shape({ id: 4, x: 9 * IN, y: 0.8 * IN, cx: 2.5 * IN, cy: 1.6 * IN, rot: 30 * 60000, fill: solid('FFC000'), effects: '<a:effectLst><a:outerShdw blurRad="50800" dist="127000" dir="0" algn="ctr" rotWithShape="1"><a:prstClr val="black"><a:alpha val="50000"/></a:prstClr></a:outerShdw></a:effectLst>' }) +
        shape({ id: 5, x: IN, y: 3.6 * IN, cx: 2.5 * IN, cy: 1.4 * IN, fill: solid('ED7D31'), text: para(run('Tight'), '<a:pPr algn="ctr"/>'), bodyPr: '<a:bodyPr anchor="ctr"/>', effects: '<a:effectLst><a:reflection blurRad="6350" stA="50000" endA="300" endPos="55000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/></a:effectLst>' }) +
        shape({ id: 6, x: 5 * IN, y: 3.6 * IN, cx: 2.5 * IN, cy: 1.4 * IN, prst: 'ellipse', fill: solid('5B9BD5'), effects: '<a:effectLst><a:reflection blurRad="38100" stA="38000" endPos="90000" dist="101600" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/></a:effectLst>' }) +
        box(7, 9 * IN, 3.6 * IN, 3 * IN, 1 * IN, para(run('Text casts the shadow', 'sz="2800" b="1"')), { effects: '<a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="60000"/></a:srgbClr></a:outerShdw></a:effectLst>' }),
    },
  ],
  'fills-lines': [
    {
      // Path gradients, preset patterns, a tiled picture, and gradient text.
      rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
      shapes:
        shape({ id: 2, x: 0.5 * IN, y: 0.5 * IN, cx: 3 * IN, cy: 2 * IN, fill: `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FFF2CC"/></a:gs><a:gs pos="100000"><a:srgbClr val="C55A11"/></a:gs></a:gsLst><a:path path="rect"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>` }) +
        shape({ id: 3, x: 4 * IN, y: 0.5 * IN, cx: 3 * IN, cy: 2 * IN, fill: `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="DEEBF7"/></a:gs><a:gs pos="60000"><a:srgbClr val="5B9BD5"/></a:gs><a:gs pos="100000"><a:srgbClr val="1F4E79"/></a:gs></a:gsLst><a:path path="rect"><a:fillToRect l="0" t="0" r="100000" b="100000"/></a:path></a:gradFill>` }) +
        shape({ id: 4, x: 7.5 * IN, y: 0.5 * IN, cx: 3 * IN, cy: 2 * IN, prst: 'ellipse', fill: `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="E2F0D9"/></a:gs><a:gs pos="100000"><a:srgbClr val="548235"/></a:gs></a:gsLst><a:path path="shape"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>` }) +
        ['pct25', 'pct75', 'dnDiag', 'upDiag', 'cross', 'diagCross', 'smCheck', 'lgCheck', 'horzBrick', 'wave', 'dkHorz', 'dotGrid']
          .map((prst, i) =>
            shape({
              id: 10 + i,
              x: (0.5 + (i % 6) * 1.6) * IN,
              y: (3 + Math.floor(i / 6) * 1.1) * IN,
              cx: 1.4 * IN,
              cy: 0.9 * IN,
              fill: `<a:pattFill prst="${prst}"><a:fgClr><a:srgbClr val="2F5597"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>`,
              line: '<a:ln w="6350"><a:solidFill><a:srgbClr val="7F7F7F"/></a:solidFill></a:ln>',
            })
          )
          .join('') +
        shape({ id: 30, x: 0.5 * IN, y: 5.3 * IN, cx: 4.5 * IN, cy: 1.8 * IN, fill: '<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rIdImg"/><a:srcRect/><a:tile tx="0" ty="0" sx="50000" sy="50000" flip="xy" algn="ctr"/></a:blipFill>' }) +
        box(31, 5.5 * IN, 5.3 * IN, 7 * IN, 1.5 * IN, para(run('Gradient text', 'sz="6000" b="1"', '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="C00000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0070C0"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>'))),
    },
    {
      // Compound lines on closed and open paths, with a dash.
      shapes:
        ['dbl', 'thickThin', 'thinThick', 'tri']
          .map((cmpd, i) =>
            shape({
              id: 2 + i,
              x: (0.7 + i * 3.1) * IN,
              y: 0.8 * IN,
              cx: 2.5 * IN,
              cy: 2 * IN,
              fill: solid('F2F2F2'),
              line: `<a:ln w="76200" cmpd="${cmpd}"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:ln>`,
            })
          )
          .join('') +
        shape({ id: 7, x: 0.7 * IN, y: 3.6 * IN, cx: 5 * IN, cy: 0, prst: 'line', line: '<a:ln w="57150" cmpd="dbl"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>' }) +
        shape({ id: 8, x: 0.7 * IN, y: 4.6 * IN, cx: 5 * IN, cy: 0, prst: 'line', line: '<a:ln w="57150" cmpd="thickThin"><a:solidFill><a:srgbClr val="548235"/></a:solidFill></a:ln>' }) +
        shape({ id: 9, x: 7 * IN, y: 3.6 * IN, cx: 4 * IN, cy: 2.5 * IN, prst: 'roundRect', line: '<a:ln w="63500" cmpd="tri"><a:solidFill><a:srgbClr val="7030A0"/></a:solidFill><a:prstDash val="dash"/></a:ln>' }),
    },
  ],
  'groups-transforms': [
    {
      shapes:
        '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2" name="Group 1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="900000"><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="2743200"/><a:chOff x="0" y="0"/><a:chExt cx="2286000" cy="1371600"/></a:xfrm></p:grpSpPr>' +
        shape({ id: 3, x: 0, y: 0, cx: 1143000, cy: 685800, fill: solid('4472C4'), text: para(run('in group', 'sz="1200"')) }) +
        shape({ id: 4, x: 1143000, y: 685800, cx: 1143000, cy: 685800, prst: 'ellipse', fill: solid('ED7D31') }) +
        '</p:grpSp>' +
        shape({ id: 5, x: 6.5 * IN, y: 0.8 * IN, cx: 3 * IN, cy: 1.2 * IN, rot: 45 * 60000, fill: solid('A5A5A5'), text: para(run('Rotated 45°')) }) +
        shape({ id: 6, x: 6.5 * IN, y: 3.5 * IN, cx: 2.5 * IN, cy: 1.5 * IN, prst: 'rightArrow', flipH: true, fill: solid('70AD47'), text: para(run('flipH keeps text')) }) +
        shape({ id: 7, x: 9.5 * IN, y: 3.5 * IN, cx: 2.5 * IN, cy: 1.5 * IN, prst: 'triangle', flipV: true, fill: solid('FFC000'), text: para(run('flipV')) }),
    },
    {
      // Text in flipped groups is never mirrored: flipH keeps it upright, flipV turns it 180°,
      // and a child's own flip combines with its group's.
      shapes: [
        ['flipH="1"', 0, 'flipH group'],
        ['flipV="1"', 0, 'flipV group'],
        ['flipH="1" rot="1800000"', 0, 'flipH group, 30°'],
        ['flipV="1"', 1, 'flipV group, flipV child'],
      ]
        .map(([attrs, childFlipV, label], i) => {
          const x = Math.round((0.5 + (i % 2) * 6.2) * IN);
          const y = Math.round((0.6 + Math.floor(i / 2) * 3.4) * IN);
          return (
            `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${10 + i * 3}" name="Flipped group ${i + 1}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
            `<p:grpSpPr><a:xfrm ${attrs}><a:off x="${x}" y="${y}"/><a:ext cx="${5.4 * IN}" cy="${2.4 * IN}"/><a:chOff x="0" y="0"/><a:chExt cx="${5.4 * IN}" cy="${2.4 * IN}"/></a:xfrm></p:grpSpPr>` +
            shape({ id: 11 + i * 3, x: 0, y: 0, cx: 3 * IN, cy: 1.1 * IN, flipV: childFlipV === 1, fill: solid('4472C4'), text: para(run(String(label), 'sz="1800"')), bodyPr: '<a:bodyPr anchor="t"><a:noAutofit/></a:bodyPr>' }) +
            shape({ id: 12 + i * 3, x: 2.4 * IN, y: 1.2 * IN, cx: 3 * IN, cy: 1.2 * IN, prst: 'rightArrow', fill: solid('ED7D31'), text: para(run('Arrow text', 'sz="1600"')), bodyPr: '<a:bodyPr anchor="ctr"><a:noAutofit/></a:bodyPr>' }) +
            '</p:grpSp>'
          );
        })
        .join(''),
    },
  ],
  'placeholders-background': [
    {
      background: '<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="accent1"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="bg1"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>',
      shapes:
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>' +
        para(run('Title inherits from layout and master')) + '</p:txBody></p:sp>' +
        '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr><a:normAutofit/></a:bodyPr><a:lstStyle/>' +
        para(run('Body placeholder with master bullets')) + para(run('Level two'), '<a:pPr lvl="1"/>') + '</p:txBody></p:sp>',
    },
  ],
  'vertical-east-asian': [
    {
      shapes:
        shape({ id: 2, x: 0.6 * IN, y: 0.5 * IN, cx: 1.4 * IN, cy: 6 * IN, fill: solid('FDF2E9'),
          bodyPr: '<a:bodyPr vert="eaVert" wrap="square"><a:noAutofit/></a:bodyPr>',
          text: para(run('縦書きのテキストです。ABC 123', 'sz="2800"', '<a:ea typeface="Noto Sans CJK JP"/>')) }) +
        shape({ id: 3, x: 2.4 * IN, y: 0.5 * IN, cx: 1.4 * IN, cy: 6 * IN, fill: solid('E3F2FD'),
          bodyPr: '<a:bodyPr vert="vert" wrap="square"><a:noAutofit/></a:bodyPr>',
          text: para(run('横に回転 vert', 'sz="2800"', '<a:ea typeface="Noto Sans CJK JP"/>')) }) +
        shape({ id: 4, x: 4.2 * IN, y: 0.5 * IN, cx: 1.2 * IN, cy: 6 * IN, fill: solid('F3E5F5'),
          bodyPr: '<a:bodyPr vert="wordArtVert" wrap="square"><a:noAutofit/></a:bodyPr>',
          text: para(run('STACK', 'sz="3200" b="1"')) }),
    },
  ],
  'open-font-families': [
    {
      shapes: [
        ['Liberation Sans', 'Liberation Sans: packaged open sans-serif face'],
        ['Liberation Sans', 'Liberation Sans: another line in the open face'],
        ['Liberation Serif', 'Liberation Serif: packaged open serif face'],
        ['Liberation Serif', 'Liberation Serif: another line in the open face'],
      ]
        .map(([face, label], i) =>
          box(2 + i, 0.6 * IN, (0.6 + i * 1.5) * IN, 11 * IN, 1.2 * IN,
            para(run(`${label}. ${LOREM}`, 'sz="2000"', `<a:latin typeface="${face}"/>`))))
        .join(''),
    },
  ],
  // Shapes on the edges of an ISO A4 slide, whose height is not whole points.
  'slide-size': [
    {
      shapes:
        edgeShapes(A4) +
        shape({ id: 5, x: 3 * IN, y: 2 * IN, cx: 4 * IN, cy: 4 * IN, prst: 'ellipse', fill: solid('5B9BD5'),
          line: '<a:ln w="19050"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:ln>' }) +
        shape({ id: 6, x: 0.6 * IN, y: A4.height - 1.35 * IN, cx: 2 * IN, cy: 0.8 * IN, prst: 'roundRect', fill: solid('C00000') }),
    },
  ],
  // The same edges on a slide whose width the reference PDF rounds down.
  'slide-size-rounding': [
    {
      shapes:
        edgeShapes(ROUNDING) +
        shape({ id: 5, x: 1.5 * IN, y: 1 * IN, cx: 2 * IN, cy: 2 * IN, prst: 'ellipse', fill: solid('5B9BD5'),
          line: '<a:ln w="19050"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:ln>' }),
    },
  ],
  'preserve-unsupported': [
    {
      shapes:
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table 1"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="1483360"/></p:xfrm>' +
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid><a:gridCol w="2286000"/><a:gridCol w="2286000"/></a:tblGrid>' +
        ['A', 'B'].map((r) => `<a:tr h="370840">${[1, 2].map((c) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${r}${c}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`).join('')}</a:tr>`).join('') +
        '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>' +
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">' +
        shape({ id: 3, x: 6 * IN, y: IN, cx: 3 * IN, cy: 1 * IN, fill: solid('C00000'), text: para(run('Choice branch')) }) +
        '</mc:Choice><mc:Fallback>' + shape({ id: 3, x: 6 * IN, y: IN, cx: 3 * IN, cy: 1 * IN, fill: solid('70AD47'), text: para(run('Fallback branch')) }) + '</mc:Fallback></mc:AlternateContent>',
    },
  ],
};

mkdirSync(OUT, { recursive: true });
const image = testPng();
for (const [name, slides] of Object.entries(fixtures)) {
  // Media only where a slide references it: an orphan part is not worth a repair prompt.
  const media = name === 'pictures' || name === 'picture-editing' || name === 'fills-lines' ? [{ name: 'image1.png', bytes: image, contentType: 'image/png' }] : [];
  const tableStyles = tableStylesFor[name];
  const effectStyles = name === 'three-d-turned-effects' ? THREE_D_EFFECT_STYLES : undefined;
  const bytes = buildPptx({
    slides,
    media,
    ...(tableStyles ? { tableStyles } : {}),
    ...(effectStyles ? { effectStyles } : {}),
    ...slideSizeFor[name],
  });
  writeFileSync(`${OUT}/${name}.pptx`, bytes);
}
// Sections, comments, a chart, and embedded objects: the slide add, duplicate, and delete checks.
writeFileSync(`${OUT}/slide-structure.pptx`, buildPptx(slideStructureDeck(image, solidPng())));
// Layouts, a hidden slide, sections, prompts, and formatted notes: the slide and layout checks.
writeFileSync(`${OUT}/slide-layouts.pptx`, buildPptx(slideLayoutsDeck()));
// One slide per fill rule.
writeFileSync(`${OUT}/fill-fidelity.pptx`, buildPptx(fillFidelityDeck(image)));
// Picture bullets at several sizes, levels, and alignments.
writeFileSync(`${OUT}/picture-bullets.pptx`, buildPptx(pictureBulletsDeck()));
// East Asian faces from the theme script fonts, by each run's language.
writeFileSync(`${OUT}/east-asian-theme-fonts.pptx`, buildPptx(eastAsianThemeFontsDeck()));
// Custom kinsoku lists, and justified and distributed East Asian text.
writeFileSync(`${OUT}/east-asian-typography.pptx`, buildPptx(eastAsianTypographyDeck()));
// One chart type per slide, drawn from the cached values (ECMA-376 Part 1, 21.2).
writeFileSync(`${OUT}/charts.pptx`, buildPptx(chartsDeck()));
// Legend-bearing charts at several text and frame sizes.
writeFileSync(`${OUT}/chart-sizes.pptx`, buildPptx(chartSizesDeck()));
// Charts with embedded workbooks, for the Edit Data dialog and `setChartData`.
writeFileSync(`${OUT}/chart-data.pptx`, buildPptx(chartDataDeck()));
// Diagrams with their saved drawings (ECMA-376 Part 1, 21.4).
writeFileSync(`${OUT}/diagrams.pptx`, buildPptx(diagramsDeck()));
// A process and a list diagram, for the diagram text and structure edits.
writeFileSync(`${OUT}/diagram-editing.pptx`, buildPptx(diagramEditingDeck()));
// Links, kept content, an SVG picture, and tab stops: the round-trip checks of the feature edits.
const roundTrip = buildRoundTripDecks(image);
for (const [name, bytes] of Object.entries(roundTrip)) writeFileSync(`${OUT}/${name}.pptx`, bytes);
console.log(`${Object.keys(fixtures).length + 8 + Object.keys(roundTrip).length} fixtures written to ${OUT}/`);
