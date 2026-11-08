// Synthetic decks that try to make the editor hang: one adversarial input per deck, each
// aimed at a loop, a recursion, or a size that an untrusted file controls. `hang-check.ts`
// opens, lays out, and paints each one with a time budget.

import { zipSync, strToU8, unzipSync } from 'fflate';
import { buildPptx, cell, para, run, shape, solid, table } from '../fixtures/builder.ts';
import {
  IN,
  LOREM,
  MC_NS,
  P_NS,
  R_NS,
  blipFill,
  box,
  one,
  patched,
  placeholder,
  range,
  SLIDE,
  tableRows,
  type HangDeck,
} from './deck-parts.ts';
import { payloadDecks } from './payload-decks.ts';
import { uiScaleDecks } from './ui-scale-decks.ts';

export type { HangDeck } from './deck-parts.ts';

/** `n` slide list entries, each with slide id `id(i)`, to its own one-byte part. */
function slideIdDeck(n: number, id: (i: number) => string): Uint8Array {
  return patched(one(''), (f) => {
    for (const i of range(n)) f.bin(`ppt/media/p${i}.bin`);
    const bin = '<Default Extension="bin" ContentType="application/octet-stream"/>';
    f.set(
      '[Content_Types].xml',
      f.get('[Content_Types].xml').replace('<Default ', `${bin}<Default `)
    );
    const rels = range(n)
      .map((i) => `<Relationship Id="rIdX${i}" Type="${R_NS}/slide" Target="media/p${i}.bin"/>`)
      .join('');
    const presRels = 'ppt/_rels/presentation.xml.rels';
    f.set(presRels, f.get(presRels).replace('</Relationships>', `${rels}</Relationships>`));
    const ids = range(n)
      .map((i) => `<p:sldId id="${id(i)}" r:id="rIdX${i}"/>`)
      .join('');
    const pres = 'ppt/presentation.xml';
    f.set(pres, f.get(pres).replace('</p:sldIdLst>', `${ids}</p:sldIdLst>`));
  });
}

const NOFIT = '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>';
const NORM = '<a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr>';

/** Nested groups, `depth` levels deep, around one text shape. */
function nestedGroups(depth: number, chExt = '<a:chExt cx="100" cy="100"/>'): string {
  let inner = box(2, para(run('deep')));
  for (let i = 0; i < depth; i += 1)
    inner =
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${1000 + i}" name="G${i}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/>${chExt}</a:xfrm></p:grpSpPr>` +
      `${inner}</p:grpSp>`;
  return inner;
}

/** A custom geometry shape with the given path XML. */
function custGeom(id: number, path: string, attrs = 'w="100" h="100"'): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Cust ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${4 * IN}" cy="${3 * IN}"/></a:xfrm>` +
    `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>` +
    `<a:pathLst><a:path ${attrs}>${path}</a:path></a:pathLst></a:custGeom>${solid('4472C4')}` +
    `<a:ln w="12700">${solid('000000')}</a:ln></p:spPr></p:sp>`
  );
}

/** A deck with its slide layout's master link pointed back at a layout (a part cycle). */
function relCycleDeck(): Uint8Array {
  const files = unzipSync(one(box(2, para(run('cycle')))));
  const relsName = 'ppt/slideLayouts/_rels/slideLayout1.xml.rels';
  files[relsName] = strToU8(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideLayout1.xml"/>' +
      '</Relationships>'
  );
  return zipSync(files);
}

/** A 1×1 PNG whose header claims `w`×`h` pixels. */
function pngHeader(w: number, h: number): Uint8Array {
  const b = new Uint8Array(33);
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  const v = new DataView(b.buffer);
  v.setUint32(16, w);
  v.setUint32(20, h);
  b.set([8, 6, 0, 0, 0], 24);
  return b;
}

function pictureShape(fill: string): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Tile"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${12 * IN}" cy="${7 * IN}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fill}</p:spPr></p:sp>`
  );
}

function imageDeck(fill: string, png: Uint8Array): Uint8Array {
  return buildPptx({
    slides: [
      {
        shapes: pictureShape(fill),
        rels: [
          [
            'rIdImg',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            '../media/image1.png',
          ],
        ],
      },
    ],
    media: [{ name: 'image1.png', bytes: png, contentType: 'image/png' }],
  });
}

const manyShapes = (n: number) =>
  Array.from({ length: n }, (_, i) =>
    shape({
      id: i + 2,
      x: (i % 100) * 0.12 * IN,
      y: Math.floor(i / 100) * 0.03 * IN,
      cx: 0.1 * IN,
      cy: 0.1 * IN,
      fill: solid('4472C4'),
      text: para(run(String(i), 'sz="400"')),
    })
  ).join('');

const grid = (cols: number) => Array.from({ length: cols }, () => Math.floor((11 * IN) / cols));

/** The full set. */
export const HANG_DECKS: readonly HangDeck[] = [
  {
    id: 'tabs-many-stops',
    target: '60000 tab characters with 60000 explicit stops behind the cursor',
    bytes: () =>
      one(
        box(
          2,
          '<a:p><a:pPr><a:tabLst>' +
            '<a:tab pos="0" algn="l"/>'.repeat(60_000) +
            '</a:tabLst></a:pPr>' +
            run('\t'.repeat(60_000)) +
            '</a:p>',
          { bodyPr: NOFIT }
        )
      ),
  },
  // Size.
  {
    id: 'slides-3000',
    target: 'open and rail with 3000 slides',
    bytes: () =>
      buildPptx({
        slides: Array.from({ length: 3000 }, (_, i) => ({
          shapes: box(2, para(run(`Slide ${i}`))),
        })),
      }),
  },
  {
    id: 'shapes-10000',
    target: '10000 text shapes on one slide',
    bytes: () => one(manyShapes(10_000)),
  },
  {
    id: 'text-500k-chars',
    target: 'one paragraph of 500000 characters',
    bytes: () => one(box(2, para(run(LOREM.repeat(5_800), 'sz="1000"')), { bodyPr: NOFIT })),
  },
  {
    id: 'paragraphs-20000',
    target: '20000 paragraphs in one shape',
    bytes: () =>
      one(
        box(2, Array.from({ length: 20_000 }, (_, i) => para(run(`p${i}`))).join(''), {
          bodyPr: NOFIT,
        })
      ),
  },
  {
    id: 'runs-50000',
    target: '50000 runs in one paragraph',
    bytes: () =>
      one(
        box(2, para(Array.from({ length: 50_000 }, () => run('ab ', 'b="1"')).join('')), {
          bodyPr: NOFIT,
        })
      ),
  },
  {
    id: 'table-300x60',
    target: 'a 300 by 60 table',
    bytes: () =>
      one(
        table({
          id: 2,
          x: IN,
          y: IN,
          cx: 11 * IN,
          cy: 5 * IN,
          columns: grid(60),
          tblPr: '<a:tblPr/>',
          rows: tableRows(300, 60),
        })
      ),
  },

  // Line breaking.
  {
    id: 'word-100k-narrow',
    target: 'one 100000-character word in a narrow box',
    bytes: () => one(box(2, para(run('W'.repeat(100_000))), { cx: 0.3 * IN, bodyPr: NOFIT })),
  },
  {
    id: 'width-zero',
    target: 'wrapped text in a zero-width box',
    bytes: () => one(box(2, para(run(LOREM.repeat(20))), { cx: 0, bodyPr: NOFIT })),
  },
  {
    id: 'insets-exceed-box',
    target: 'insets wider than the box',
    bytes: () =>
      one(
        box(2, para(run(LOREM.repeat(20))), {
          bodyPr: `<a:bodyPr wrap="square" lIns="${4 * IN}" rIns="${4 * IN}"><a:noAutofit/></a:bodyPr>`,
        })
      ),
  },
  {
    id: 'combining-marks-50k',
    target: 'one grapheme cluster of 50000 combining marks',
    bytes: () => one(box(2, para(run('a' + '́'.repeat(50_000))), { bodyPr: NOFIT })),
  },
  {
    id: 'bullet-combining-marks-50k',
    target: 'a bullet grapheme of 50000 combining marks',
    bytes: () =>
      one(
        box(
          2,
          para(run('item'), `<a:pPr><a:buChar char="${'a' + '\u0301'.repeat(50_000)}"/></a:pPr>`),
          { bodyPr: NOFIT }
        )
      ),
  },
  {
    id: 'cjk-100k-no-breaks',
    target: '100000 CJK characters in kinsoku rules',
    bytes: () =>
      one(
        box(
          2,
          // Not `run()`: its `lang="en-US"` and a second `lang` are not well-formed XML.
          para(
            `<a:r><a:rPr lang="ja-JP" altLang="en-US"/><a:t>${'、「'.repeat(50_000)}</a:t></a:r>`,
            '<a:pPr eaLnBrk="1" hangingPunct="1"/>'
          ),
          { bodyPr: NOFIT }
        )
      ),
  },
  {
    id: 'bidi-alternating',
    target: '20000 alternating right-to-left and left-to-right words',
    bytes: () =>
      one(box(2, para(run('abc אבג '.repeat(10_000)), '<a:pPr rtl="1"/>'), { bodyPr: NOFIT })),
  },

  // Autofit and spacing.
  {
    id: 'normautofit-never-fits',
    target: 'shrink-on-overflow text that cannot fit',
    bytes: () =>
      one(
        box(2, para(run(LOREM.repeat(400), 'sz="4000"')), {
          cx: 0.5 * IN,
          cy: 0.1 * IN,
          bodyPr: NORM,
        })
      ),
  },
  {
    id: 'spautofit-zero-width',
    target: 'resize-to-fit text in a zero-width box',
    bytes: () => one(box(2, para(run(LOREM.repeat(50))), { cx: 0, cy: 0 })),
  },
  {
    id: 'font-size-max',
    target: 'a 4000 pt run',
    bytes: () => one(box(2, para(run(LOREM, 'sz="400000"')), { bodyPr: NOFIT })),
  },
  {
    id: 'line-spacing-zero',
    target: 'zero and negative line spacing with 5000 lines',
    bytes: () =>
      one(
        box(
          2,
          Array.from({ length: 5000 }, () =>
            para(
              run('x'),
              '<a:pPr><a:lnSpc><a:spcPct val="0"/></a:lnSpc><a:spcBef><a:spcPts val="-158400"/></a:spcBef></a:pPr>'
            )
          ).join(''),
          { bodyPr: NORM }
        )
      ),
  },
  {
    id: 'tab-default-zero',
    target: 'tabs with a zero default tab size',
    bytes: () =>
      one(box(2, para(run('a\tb\tc\t'.repeat(2000)), '<a:pPr defTabSz="0"/>'), { bodyPr: NOFIT })),
  },
  {
    id: 'tab-stops-negative',
    target: 'tabs with negative and duplicate tab stops',
    bytes: () =>
      one(
        box(
          2,
          para(
            run('a\t'.repeat(3000)),
            '<a:pPr defTabSz="-914400"><a:tabLst><a:tab pos="-100" algn="l"/><a:tab pos="0" algn="r"/><a:tab pos="0" algn="dec"/></a:tabLst></a:pPr>'
          ),
          { bodyPr: NOFIT }
        )
      ),
  },
  {
    id: 'indent-huge',
    target: 'a margin and indent wider than the box',
    bytes: () =>
      one(
        box(2, para(run(LOREM.repeat(30)), '<a:pPr marL="51206400" indent="-51206400"/>'), {
          bodyPr: NOFIT,
        })
      ),
  },
  {
    id: 'columns-1000',
    target: '16 columns with huge spacing',
    bytes: () =>
      one(
        box(2, para(run(LOREM.repeat(200))), {
          bodyPr: `<a:bodyPr wrap="square" numCol="16" spcCol="51206400"><a:normAutofit/></a:bodyPr>`,
        })
      ),
  },

  // Numbering.
  {
    id: 'roman-start-max',
    target: 'a roman numeral list that starts at 2147483647',
    bytes: () =>
      one(
        box(
          2,
          para(
            run('item'),
            '<a:pPr marL="342900" indent="-342900"><a:buAutoNum type="romanUcPeriod" startAt="2147483647"/></a:pPr>'
          ).repeat(3),
          { bodyPr: NOFIT }
        )
      ),
  },
  {
    id: 'alpha-start-max',
    target: 'a letter list that starts at 32767',
    bytes: () =>
      one(
        box(
          2,
          para(
            run('item'),
            '<a:pPr marL="342900" indent="-342900"><a:buAutoNum type="alphaLcParenR" startAt="32767"/></a:pPr>'
          ).repeat(3),
          { bodyPr: NOFIT }
        )
      ),
  },
  {
    id: 'numbered-20000',
    target: '20000 numbered paragraphs',
    bytes: () =>
      one(
        box(
          2,
          para(
            run('i'),
            '<a:pPr marL="342900" indent="-342900"><a:buAutoNum type="arabicPeriod"/></a:pPr>'
          ).repeat(20_000),
          { bodyPr: NOFIT }
        )
      ),
  },

  // Groups and geometry.
  { id: 'groups-deep-5000', target: '5000 nested groups', bytes: () => one(nestedGroups(5000)) },
  {
    id: 'groups-zero-chext',
    target: 'nested groups with a zero child extent',
    bytes: () => one(nestedGroups(50, '<a:chExt cx="0" cy="0"/>')),
  },
  {
    id: 'custgeom-200k-points',
    target: 'a custom path of 200000 points',
    bytes: () =>
      one(
        custGeom(
          2,
          '<a:moveTo><a:pt x="0" y="0"/></a:moveTo>' +
            Array.from(
              { length: 200_000 },
              (_, i) => `<a:lnTo><a:pt x="${i % 100}" y="${(i * 7) % 100}"/></a:lnTo>`
            ).join('')
        )
      ),
  },
  {
    id: 'custgeom-arc-huge',
    target: 'arcs with huge sweep and zero radius',
    bytes: () =>
      one(
        custGeom(
          2,
          '<a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:arcTo wR="50" hR="50" stAng="0" swAng="2147483647"/><a:arcTo wR="0" hR="0" stAng="0" swAng="21600000"/><a:close/>'
        )
      ),
  },
  {
    id: 'custgeom-zero-path-size',
    target: 'a custom path with a zero path size',
    bytes: () =>
      one(
        custGeom(
          2,
          '<a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="10" y="10"/></a:lnTo>',
          'w="0" h="0"'
        )
      ),
  },
  {
    id: 'guide-chain-5000',
    target: 'a chain of 5000 shape guides',
    bytes: () => {
      const gd = Array.from(
        { length: 5000 },
        (_, i) => `<a:gd name="g${i}" fmla="*/ ${i ? `g${i - 1}` : 'w'} 1 1"/>`
      ).join('');
      const geo = custGeom(
        2,
        '<a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="g4999" y="h"/></a:lnTo>'
      ).replace('<a:gdLst/>', `<a:gdLst>${gd}</a:gdLst>`);
      return one(geo);
    },
  },
  {
    id: 'guide-cycle',
    target: 'shape guides that name each other',
    bytes: () =>
      one(
        custGeom(
          2,
          '<a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="ga" y="gb"/></a:lnTo>'
        ).replace(
          '<a:gdLst/>',
          '<a:gdLst><a:gd name="ga" fmla="+- gb 1 0"/><a:gd name="gb" fmla="+- ga 1 0"/></a:gdLst>'
        )
      ),
  },
  {
    id: 'preset-adjust-extreme',
    target: 'preset shapes with extreme adjust values',
    bytes: () =>
      one(
        [
          'gear9',
          'star32',
          'wave',
          'doubleWave',
          'cloudCallout',
          'blockArc',
          'circularArrow',
          'leftRightRibbon',
        ]
          .map((prst, i) =>
            shape({
              id: i + 2,
              x: (i % 4) * 3 * IN,
              y: Math.floor(i / 4) * 3 * IN,
              cx: 2 * IN,
              cy: 2 * IN,
              prst,
              fill: solid('70AD47'),
              avLst:
                '<a:gd name="adj" fmla="val 2147483647"/><a:gd name="adj1" fmla="val -2147483647"/><a:gd name="adj2" fmla="val 2147483647"/><a:gd name="adj3" fmla="val 0"/>',
            })
          )
          .join('')
      ),
  },
  {
    id: 'extents-max',
    target: 'shapes at the largest coordinates',
    bytes: () =>
      one(
        shape({
          id: 2,
          x: 2147483647,
          y: -2147483647,
          cx: 2147483647,
          cy: 2147483647,
          rot: 2147483647,
          prst: 'star32',
          fill: solid('FF0000'),
          text: para(run(LOREM)),
        })
      ),
  },

  // Lines.
  {
    id: 'dash-zero-length',
    target: 'a custom dash with zero lengths on a long line',
    bytes: () =>
      one(
        shape({
          id: 2,
          x: 0,
          y: 0,
          cx: 12 * IN,
          cy: 6 * IN,
          prst: 'ellipse',
          line: `<a:ln w="1">${solid('000000')}<a:custDash><a:ds d="0" sp="0"/></a:custDash></a:ln>`,
        })
      ),
  },
  {
    id: 'dash-tiny-thick-line',
    target: 'a 1-EMU line with a dotted dash on a huge shape',
    bytes: () =>
      one(
        shape({
          id: 2,
          x: 0,
          y: 0,
          cx: 2147483647,
          cy: 2147483647,
          prst: 'ellipse',
          line: `<a:ln w="1">${solid('000000')}<a:prstDash val="sysDot"/></a:ln>`,
        })
      ),
  },

  // Tables.
  {
    id: 'table-tiny-dashed-borders',
    target: 'one-EMU dotted borders on a large table cell',
    bytes: () =>
      one(
        table({
          id: 2,
          x: 0,
          y: 0,
          cx: 2147483647,
          cy: 2147483647,
          columns: [2147483647],
          tblPr: '<a:tblPr/>',
          rows:
            '<a:tr h="2147483647">' +
            cell('', {
              tcPr: ['lnL', 'lnR', 'lnT', 'lnB', 'lnTlToBr', 'lnBlToTr']
                .map(
                  (side) =>
                    `<a:${side} w="1">${solid('000000')}<a:prstDash val="sysDot"/></a:${side}>`
                )
                .join(''),
            }) +
            '</a:tr>',
        })
      ),
  },
  {
    id: 'table-span-huge',
    target: 'table cells with huge grid and row spans',
    bytes: () =>
      one(
        table({
          id: 2,
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: 2 * IN,
          columns: grid(3),
          tblPr: '<a:tblPr/>',
          rows: `<a:tr h="${IN}">${cell('a', { attrs: ' gridSpan="2147483647" rowSpan="2147483647"' })}${cell('b')}${cell('c')}</a:tr><a:tr h="${IN}">${cell('d', { attrs: ' hMerge="1" vMerge="1"' })}${cell('e')}${cell('f')}</a:tr>`,
        })
      ),
  },
  {
    id: 'table-cell-text-tall',
    target: 'a table cell with 5000 lines',
    bytes: () =>
      one(
        table({
          id: 2,
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: 2 * IN,
          columns: grid(2),
          tblPr: '<a:tblPr/>',
          rows: `<a:tr h="${IN}">${cell(Array.from({ length: 5000 }, (_, i) => `line ${i}`).join('\n'))}${cell('b')}</a:tr>`,
        })
      ),
  },

  // Images.
  {
    id: 'image-tile-tiny-scale',
    target: 'a picture tile at 0.001% scale',
    bytes: () =>
      imageDeck(
        blipFill('<a:tile tx="0" ty="0" sx="1" sy="1" flip="none" algn="tl"/>'),
        pngHeader(1, 1)
      ),
  },
  {
    id: 'image-header-huge',
    target: 'a picture whose header claims 2^31 by 2^31 pixels',
    bytes: () =>
      imageDeck(
        blipFill('<a:stretch><a:fillRect/></a:stretch>'),
        pngHeader(0x7fffffff, 0x7fffffff)
      ),
  },

  // Package.
  { id: 'layout-master-cycle', target: 'a layout whose master is itself', bytes: relCycleDeck },
  {
    id: 'fonts-500-families',
    target: '500 distinct font families on one slide',
    bytes: () =>
      one(
        Array.from({ length: 500 }, (_, i) =>
          box(i + 2, para(run('x', '', `<a:latin typeface="Family ${i}"/>`)), {
            cx: IN / 4,
            cy: IN / 4,
          })
        ).join('')
      ),
  },
  {
    id: 'font-origin-stalled',
    target: 'a font origin whose download never finishes',
    stalledFonts: true,
    bytes: () => one(box(2, para(run('Stalled fonts', '', '<a:latin typeface="Open Sans"/>')))),
  },
  // Super-linear costs that fuzzing found: each took 12 to 83 s before its fix.
  {
    id: 'slide-ids-invalid-16000',
    target: '16000 slide list entries whose ids need a fresh id',
    bytes: () => slideIdDeck(16_000, (i) => `x${i}`),
  },
  {
    id: 'slide-ids-max-16000',
    target: '16000 slide list entries that repeat the largest slide id',
    bytes: () => slideIdDeck(16_000, () => '2147483647'),
  },
  {
    id: 'requires-prefixes-2000',
    target: '800 alternate contents that each require 2000 prefixes',
    bytes: () =>
      patched(one(''), (f) => {
        const decls = range(2000)
          .map((i) => `xmlns:q${i}="${P_NS}"`)
          .join(' ');
        const req = range(2000)
          .map((i) => `q${1999 - i}`)
          .join(' ');
        const ac = (id: number): string =>
          `<mc:AlternateContent xmlns:mc="${MC_NS}"><mc:Choice Requires="${req}">` +
          `${box(id, para(run('ac')), { cx: IN, cy: IN })}</mc:Choice><mc:Fallback/></mc:AlternateContent>`;
        const all = range(800)
          .map((i) => ac(100 + i))
          .join('');
        const xml = f.get(SLIDE).replace('<p:sld ', `<p:sld ${decls} `);
        f.set(SLIDE, xml.replace('</p:spTree>', `${all}</p:spTree>`));
      }),
  },
  {
    id: 'placeholders-unmatched-24000',
    target: '24000 slide placeholders that match none of 24000 layout placeholders',
    // Painting 24000 shapes is linear but slow; the quadratic match took about 12 s.
    slowMs: 6000,
    bytes: () =>
      patched(one(''), (f) => {
        const rels = f.get('ppt/slides/_rels/slide1.xml.rels');
        const layout = `ppt/slideLayouts/${/slideLayouts\/([^"]+)"/.exec(rels)![1]}`;
        const dates = range(24_000)
          .map((i) => placeholder(5000 + i, 'dt', 900_000 + i))
          .join('');
        f.set(layout, f.get(layout).replace('</p:spTree>', `${dates}</p:spTree>`));
        const titles = range(24_000)
          .map((i) => placeholder(100_000 + i, 'title', 100_000 + i))
          .join('');
        f.set(SLIDE, f.get(SLIDE).replace('</p:spTree>', `${titles}</p:spTree>`));
      }),
  },
  {
    id: 'namespace-scopes-64000',
    target: '64000 elements that each declare a prefix under 64000 inherited ones',
    bytes: () =>
      patched(one(''), (f) => {
        const decls = range(64_000)
          .map((i) => `xmlns:q${i}="urn:q${i}"`)
          .join(' ');
        const exts = range(64_000)
          .map((i) => `<p:ext uri="{U${i}}" xmlns:z="urn:z"/>`)
          .join('');
        const xml = f.get(SLIDE).replace('<p:sld ', `<p:sld ${decls} `);
        f.set(SLIDE, xml.replace('</p:sld>', `<p:extLst>${exts}</p:extLst></p:sld>`));
      }),
  },
  {
    id: 'notes-500k-chars',
    target: 'speaker notes of 500000 characters',
    bytes: () =>
      buildPptx({ slides: [{ shapes: box(2, para(run('notes'))), notes: LOREM.repeat(5_800) }] }),
  },
  // Sizes that the editor UI reads on every change: the toolbar state, the slide panel, the
  // layout menu, the selection, and the live region. Each one must open and edit in the
  // browser without a long freeze (`e2e/hang.stress.ts`).
  ...uiScaleDecks(),
  // Adversarial payloads: pictures, SVG, metafiles, effects, charts, text, and package parts.
  ...payloadDecks(),
];
