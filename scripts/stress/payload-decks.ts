// Synthetic decks with adversarial payloads: pictures that decode to huge bitmaps, SVG and
// metafile pictures that expand or loop, effects whose paint cost grows with a radius, charts
// with huge data, and package parts that try to exhaust the reader. `HANG_DECKS` includes them,
// so `hang-check.ts` and `e2e/hang.bench.ts` open each one with a time budget.

import { buildPptx, para, run, shape, solid } from '../fixtures/builder.ts';
import { IN, LOREM, R_NS, box, one, patched, range, SLIDE, type HangDeck } from './deck-parts.ts';
import {
  bitPng,
  flatJpeg,
  manyFrameGif,
  emfRecord,
  emf,
  emfLong,
  emfBigBitmap,
  svgUseBomb,
  svgUseCycle,
  svgDataUriBomb,
  svgMarkerBomb,
  svgDuplicateIdBomb,
  svgNestedMaskBomb,
  svgDefsUseBomb,
  svgSmilHrefBomb,
  svgOpacityOverdraw,
  svgUseHugePath,
  svgFilterMorphology,
  svgBlurHuge,
  svgDeep,
  svgExternal,
  svgHugeSize,
  svgEntities,
  svgEvenOddCrossing,
  svgRoundRects,
  pic,
  pictureDeck,
  png,
  fullPicture,
  svgPictureDeck,
  metafileDeck,
  stored,
  effectShape,
  MAX_RAD,
  clr,
  ser,
  chartXml,
  lineChart,
  chartDeck,
  NOFIT,
  MIXED,
  manyMastersDeck,
  externalLinksDeck,
  mediaDeck,
  nestedAlternateContent,
} from './payload-parts.ts';

// --- The decks ---------------------------------------------------------------------------------

/** The adversarial payload decks. */
export function payloadDecks(): HangDeck[] {
  return [
    // Pictures.
    {
      id: 'png-decodes-100mp',
      target: 'a 10000x10000 PNG of 30 KB (decodes to 400 MB)',
      bytes: () => fullPicture(png('image0.png', bitPng(10_000, 10_000))),
    },
    {
      id: 'png-decodes-900mp',
      target: 'a 30000x30000 PNG of 140 KB (decodes to 3.6 GB)',
      bytes: () => fullPicture(png('image0.png', bitPng(30_000, 30_000))),
    },
    {
      id: 'jpeg-decodes-100mp',
      target: 'a 10000x10000 JPEG of 400 KB',
      bytes: () =>
        stored(
          fullPicture({
            name: 'image0.jpeg',
            bytes: flatJpeg(10_000, 10_000),
            contentType: 'image/jpeg',
          })
        ),
    },
    {
      id: 'png-200-large',
      target: '200 distinct 4000x4000 PNG pictures on one slide (12.8 GB decoded)',
      bytes: () =>
        pictureDeck(
          range(200)
            .map((i) =>
              pic(
                i + 2,
                `rIdM${i}`,
                (i % 20) * 0.6 * IN,
                Math.floor(i / 20) * 0.65 * IN,
                0.6 * IN,
                0.65 * IN
              )
            )
            .join(''),
          range(200).map((i) => png(`image${i}.png`, bitPng(4000, 4000, i)))
        ),
    },
    {
      id: 'png-12-at-100mp',
      target: '12 distinct 10000x10000 PNG pictures on one slide (4.8 GB decoded)',
      bytes: () =>
        pictureDeck(
          range(12)
            .map((i) =>
              pic(
                i + 2,
                `rIdM${i}`,
                (i % 4) * 3 * IN,
                Math.floor(i / 4) * 2.25 * IN,
                3 * IN,
                2.25 * IN
              )
            )
            .join(''),
          range(12).map((i) => png(`image${i}.png`, bitPng(10_000, 10_000, i)))
        ),
    },
    {
      id: 'png-same-1000-refs',
      target: 'one 8000x8000 PNG drawn 1000 times on one slide',
      bytes: () =>
        pictureDeck(
          range(1000)
            .map((i) =>
              pic(
                i + 2,
                'rIdM0',
                (i % 40) * 0.3 * IN,
                Math.floor(i / 40) * 0.27 * IN,
                0.3 * IN,
                0.27 * IN
              )
            )
            .join(''),
          [png('image0.png', bitPng(8000, 8000))]
        ),
    },
    {
      id: 'png-damaged-data',
      target: 'a 4000x4000 PNG whose image data is random bytes',
      bytes: () =>
        fullPicture(
          png(
            'image0.png',
            bitPng(
              4000,
              4000,
              0,
              Uint8Array.from(range(50_000), (i) => (i * 7919) & 0xff)
            )
          )
        ),
    },
    {
      id: 'gif-20000-frames',
      target: 'a looping GIF of 20000 frames with no delay',
      bytes: () =>
        fullPicture({
          name: 'image0.gif',
          bytes: manyFrameGif(4000, 4000, 20_000),
          contentType: 'image/gif',
        }),
    },
    {
      id: 'picture-crop-extreme',
      target: 'a picture crop of -2^31 and 2^31 thousandths',
      bytes: () =>
        fullPicture(
          png('image0.png', bitPng(4000, 4000)),
          '',
          `<a:srcRect l="-2147483648" t="2147483647" r="-2147483648" b="2147483647"/><a:stretch><a:fillRect l="2147483647" t="-2147483648"/></a:stretch>`
        ),
    },
    {
      id: 'picture-recolor-huge',
      target: 'duotone, grayscale, alpha, and luminance on a 10000x10000 picture',
      bytes: () =>
        fullPicture(
          png('image0.png', bitPng(10_000, 10_000)),
          '<a:alphaModFix amt="50000"/><a:biLevel thresh="50000"/><a:grayscl/><a:lum bright="70000" contrast="-70000"/>' +
            '<a:duotone><a:prstClr val="black"/><a:srgbClr val="FF0000"/></a:duotone>'
        ),
    },
    {
      id: 'picture-tile-huge',
      target: 'a 10000x10000 picture tiled at 1% scale',
      bytes: () =>
        fullPicture(
          png('image0.png', bitPng(10_000, 10_000)),
          '',
          '<a:tile tx="0" ty="0" sx="1000" sy="1000" flip="xy" algn="tl"/>'
        ),
    },

    // SVG pictures.
    {
      id: 'svg-use-bomb',
      target: 'an SVG picture whose <use> chain expands to 10^10 shapes',
      bytes: () => svgPictureDeck(svgUseBomb()),
    },
    {
      id: 'svg-data-uri-bomb',
      target: 'an SVG picture with the <use> bomb as a nested SVG data: URI in <image> and feImage',
      bytes: () => svgPictureDeck(svgDataUriBomb()),
    },
    {
      id: 'svg-marker-bomb',
      target: 'an SVG picture with a 1000-element marker at 100000 vertices, and nested markers',
      bytes: () => svgPictureDeck(svgMarkerBomb()),
    },
    {
      id: 'svg-duplicate-id-bomb',
      target: 'an SVG picture with the <use> bomb and an empty element that repeats each id',
      bytes: () => svgPictureDeck(svgDuplicateIdBomb()),
    },
    {
      id: 'svg-nested-mask-bomb',
      target: 'an SVG picture with ten levels of ten masked rectangles (10^10)',
      bytes: () => svgPictureDeck(svgNestedMaskBomb()),
    },
    {
      id: 'svg-defs-use-bomb',
      target: 'an SVG picture with a <use> chain of 10^9 copies in <defs> only',
      bytes: () => svgPictureDeck(svgDefsUseBomb()),
    },
    {
      id: 'svg-smil-href-bomb',
      target: 'an SVG picture whose animation sets a <use> link to a 10^9 chain',
      bytes: () => svgPictureDeck(svgSmilHrefBomb()),
    },
    {
      id: 'svg-opacity-overdraw',
      target: 'an SVG picture with 20000 copies of full-size rectangles drawn with opacity',
      bytes: () => svgPictureDeck(svgOpacityOverdraw()),
    },
    {
      id: 'svg-use-huge-path',
      target: 'an SVG picture with 1000 copies of a path of 100000 full-size segments',
      bytes: () => svgPictureDeck(svgUseHugePath()),
    },
    {
      id: 'svg-filter-morphology',
      target: 'an SVG picture with a morphology filter of radius 500 on 20 rectangles',
      bytes: () => svgPictureDeck(svgFilterMorphology()),
    },
    {
      id: 'svg-use-cycle',
      target: 'an SVG picture with a <use> of its own ancestor',
      bytes: () => svgPictureDeck(svgUseCycle()),
    },
    {
      id: 'svg-filters-huge',
      target: 'an SVG picture 20000 px square with 200 huge blur filters',
      bytes: () => svgPictureDeck(svgBlurHuge()),
    },
    {
      id: 'svg-nesting-100k',
      target: 'an SVG picture with 100000 nested groups',
      bytes: () => svgPictureDeck(svgDeep(100_000)),
    },
    {
      id: 'svg-external-refs',
      target: 'an SVG picture with external images, styles, frames, and a script',
      bytes: () => svgPictureDeck(svgExternal()),
    },
    {
      id: 'svg-size-huge',
      target: 'an SVG picture 10^9 px square',
      bytes: () => svgPictureDeck(svgHugeSize()),
    },
    {
      id: 'svg-entity-bomb',
      target: 'an SVG picture with nested entity expansion',
      bytes: () => svgPictureDeck(svgEntities()),
    },
    {
      id: 'svg-evenodd-crossing',
      target: 'an SVG picture with a self-crossing path of 5000 segments and the even-odd rule',
      bytes: () => svgPictureDeck(svgEvenOddCrossing()),
    },
    {
      id: 'svg-round-rects',
      target: 'an SVG picture with 2000 full-size stroked rounded rectangles',
      bytes: () => svgPictureDeck(svgRoundRects()),
    },

    // Metafiles.
    {
      id: 'emf-records-1m',
      target: 'an EMF of 1000000 rectangle records',
      bytes: () => metafileDeck(emfLong(1_000_000)),
    },
    {
      id: 'emf-bounds-huge',
      target: 'an EMF that claims 2^31 by 2^31 pixels',
      bytes: () => metafileDeck(emf([emfRecord(43, [0, 0, 100, 100])], 0x7ffffff0, 0x7ffffff0)),
    },
    {
      id: 'emf-bitmap-huge',
      target: 'an EMF bitmap record of 30000x30000 with 16 bytes of pixels',
      bytes: () => metafileDeck(emfBigBitmap(30_000, 30_000)),
    },
    {
      id: 'emf-100-pictures',
      target: '100 distinct EMF pictures of 10000 records each on one slide',
      bytes: () =>
        pictureDeck(
          range(100)
            .map((i) =>
              pic(
                i + 2,
                `rIdM${i}`,
                (i % 10) * 1.2 * IN,
                Math.floor(i / 10) * 0.65 * IN,
                1.2 * IN,
                0.65 * IN
              )
            )
            .join(''),
          range(100).map((i) => {
            const b = emfLong(10_000);
            b[88 + 16] = i; // a distinct brush color per picture
            return { name: `image${i}.emf`, bytes: b, contentType: 'image/x-emf' };
          })
        ),
    },

    // Effects.
    {
      id: 'effect-blur-max',
      target: 'shadow, glow, soft edge, and reflection with radius 2^31 EMU',
      bytes: () =>
        one(
          effectShape(
            2,
            `<a:glow rad="${MAX_RAD}">${clr}</a:glow><a:innerShdw blurRad="${MAX_RAD}" dist="${MAX_RAD}">${clr}</a:innerShdw><a:outerShdw blurRad="${MAX_RAD}" dist="${MAX_RAD}" dir="2700000">${clr}</a:outerShdw><a:reflection blurRad="${MAX_RAD}" stA="50000" endPos="100000" dist="${MAX_RAD}" dir="5400000" sy="-100000" algn="bl"/><a:softEdge rad="${MAX_RAD}"/>`
          )
        ),
    },
    {
      id: 'effect-glow-large',
      target: 'a 2-inch glow and soft edge on a full-slide shape',
      bytes: () =>
        one(
          effectShape(2, `<a:glow rad="${2 * IN}">${clr}</a:glow><a:softEdge rad="${2 * IN}"/>`, {
            x: 0,
            y: 0,
            cx: 12 * IN,
            cy: 6.75 * IN,
          })
        ),
    },
    {
      id: 'effect-500-shadows',
      target: '500 shapes, each with a blurred shadow, glow, and reflection',
      bytes: () =>
        one(
          range(500)
            .map((i) =>
              effectShape(
                i + 2,
                `<a:glow rad="${IN / 4}">${clr}</a:glow><a:outerShdw blurRad="${IN / 2}" dist="${IN / 8}" dir="2700000">${clr}</a:outerShdw><a:reflection blurRad="${IN / 4}" stA="50000" endPos="90000" dist="0" dir="5400000" sy="-100000" algn="bl"/>`,
                {
                  x: (i % 25) * 0.45 * IN,
                  y: Math.floor(i / 25) * 0.32 * IN,
                  cx: 0.4 * IN,
                  cy: 0.3 * IN,
                }
              )
            )
            .join('')
        ),
    },
    {
      id: 'gradient-10000-stops',
      target: 'a gradient fill with 10000 stops on 50 shapes',
      bytes: () => {
        const stops = range(10_000)
          .map(
            (i) =>
              `<a:gs pos="${Math.round((i * 100_000) / 10_000)}"><a:srgbClr val="${((i * 2654435761) >>> 8).toString(16).padStart(6, '0').slice(0, 6)}"/></a:gs>`
          )
          .join('');
        const fill = `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`;
        return one(
          range(50)
            .map((i) =>
              shape({
                id: i + 2,
                x: (i % 10) * IN,
                y: Math.floor(i / 10) * IN,
                cx: IN,
                cy: IN,
                fill,
              })
            )
            .join('')
        );
      },
    },
    {
      id: 'pattern-line-huge',
      target: 'a pattern fill and a 2^31 EMU line on a shape 2^31 EMU wide',
      bytes: () =>
        one(
          shape({
            id: 2,
            x: 0,
            y: 0,
            cx: MAX_RAD,
            cy: MAX_RAD,
            fill: '<a:pattFill prst="smCheck"><a:fgClr><a:srgbClr val="000000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>',
            line: `<a:ln w="${MAX_RAD}" cmpd="tri"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="sysDot"/><a:headEnd type="arrow" w="lg" len="lg"/><a:tailEnd type="stealth" w="lg" len="lg"/></a:ln>`,
          })
        ),
    },
    {
      id: 'custdash-10000',
      target: 'a custom dash of 10000 segments on 10 shapes',
      bytes: () => {
        const ds = range(10_000)
          .map((i) => `<a:ds d="${1 + (i % 3)}" sp="1"/>`)
          .join('');
        const line = `<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:custDash>${ds}</a:custDash></a:ln>`;
        return one(
          range(10)
            .map((i) =>
              shape({ id: i + 2, x: i * IN, y: IN, cx: IN, cy: 4 * IN, prst: 'ellipse', line })
            )
            .join('')
        );
      },
    },
    {
      id: 'line-ends-tiny-lines',
      target: '2000 one-EMU lines with large arrowheads and thick strokes',
      bytes: () =>
        one(
          range(2000)
            .map((i) =>
              shape({
                id: i + 2,
                x: (i % 50) * 0.24 * IN,
                y: Math.floor(i / 50) * 0.16 * IN,
                cx: 1,
                cy: 0,
                prst: 'line',
                line: `<a:ln w="508000"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:headEnd type="triangle" w="lg" len="lg"/><a:tailEnd type="oval" w="lg" len="lg"/></a:ln>`,
              })
            )
            .join('')
        ),
    },
    {
      id: 'group-scale-extreme',
      target: 'groups that scale their children by 1e-6 and 1e6, rotated and flipped',
      bytes: () => {
        const grp = (id: number, ext: number, ch: number, inner: string) =>
          `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="G${id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
          `<p:grpSpPr><a:xfrm rot="2700000" flipH="1" flipV="1"><a:off x="${IN}" y="${IN}"/><a:ext cx="${ext}" cy="${ext}"/><a:chOff x="0" y="0"/><a:chExt cx="${ch}" cy="${ch}"/></a:xfrm></p:grpSpPr>${inner}</p:grpSp>`;
        const leaf = (id: number, size: number) =>
          shape({
            id,
            x: 0,
            y: 0,
            cx: size,
            cy: size,
            fill: solid('70AD47'),
            line: `<a:ln w="${MAX_RAD}">${solid('000000')}</a:ln>`,
            text: para(run('scaled')),
          });
        return one(
          grp(2, 4 * IN, 4_000_000_000_000, leaf(3, 4_000_000_000_000)) +
            grp(4, 4 * IN, 1, leaf(5, 1)) +
            grp(6, 4 * IN, 4, grp(7, MAX_RAD, 1, leaf(8, 1)))
        );
      },
    },
    {
      id: 'bevel-3d-500',
      target: '500 shapes with 3D bevels, extrusion, and camera rotation',
      bytes: () =>
        one(
          range(500)
            .map((i) =>
              shape({
                id: i + 2,
                x: (i % 25) * 0.45 * IN,
                y: Math.floor(i / 25) * 0.32 * IN,
                cx: 0.4 * IN,
                cy: 0.3 * IN,
                fill: solid('4472C4'),
                effects:
                  '<a:scene3d><a:camera prst="perspectiveContrastingRightFacing" fov="10800000"><a:rot lat="2000000" lon="19000000" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d>' +
                  `<a:sp3d z="${MAX_RAD}" extrusionH="${MAX_RAD}" contourW="${MAX_RAD}"><a:bevelT w="${MAX_RAD}" h="${MAX_RAD}" prst="artDeco"/><a:bevelB w="${MAX_RAD}" h="${MAX_RAD}"/></a:sp3d>`,
              })
            )
            .join('')
        ),
    },

    // Charts.
    {
      id: 'chart-100k-points',
      target: 'a line chart of 100000 points with 100000 categories',
      bytes: () =>
        chartDeck([
          chartXml(
            lineChart(
              ser(
                0,
                range(100_000).map((i) => Math.sin(i / 50) * 100),
                range(100_000).map((i) => `c${i}`)
              )
            )
          ),
        ]),
    },
    {
      id: 'chart-2000-series',
      target: 'a column chart of 2000 series with a legend',
      bytes: () =>
        chartDeck([
          chartXml(
            lineChart(
              range(2000)
                .map((i) =>
                  ser(i, [i, i + 1, i + 2], ['a', 'b', 'c'], `<c:tx><c:v>Series ${i}</c:v></c:tx>`)
                )
                .join(''),
              undefined,
              'barChart'
            )
          ),
        ]),
    },
    {
      id: 'chart-log-axis-bad',
      target: 'a log axis with zero, negative, NaN, and Infinity values and base 1',
      bytes: () =>
        chartDeck(
          [
            chartXml(
              lineChart(
                ser(
                  0,
                  [0, -5, 'NaN', 'Infinity', '-Infinity', '1e308', '1e-308', 5],
                  ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h']
                ),
                '<c:logBase val="1"/><c:orientation val="minMax"/><c:max val="1e308"/><c:min val="-1e308"/>'
              )
            ),
            chartXml(
              lineChart(
                ser(0, [1, 2, 3]),
                '<c:logBase val="1.0000001"/><c:orientation val="minMax"/><c:max val="1e308"/><c:min val="1e-308"/>'
              )
            ),
          ],
          2
        ),
    },
    {
      id: 'chart-axis-units-tiny',
      target: 'a value axis from -1e308 to 1e308 with a major unit of 1e-300',
      bytes: () =>
        chartDeck([
          chartXml(
            lineChart(
              ser(0, [1, 2, 3]),
              '<c:orientation val="minMax"/><c:max val="1e308"/><c:min val="-1e308"/>'
            )
              .replace(
                '<c:crossAx val="1"/></c:valAx>',
                '<c:crossAx val="1"/><c:majorUnit val="1e-300"/><c:minorUnit val="1e-300"/></c:valAx>'
              )
              .replace('<c:majorGridlines/>', '<c:majorGridlines/><c:minorGridlines/>')
          ),
        ]),
    },
    {
      id: 'chart-labels-huge',
      target: 'chart category labels of 100000 characters and data labels on 5000 points',
      bytes: () =>
        chartDeck([
          chartXml(
            lineChart(
              ser(
                0,
                range(5000).map((i) => i),
                range(5000).map((i) => (i < 3 ? LOREM.repeat(1150) : `c${i}`)),
                '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="1"/><c:showSerName val="1"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
              )
            )
          ),
        ]),
    },
    {
      id: 'chart-pt-idx-huge',
      target: 'chart caches with ptCount and point indexes of 2^31',
      bytes: () =>
        chartDeck([
          chartXml(
            lineChart(
              `<c:ser><c:idx val="2147483647"/><c:order val="2147483647"/><c:cat><c:strRef><c:f>A</c:f><c:strCache><c:ptCount val="2147483647"/><c:pt idx="2147483646"><c:v>x</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
                `<c:val><c:numRef><c:f>B</c:f><c:numCache><c:ptCount val="2147483647"/><c:pt idx="2147483646"><c:v>1</c:v></c:pt><c:pt idx="0"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>`
            )
          ),
        ]),
    },
    {
      id: 'chart-pie-50000',
      target: 'a pie chart of 50000 slices with varied colors and labels',
      bytes: () =>
        chartDeck([
          chartXml(
            `<c:pieChart><c:varyColors val="1"/>${ser(
              0,
              range(50_000).map((i) => 1 + (i % 7)),
              range(50_000).map((i) => `s${i}`),
              '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/></c:dLbls>'
            )}<c:firstSliceAng val="0"/></c:pieChart>`
          ),
        ]),
    },
    {
      id: 'chart-200-frames',
      target: '200 chart frames on one slide, each with 2000 points',
      bytes: () =>
        chartDeck(
          [
            chartXml(
              lineChart(
                ser(
                  0,
                  range(2000).map((i) => i % 97)
                )
              ),
              false
            ),
          ],
          200
        ),
    },
    {
      id: 'chart-part-missing-cyclic',
      target: 'chart frames to a missing part, to the slide itself, and to a non-chart part',
      bytes: () =>
        chartDeck(
          ['', '', ''],
          3,
          (i) => ['../charts/none.xml', '../slides/slide1.xml', '../theme/theme1.xml'][i]!
        ),
    },

    // Text.
    {
      id: 'text-mixed-scripts-zwj',
      target: 'one run of 200000 characters of mixed scripts and emoji ZWJ sequences',
      bytes: () => one(box(2, para(run(MIXED.repeat(3500), 'sz="800"')), { bodyPr: NOFIT })),
    },
    {
      id: 'text-vertical-huge',
      target: 'vert, eaVert, and wordArtVert text of 100000 characters each',
      bytes: () =>
        one(
          ['vert', 'eaVert', 'wordArtVert', 'vert270', 'mongolianVert']
            .map((v, i) =>
              box(i + 2, para(run(LOREM.repeat(1150), 'sz="1000"')), {
                x: i * 2.4 * IN,
                cx: 2.3 * IN,
                cy: 6 * IN,
                bodyPr: `<a:bodyPr vert="${v}" wrap="square"><a:normAutofit/></a:bodyPr>`,
              })
            )
            .join('')
        ),
    },
    {
      id: 'text-warp-huge',
      target: 'WordArt warps of 50000 characters each',
      bytes: () =>
        one(
          ['textArchUp', 'textCircle', 'textWave4', 'textInflate', 'textCascadeUp']
            .map((w, i) =>
              box(i + 2, para(run(LOREM.repeat(580), 'sz="4000"')), {
                x: i * 2.4 * IN,
                cx: 2.3 * IN,
                cy: 6 * IN,
                bodyPr: `<a:bodyPr wrap="square"><a:prstTxWarp prst="${w}"><a:avLst/></a:prstTxWarp><a:normAutofit/></a:bodyPr>`,
              })
            )
            .join('')
        ),
    },
    {
      id: 'text-1-emu-box-autofit',
      target: '20 text boxes of 1x1 EMU with shrink and grow autofit',
      bytes: () =>
        one(
          range(20)
            .map((i) =>
              box(i + 2, para(run(LOREM.repeat(50), 'sz="6000"')), {
                cx: 1,
                cy: 1,
                bodyPr:
                  i % 2
                    ? '<a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"><a:normAutofit fontScale="1000" lnSpcReduction="90000"/></a:bodyPr>'
                    : '<a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr>',
              })
            )
            .join('')
        ),
    },
    {
      id: 'embedded-fonts-bad',
      target: '300 embedded font entries to damaged, empty, and huge font parts',
      bytes: () =>
        patched(one(box(2, para(run('embedded', '', '<a:latin typeface="Embedded 0"/>')))), (f) => {
          const n = 300;
          const lst = range(n)
            .map(
              (i) =>
                `<p:embeddedFont><p:font typeface="Embedded ${i}"/><p:regular r:id="rIdF${i}"/><p:bold r:id="rIdF${i}"/></p:embeddedFont>`
            )
            .join('');
          const pres = 'ppt/presentation.xml';
          f.set(
            pres,
            f
              .get(pres)
              .replace(
                '<p:defaultTextStyle',
                `<p:embeddedFontLst>${lst}</p:embeddedFontLst><p:defaultTextStyle`
              )
          );
          const rels = 'ppt/_rels/presentation.xml.rels';
          f.set(
            rels,
            f.get(rels).replace(
              '</Relationships>',
              `${range(n)
                .map(
                  (i) =>
                    `<Relationship Id="rIdF${i}" Type="${R_NS}/font" Target="fonts/font${i}.fntdata"/>`
                )
                .join('')}</Relationships>`
            )
          );
          const types = '[Content_Types].xml';
          f.set(
            types,
            f
              .get(types)
              .replace(
                '<Default ',
                '<Default Extension="fntdata" ContentType="application/x-fontdata"/><Default '
              )
          );
          for (const i of range(n)) {
            // An sfnt header that claims 65535 tables at offsets past the end, or obfuscated bytes.
            const b = new Uint8Array(i % 3 === 0 ? 64 * 1024 : 64).fill(i % 3 === 1 ? 0 : 0xff);
            b.set([0, 1, 0, 0, 0xff, 0xff], 0);
            f.put(`ppt/fonts/font${i}.fntdata`, b);
          }
        }),
    },

    // Package.
    {
      id: 'zip-50000-entries',
      target: 'a package of 50000 tiny entries',
      bytes: () =>
        patched(one(''), (f) => range(50_000).forEach((i) => f.bin(`ppt/media/x${i}.bin`))),
    },
    {
      id: 'zip-deflate-bomb',
      target: 'one XML part of 64 MB of spaces that deflates to 64 KB',
      bytes: () =>
        patched(one(''), (f) =>
          f.set(SLIDE, f.get(SLIDE).replace('</p:sld>', `${' '.repeat(64 * 1024 * 1024)}</p:sld>`))
        ),
    },
    {
      id: 'custdata-huge',
      target: 'custDataLst, comments, and extension lists of 20 MB',
      bytes: () =>
        patched(one(box(2, para(run('custom data')))), (f) => {
          const tags = range(200_000)
            .map((i) => `<p:custData r:id="rIdT${i % 5}"/>`)
            .join('');
          f.set(
            SLIDE,
            f.get(SLIDE).replace('</p:cSld>', `<p:custDataLst>${tags}</p:custDataLst></p:cSld>`)
          );
          const ext = `<p:extLst>${range(100_000)
            .map((i) => `<p:ext uri="{E${i}}"><x:y xmlns:x="urn:x">${i}</x:y></p:ext>`)
            .join('')}</p:extLst>`;
          const pres = 'ppt/presentation.xml';
          f.set(pres, f.get(pres).replace('</p:presentation>', `${ext}</p:presentation>`));
        }),
    },
    {
      id: 'alternate-content-deep',
      target: '5000 nested mc:AlternateContent levels',
      bytes: () => one(nestedAlternateContent(5000)),
    },
    {
      id: 'masters-500',
      target: '500 slide masters, each with a 40 KB theme',
      bytes: () => manyMastersDeck(500, 40_000),
    },
    {
      id: 'external-links-5000',
      target: '5000 external hyperlinks and a picture linked to a URL',
      bytes: () => externalLinksDeck(5000),
    },
    {
      id: 'media-external',
      target: 'video and audio linked to URLs, an autoplay timing, and a 5 MB clip',
      bytes: mediaDeck,
    },
    {
      id: 'media-file-100mb',
      target: 'a 100 MB media part',
      bytes: () =>
        stored(
          buildPptx({
            slides: [
              {
                shapes: box(2, para(run('big media'))),
                rels: [['rIdV', `${R_NS}/video`, '../media/media1.mp4']],
              },
            ],
            media: [
              {
                name: 'media1.mp4',
                bytes: new Uint8Array(100 * 1024 * 1024).fill(3),
                contentType: 'video/mp4',
              },
            ],
          })
        ),
      slowMs: 6000,
    },
  ];
}
