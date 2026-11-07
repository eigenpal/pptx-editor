// Probes of the SVG picture subset (svg-picture.ts): bypasses of a counting check, features
// whose cost grows per pixel, vertex, or layer, inputs that make a check slow, and SVG files
// from common tools. A refused picture draws its raster alternative.

import { expect, test } from 'bun:test';
import { MAX_SVG_BYTES, svgDrawable } from '../svg-picture.ts';

const NS = 'xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"';
const W = 'width="1600" height="900" viewBox="0 0 1600 900"';
const svg = (body: string, prolog = '') => `${prolog}<svg ${NS} ${W}>${body}</svg>`;
let chain = '<g id="a0"><rect width="1600" height="900" fill="red" opacity="0.5"/></g>';
for (let i = 1; i <= 9; i++) chain += `<g id="a${i}">${`<use href="#a${i - 1}"/>`.repeat(10)}</g>`;
const M = `<defs><marker id="m" markerWidth="9" markerHeight="9">${'<rect width="3" height="3"/>'.repeat(100)}</marker></defs>`;
const P = `d="M0 0${' l0.3 0.1'.repeat(5_000)}" stroke="black" fill="none"`;
const longD = (n: number) => `M0 0${' L1600 900 L0 900'.repeat(n / 2)}`;
const fullRect = 'width="1600" height="900"';

const CASES: Record<string, () => string> = {
  smilSetHref: () =>
    svg(
      `<defs>${chain}<g id="small"/></defs><use href="#small"><set attributeName="href" to="#a9" begin="0s"/></use>`
    ),
  smilSetXlinkHref: () =>
    svg(
      `<defs>${chain}<g id="small"/></defs><use xlink:href="#small"><set attributeName="xlink:href" to="#a9" begin="0s"/></use>`
    ),
  smilAnimateValuesHref: () =>
    svg(
      `<defs>${chain}<g id="small"/></defs><use href="#small"><animate attributeName="href" values="#a9;#a9" dur="1s" repeatCount="indefinite"/></use>`
    ),
  smilSetMarker: () => svg(`${M}<path ${P}><set attributeName="marker-mid" to="url(#m)"/></path>`),
  styleCommentSplitMarker: () =>
    svg(`<style>path{mar<!---->ker-mid:url(#m)}</style>${M}<path ${P}/>`),
  styleCdataSplitMarker: () =>
    svg(`<style>path{mar<![CDATA[ker-mid:url(#m)]]>}</style>${M}<path ${P}/>`),
  styleCommentSplitUrl: () =>
    svg(
      `<defs>${chain}</defs><style>.x{fill:u<!---->rl(#p)}</style><defs><pattern id="p" width="1" height="1"><use href="#a9"/></pattern></defs><rect class="x" ${fullRect}/>`
    ),
  useHugePath: () =>
    svg(
      `<defs><path id="p" d="${longD(100_000)}" fill="none" stroke="black"/></defs>${'<use href="#p"/>'.repeat(1_000)}`
    ),
  filterChainBlur: () =>
    svg(
      `<defs><filter id="f" x="0" y="0" width="1" height="1">${'<feGaussianBlur stdDeviation="20"/>'.repeat(2_000)}</filter></defs><rect ${fullRect} fill="blue" filter="url(#f)"/>`
    ),
  filterTurbulence: () =>
    svg(
      `<defs><filter id="f" x="0" y="0" width="1" height="1"><feTurbulence baseFrequency="0.01" numOctaves="1000"/></filter></defs>${`<rect ${fullRect} filter="url(#f)"/>`.repeat(50)}`
    ),
  filterConvolve: () =>
    svg(
      `<defs><filter id="f" x="0" y="0" width="1" height="1"><feConvolveMatrix order="60" kernelMatrix="${'1 '.repeat(3600)}"/></filter></defs><rect ${fullRect} fill="blue" filter="url(#f)"/>`
    ),
  filterMorphology: () =>
    svg(
      `<defs><filter id="f" x="0" y="0" width="1" height="1"><feMorphology operator="dilate" radius="500"/></filter></defs>${`<rect ${fullRect} fill="blue" filter="url(#f)"/>`.repeat(20)}`
    ),
  maskOverdraw: () =>
    svg(
      `<defs><mask id="k"><rect ${fullRect} fill="white"/></mask></defs>${`<rect ${fullRect} fill="red" mask="url(#k)"/>`.repeat(5_000)}`
    ),
  opacityOverdraw: () =>
    svg(
      `<defs><g id="r"><rect ${fullRect} fill="red"/><rect ${fullRect} fill="blue"/></g></defs>${'<use href="#r" opacity="0.5"/>'.repeat(20_000)}`
    ),
  dasharray: () =>
    svg(`${`<path d="${longD(20)}" stroke="black" stroke-dasharray="0.01"/>`.repeat(200)}`),
  hugeText: () =>
    svg(
      `<defs><text id="t" y="20" font-size="1">${'Wi '.repeat(100_000)}</text></defs>${'<use href="#t"/>'.repeat(100)}`
    ),
  textPath: () =>
    svg(
      `<defs><path id="tp" d="${longD(2_000)}"/></defs><text font-size="2"><textPath href="#tp">${'abc'.repeat(100_000)}</textPath></text>`
    ),
  tinyPattern: () =>
    svg(
      `<defs><pattern id="p" width="0.01" height="0.01" patternUnits="userSpaceOnUse"><rect width="0.005" height="0.005"/></pattern></defs>${`<rect ${fullRect} fill="url(#p)"/>`.repeat(100)}`
    ),
  checkerDuplicateIds: () =>
    svg(`<defs>${'<g id="x"/>'.repeat(90_000)}</defs>${'<use href="#x"/>'.repeat(90_000)}`),
  checkerRulesTimesElements: () =>
    svg(`<style>${'.a{fill:url(#b)}'.repeat(400_000)}</style>${'<g class="z"/>'.repeat(150_000)}`),
  checkerSheetNoBraces: () => svg(`<style>${'a'.repeat(2_000_000)}</style>`),
  checkerUrlNoParen: () => svg(`<g fill="${'url('.repeat(500_000)}"/>`),
  checkerCommentOpen: () => svg(`<g style="${'/*'.repeat(1_000_000)}"/>`),
  checkerFontFaceOpen: () => svg(`<style>${'@font-face{'.repeat(300_000)}</style>`),
  controlSmilWithoutSet: () => svg(`<defs>${chain}<g id="small"/></defs><use href="#small"/>`),
  useHugePath10: () =>
    svg(
      `<defs><path id="p" d="${longD(100_000)}" fill="none" stroke="black"/></defs>${'<use href="#p"/>'.repeat(10)}`
    ),
  useHugePath100: () =>
    svg(
      `<defs><path id="p" d="${longD(100_000)}" fill="none" stroke="black"/></defs>${'<use href="#p"/>'.repeat(100)}`
    ),
  smilSetHrefSmall: () =>
    svg(
      `<defs><g id="b0"><rect ${fullRect} fill="red" opacity="0.5"/></g><g id="b1">${'<use href="#b0"/>'.repeat(100)}</g><g id="b2">${'<use href="#b1"/>'.repeat(100)}</g><g id="b3">${'<use href="#b2"/>'.repeat(100)}</g><g id="small"/></defs><use href="#small"><set attributeName="href" to="#b3"/></use>`
    ),
  cdataMarkerBig: () =>
    svg(
      `<style>path{mar<![CDATA[ker-mid:url(#mb)]]>}</style><defs><marker id="mb" markerWidth="9" markerHeight="9" overflow="visible">${'<rect width="30" height="30" opacity="0.5"/>'.repeat(2000)}</marker></defs><path d="M0 0${' l0.08 0.045'.repeat(20_000)}" stroke="black" fill="none"/>`
    ),
  smilMarkerBig: () =>
    svg(
      `<defs><marker id="mb" markerWidth="9" markerHeight="9" overflow="visible">${'<rect width="30" height="30" opacity="0.5"/>'.repeat(2000)}</marker></defs><path d="M0 0${' l0.08 0.045'.repeat(20_000)}" stroke="black" fill="none"><set attributeName="marker-mid" to="url(#mb)"/></path>`
    ),
  controlNoMarkerBig: () =>
    svg(
      `<defs><marker id="mb" markerWidth="9" markerHeight="9" overflow="visible">${'<rect width="30" height="30" opacity="0.5"/>'.repeat(2000)}</marker></defs><path d="M0 0${' l0.08 0.045'.repeat(20_000)}" stroke="black" fill="none"/>`
    ),
  defsOnlyChain: () => svg(`<defs>${chain}</defs>`),
  defsOnlyChain1e6: () =>
    svg(
      `<defs><g id="b0"><rect width="9" height="9"/></g><g id="b1">${'<use href="#b0"/>'.repeat(100)}</g><g id="b2">${'<use href="#b1"/>'.repeat(100)}</g><g id="b3">${'<use href="#b2"/>'.repeat(100)}</g></defs>`
    ),
  symbolOnlyChain: () => svg(`<symbol id="s">${chain}</symbol>`),
  drawioSwitch: () =>
    svg(
      `<g><switch><foreignObject pointer-events="none" width="100%" height="100%" requiredFeatures="http://www.w3.org/TR/SVG11/feature#Extensibility" style="overflow: visible; text-align: left;"><div xmlns="http://www.w3.org/1999/xhtml">Label</div></foreignObject><text x="10" y="20">Label</text></switch></g>`
    ),
  illustratorSwitch: () =>
    svg(
      `<switch><foreignObject requiredExtensions="http://ns.adobe.com/AdobeIllustrator/10.0/" x="0" y="0" width="1" height="1"/><g><style>.st0{fill:url(#SVGID_1_);}</style><linearGradient id="SVGID_1_"><stop offset="0"/></linearGradient><rect class="st0" width="9" height="9"/></g></switch>`
    ),
  mermaidHtmlLabels: () =>
    svg(
      `<g class="node"><rect width="80" height="40"/><g class="label"><foreignObject width="60" height="20"><div xmlns="http://www.w3.org/1999/xhtml"><span class="nodeLabel">A</span></div></foreignObject></g></g>`
    ),
  inkscapeGradients: () =>
    svg(
      `<defs><linearGradient id="lg1"><stop offset="0" style="stop-color:#000"/></linearGradient><linearGradient xlink:href="#lg1" id="lg2" gradientUnits="userSpaceOnUse"/><filter id="f1" style="color-interpolation-filters:sRGB"><feGaussianBlur stdDeviation="2"/></filter></defs><g inkscape:label="Layer 1" xmlns:inkscape="http://www.inkscape.org/namespaces/inkscape"><path style="fill:url(#lg2);filter:url(#f1)" d="M0 0h9v9z"/></g>`
    ),
  matplotlibScatter: () =>
    svg(
      `<defs><style type="text/css">*{stroke-linejoin: round; stroke-linecap: butt}</style><path id="m1" d="M 0 3 C 0.8 3 1.5 2.7 2.1 2.1 z" style="stroke:#1f77b4"/><clipPath id="p1"><rect width="1600" height="900"/></clipPath></defs><g clip-path="url(#p1)">${'<use xlink:href="#m1" x="10" y="20" style="fill:#1f77b4"/>'.repeat(20_000)}</g>`
    ),
  figmaShadow: () =>
    svg(
      `<g filter="url(#filter0_d)"><rect width="9" height="9" fill="url(#pattern0)"/></g><defs><filter id="filter0_d" filterUnits="userSpaceOnUse" color-interpolation-filters="sRGB"><feFlood flood-opacity="0" result="BackgroundImageFix"/><feColorMatrix in="SourceAlpha" type="matrix" values="0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 127 0" result="hardAlpha"/><feOffset dy="4"/><feGaussianBlur stdDeviation="2"/><feBlend mode="normal" in2="BackgroundImageFix"/></filter><pattern id="pattern0" patternContentUnits="objectBoundingBox" width="1" height="1"><use xlink:href="#image0" transform="scale(0.01)"/></pattern><image id="image0" width="100" height="100" xlink:href="data:image/png;base64,iVBORw0KGgo="/></defs>`
    ),
  excalidrawFonts: () =>
    svg(
      `<defs><style class="style-fonts">@font-face { font-family: "Virgil"; src: url(data:font/woff2;base64,AAAA); } @font-face { font-family: "X"; src: url("data:application/font-woff;charset=utf-8;base64,AAAA") format("woff"); }</style></defs><text font-family="Virgil">hi</text>`
    ),
  highchartsMedia: () =>
    svg(
      `<style>@media print { .highcharts-tooltip { display:none } } .series path{clip-path:url(#hc-clip)}</style><defs><clipPath id="hc-clip"><rect width="9" height="9"/></clipPath></defs>${'<g class="series"><path d="M0 0L9 9"/></g>'.repeat(3_000)}`
    ),
  nestedSvgImage: () =>
    svg(`<image width="9" height="9" href="data:image/svg+xml;base64,PHN2Zy8+"/>`),
  imageXPng: () => svg(`<image width="9" height="9" href="data:image/x-png;base64,iVBOR"/>`),
  emptyTypeFont: () => svg(`<style>@font-face{font-family:a;src:url(data:;base64,AAAA)}</style>`),
  officeIcon: () =>
    svg(
      `<g><path d="M48 12C29.2 12 14 27.2 14 46s15.2 34 34 34 34-15.2 34-34-15.2-34-34-34z" fill="#4472C4"/></g>`
    ),
};

/** Ordinary files inside the subset, and two harmless sheets the check reads in linear time. */
const ACCEPTED = new Set([
  'checkerCommentOpen',
  'checkerSheetNoBraces',
  'emptyTypeFont',
  'excalidrawFonts',
  'highchartsMedia',
  'matplotlibScatter',
  'officeIcon',
]);
/** Ordinary files that need an element or a property outside the subset: they draw the PNG. */
const FALLBACK: Record<string, string> = {
  drawioSwitch: 'a <switch> and a <foreignObject>',
  illustratorSwitch: 'a <switch> and a <foreignObject>',
  mermaidHtmlLabels: 'a <foreignObject>',
  inkscapeGradients: 'a filter',
  figmaShadow: 'a filter and a pattern',
  nestedSvgImage: 'an SVG document in an <image>',
  imageXPng: 'the image/x-png type',
};

const enc = (s: string) => new TextEncoder().encode(s);

for (const [name, make] of Object.entries(CASES)) {
  const accepted = ACCEPTED.has(name);
  const why = FALLBACK[name];
  const label = accepted
    ? 'draws as vector art'
    : why
      ? `draws the PNG: it needs ${why}`
      : 'draws its raster alternative';
  test(`${name} ${label}`, () => {
    expect(svgDrawable(enc(make()))).toBe(accepted);
  });
}

/** The fastest of three checks of `bytes`, in milliseconds. */
function checkMs(bytes: Uint8Array): number {
  let best = Infinity;
  for (let k = 0; k < 3; k += 1) {
    // A copy: the check keeps its answer for each array.
    const copy = bytes.slice();
    const t0 = performance.now();
    svgDrawable(copy);
    best = Math.min(best, performance.now() - t0);
  }
  return best;
}

/** Inputs that made a checker slow, at `n` bytes or so. */
const SLOW: Record<string, (n: number) => string> = {
  sheetNoBraces: (n) => svg(`<style>${'a'.repeat(n)}</style>`),
  urlNoParen: (n) => svg(`<g fill="${'url('.repeat(n / 4)}"/>`),
  duplicateIds: (n) =>
    svg(`<defs>${'<g id="x"/>'.repeat(n / 10)}</defs>${'<use href="#x"/>'.repeat(n / 10)}`),
  rulesTimesElements: (n) =>
    svg(`<style>${'.a{fill:url(#b)}'.repeat(n / 20)}</style>${'<g class="z"/>'.repeat(n / 20)}`),
  commentOpen: (n) => svg(`<g style="${'/*'.repeat(n / 2)}"/>`),
};

for (const [name, make] of Object.entries(SLOW))
  test(`the check of ${name} grows linearly: twice the input takes less than three times as long`, () => {
    const n = MAX_SVG_BYTES / 4;
    const small = checkMs(enc(make(n)));
    const large = checkMs(enc(make(2 * n)));
    expect(large).toBeLessThan(Math.max(small, 1) * 3);
  });
