// An SVG picture draws as vector art only when it is a plain icon (svg-picture.ts). Each icon
// below draws as vector art. Each other picture draws its raster alternative. More probes,
// and the cost of the check itself: svg-picture-subset.test.ts.

import { expect, test } from 'bun:test';
import {
  MAX_SVG_BYTES,
  MAX_SVG_EVEN_ODD,
  MAX_SVG_LAYERS,
  MAX_SVG_SHAPES,
  MAX_SVG_VERTICES,
  svgDrawable,
} from '../svg-picture.ts';

const enc = (s: string) => new TextEncoder().encode(s);
const NS = 'xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"';
const svg = (body: string, prolog = '', root = 'viewBox="0 0 96 96"') =>
  enc(`${prolog}<svg ${NS} ${root}>${body}</svg>`);
const drawn = (bytes: Uint8Array) => svgDrawable(bytes);

const ICONS: Record<string, Uint8Array> = {
  themeClassIcon: svg(
    '<style>.MsftOfcThm_Accent1_Fill_v2 {fill:#4472C4;}</style>' +
      '<path d="M48 12C29.2 12 14 27.2 14 46s15.2 34 34 34 34-15.2 34-34-15.2-34-34-34z"' +
      ' class="MsftOfcThm_Accent1_Fill_v2" fill="#4472C4"/>',
    '',
    'width="96" height="96" viewBox="0 0 96 96" id="Icons_Person" overflow="hidden"'
  ),
  percentSize: svg(
    '<g><path d="M48 12C29 12 14 27 14 46z" fill="#4472C4"/></g>',
    '',
    'width="100%" height="100%" viewBox="0 0 96 96" overflow="hidden"'
  ),
  generatorDoctype: svg(
    '<style type="text/css">.st0{fill:#FFFFFF;}</style><path class="st0" d="M0 0h10v10z"/>',
    '<?xml version="1.0" encoding="utf-8"?>\n<!-- Generator: an editor -->\n' +
      '<!DOCTYPE svg PUBLIC "-//W3C//DTD SVG 1.1//EN" ' +
      '"http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd">\n',
    'version="1.1" id="Layer_1" x="0px" y="0px" viewBox="0 0 512 512"' +
      ' style="enable-background:new 0 0 512 512;" xml:space="preserve"'
  ),
  logo: svg(
    '<title>Logo</title><desc>A ring and a bar</desc>' +
      '<g transform="translate(8 8) scale(0.8)"><circle cx="40" cy="40" r="30" fill="none"' +
      ' stroke="rgb(20, 40, 160)" stroke-width="8"/><rect x="10" y="70" width="60" height="8"' +
      ' rx="4" fill="#e23" opacity="0.8"/></g>'
  ),
  arrow: svg(
    '<line x1="8" y1="48" x2="80" y2="48" stroke="currentColor" stroke-width="6"' +
      ' stroke-linecap="round"/><polyline points="60,28 82,48 60,68" fill="none"' +
      ' stroke="currentColor" stroke-width="6" stroke-linejoin="round"/>'
  ),
  ring: svg(
    '<path fill-rule="evenodd" d="M48 8a40 40 0 1 0 0 80a40 40 0 1 0 0-80zm0 16a24 24 0 1 1' +
      ' 0 48a24 24 0 1 1 0-48z" fill="#0a0"/>'
  ),
  symbolUse: svg(
    '<defs><symbol id="dot" viewBox="0 0 10 10"><circle cx="5" cy="5" r="4"/></symbol>' +
      '<clipPath id="c"><rect width="96" height="96" rx="12"/></clipPath></defs>' +
      '<g clip-path="url(#c)" style="fill:navy;stroke:none">' +
      '<use href="#dot" x="10" y="10" width="20" height="20"/>' +
      '<use xlink:href="#dot" x="60" y="60" width="20" height="20"/></g>'
  ),
  noViewBox: svg(
    '<g transform="translate(-592 -312)"><path d="M640 320 L 680 360Z" fill="#000000"/></g>',
    '',
    'width="96px" height="96" xml:space="preserve" overflow="hidden"'
  ),
  editorData: svg(
    '<sodipodi:namedview id="n" pagecolor="#fff" inkscape:zoom="2"><inkscape:grid id="g"/>' +
      '</sodipodi:namedview><g inkscape:label="Layer 1" inkscape:groupmode="layer">' +
      '<path d="M1 1h20v20z" sodipodi:nodetypes="ccc"/></g>',
    '',
    'viewBox="0 0 24 24" inkscape:version="1.0"' +
      ' xmlns:sodipodi="http://sodipodi.sourceforge.net/DTD/sodipodi-0.dtd"' +
      ' xmlns:inkscape="http://www.inkscape.org/namespaces/inkscape"'
  ),
  textInTitleAndDesc: svg('<title>Icon <!-- a note --> name</title><desc>About it</desc><rect/>'),
  metadata: svg(
    '<metadata><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"/></metadata>' +
      '<ellipse cx="48" cy="48" rx="40" ry="20" fill="#ccc" data-name="shape" aria-hidden="true"/>'
  ),
};

for (const [name, bytes] of Object.entries(ICONS))
  test(`an icon draws as vector art: ${name}`, () => {
    expect(drawn(bytes)).toBe(true);
  });

/** A path that crosses itself `n` times across the box. */
const crossing = (n: number) => {
  let d = 'M0 0';
  for (let i = 0; i < n; i += 1) d += ` L${i % 2 ? 0 : 96} ${(i * 37) % 96}`;
  return d;
};
const rect = '<rect width="96" height="96" fill="red" stroke="blue" stroke-width="4"/>';
const layer = '<g opacity="0.5"><rect width="96" height="96"/></g>';
/** The most vertices of one path with the even-odd rule: v / MAX_SVG_VERTICES + v² / MAX = 1. */
const a = 1 / MAX_SVG_VERTICES;
const evenOddRoot = (-a + Math.sqrt(a * a + 4 / MAX_SVG_EVEN_ODD)) / (2 / MAX_SVG_EVEN_ODD);
const [evenOdd, pastEvenOdd] = [Math.floor(evenOddRoot) - 5, Math.ceil(evenOddRoot) + 5];

const LIMITS: Record<string, [Uint8Array, Uint8Array]> = {
  shapes: [svg(rect.repeat(MAX_SVG_SHAPES - 1)), svg(rect.repeat(MAX_SVG_SHAPES + 1))],
  vertices: [
    svg(`<path d="${crossing(MAX_SVG_VERTICES - 10)}"/>`),
    svg(`<path d="${crossing(MAX_SVG_VERTICES)}"/>`),
  ],
  layers: [
    svg('<g opacity="0.5"/>'.repeat(MAX_SVG_LAYERS - 1)),
    svg('<g opacity="0.5"/>'.repeat(MAX_SVG_LAYERS + 1)),
  ],
  evenOdd: [
    svg(`<path fill-rule="evenodd" d="${crossing(evenOdd)}"/>`),
    svg(`<path fill-rule="evenodd" d="${crossing(pastEvenOdd)}"/>`),
  ],
  evenOddInSheet: [
    svg(`<style>*{clip-rule:evenodd}</style><path d="${crossing(evenOdd)}"/>`),
    svg(`<style>*{clip-rule:evenodd}</style><path d="${crossing(pastEvenOdd)}"/>`),
  ],
  sharesAddUp: [
    svg(rect.repeat(MAX_SVG_SHAPES / 2 - 10) + layer.repeat(MAX_SVG_LAYERS / 3)),
    svg(rect.repeat(MAX_SVG_SHAPES / 2) + layer.repeat(MAX_SVG_LAYERS / 2)),
  ],
  useCopies: [
    svg(`<defs><path id="p" d="${crossing(99)}"/></defs>${'<use href="#p"/>'.repeat(30)}`),
    svg(`<defs><path id="p" d="${crossing(99)}"/></defs>${'<use href="#p"/>'.repeat(41)}`),
  ],
  sheetOpacity: [
    svg(`<style>g{opacity:.5}</style>${'<g/>'.repeat(MAX_SVG_LAYERS - 10)}`),
    svg(`<style>g{opacity:.5}</style>${'<g/>'.repeat(MAX_SVG_LAYERS)}`),
  ],
};

for (const [name, [inside, outside]] of Object.entries(LIMITS))
  test(`an icon at the ${name} limit draws, and one past it draws the raster alternative`, () => {
    expect(drawn(inside)).toBe(true);
    expect(drawn(outside)).toBe(false);
  });

test('an SVG picture longer than the size limit draws the raster alternative', () => {
  const comment = `<!--${' '.repeat(MAX_SVG_BYTES)}-->`;
  expect(drawn(svg(`<rect/>${comment}`))).toBe(false);
});

/** Nine levels of ten `<use>` copies: 10^9 paths. `ref` spells the link to `a{i}`. */
function chain(ref = (i: number) => `#a${i}`, extra = ''): string {
  let s = '<g id="a0"><path d="M0 0h1"/></g>';
  for (let i = 1; i <= 9; i += 1)
    s += `<g id="a${i}">${`<use href="${ref(i - 1)}"/>`.repeat(10)}</g>${extra.replace('#', `${i}`)}`;
  return s;
}
const M = `<defs><marker id="m">${'<rect/>'.repeat(100)}</marker></defs>`;
const P = `d="M0 0${' l1 1'.repeat(500)}"`;
let masks = '<mask id="k0"><rect width="9" height="9" fill="white"/></mask>';
for (let i = 1; i <= 10; i += 1)
  masks += `<mask id="k${i}">${`<rect width="9" height="9" mask="url(#k${i - 1})"/>`.repeat(10)}</mask>`;
let clips = '<clipPath id="c0"><rect width="9" height="9"/></clipPath>';
for (let i = 1; i <= 10; i += 1)
  clips += `<clipPath id="c${i}">${`<rect width="9" height="9" clip-path="url(#c${i - 1})"/>`.repeat(10)}</clipPath>`;
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';

/** Pictures outside the icon subset, each refused for the reason in its name. */
const REFUSED: Record<string, Uint8Array> = {
  useChain: svg(`<defs>${chain()}</defs><use href="#a9"/>`),
  useChainInDefs: svg(`<defs>${chain()}<g><use href="#a9"/></g></defs>`),
  useChainInSymbol: svg(`<symbol id="s">${chain()}<use href="#a9"/></symbol>`),
  clipChain: svg(`<defs>${clips}</defs><rect clip-path="url(#c10)"/>`),
  clipToNonClip: svg(`<defs><g id="g"/></defs><rect clip-path="url(#g)"/>`),
  useCycle: svg('<g id="a"><use href="#a"/></g>'),
  useMissing: svg('<use href="#nothing"/>'),
  useExternal: svg('<use href="other.svg#a"/>'),
  useLongChain: svg(
    Array.from({ length: 10 }, (_, i) => `<g id="u${i}"><use href="#u${i + 1}"/></g>`).join('') +
      '<g id="u10"/>'
  ),
  duplicateIds: svg(`<defs>${chain(undefined, '<g id="a#"/>')}</defs><use href="#a9"/>`),
  percentId: svg(`<defs>${chain((i) => `#%61${i}`)}</defs><use href="#%619"/>`),
  noViewBox: svg('<rect/>'.repeat(2), '', ''),
  noViewBoxPercentSize: svg('<rect/>', '', 'width="100%" height="96"'),
  noViewBoxInfiniteSize: svg('<rect/>', '', 'width="1e999" height="96"'),
  zeroViewBox: svg('<rect/>', '', 'viewBox="0 0 0 96"'),
  infiniteViewBox: svg('<rect/>', '', 'viewBox="0 0 1e999 96"'),
  slice: svg('<rect/>', '', 'viewBox="0 0 96 96" preserveAspectRatio="xMidYMid slice"'),
  upperCaseElement: svg('<RECT/>'),
  otherNamespaceHoldsSvg: svg('<x:g xmlns:x="urn:x"><rect/></x:g>'),
  otherNamespaceHoldsXhtml: svg(
    '<x:g xmlns:x="urn:x"><p xmlns="http://www.w3.org/1999/xhtml"/></x:g>'
  ),
  otherNamespaceStyle: svg('<x:style xmlns:x="urn:x">rect{filter:blur(9px)}</x:style><rect/>'),
  otherNamespaceHoldsSet: svg('<x:g xmlns:x="urn:x"><x:set attributeName="fill"/></x:g><rect/>'),
  unboundPrefixElement: svg('<x:g/>'),
  rootInOtherNamespace: enc('<svg xmlns="urn:x" viewBox="0 0 9 9"/>'),
  twoRoots: enc(`<svg ${NS} viewBox="0 0 9 9"/><svg ${NS} viewBox="0 0 9 9"/>`),
  manyAttributes: svg(
    `<rect ${Array.from({ length: 65 }, (_, i) => `data-a${i}="1"`).join(' ')}/>`
  ),
  text: svg('<text x="9" y="9">Label</text>'),
  gradient: svg(
    '<defs><linearGradient id="g"><stop offset="0"/></linearGradient></defs><rect fill="url(#g)"/>'
  ),
  image: svg(`<image width="9" height="9" href="data:image/png;base64,${PNG}"/>`),
  externalImage: svg('<image width="9" height="9" href="https://example.test/a.png"/>'),
  filter: svg('<defs><filter id="f"><feGaussianBlur/></filter></defs><rect filter="url(#f)"/>'),
  mask: svg(`<defs>${masks}</defs><rect mask="url(#k10)"/>`),
  pattern: svg('<defs><pattern id="p"/></defs><rect fill="url(#p)"/>'),
  marker: svg(`${M}<path marker-mid="url(#m)" ${P}/>`),
  markerInSheet: svg(`<style>path{marker-mid:url(#m)}</style>${M}<path ${P}/>`),
  markerInStyle: svg(`${M}<path style="marker:url(#m)" ${P}/>`),
  animation: svg('<use href="#a"><set attributeName="href" to="#b"/></use><g id="a"/>'),
  foreignObject: svg('<foreignObject width="9" height="9"/>'),
  switch: svg('<switch><rect/></switch>'),
  script: svg('<script>1</script>'),
  anchor: svg('<a href="https://example.test"><rect/></a>'),
  eventAttribute: svg('<rect onclick="x()"/>'),
  urlPaint: svg('<rect fill="url(#p)"/>'),
  urlPaintInStyle: svg('<rect style="fill:url(#p)"/>'),
  urlPaintFallback: svg('<rect fill="url(#p) red"/>'),
  dasharray: svg('<path d="M0 0h96" stroke="red" stroke-dasharray="0.01"/>'),
  miterClip: svg('<path d="M0 0h96" stroke="red" stroke-linejoin="miter-clip"/>'),
  highMiter: svg('<path d="M0 0h96" stroke="red" stroke-miterlimit="1e9"/>'),
  percentLength: svg('<rect width="100%" height="100%"/>'),
  emLength: svg('<rect width="1em" height="1em"/>'),
  calcLength: svg('<rect style="stroke-width:calc(5000px)"/>'),
  cssScale: svg('<g style="scale:2000"><rect/></g>'),
  blend: svg('<g style="mix-blend-mode:difference"><rect/></g>'),
  paintOrder: svg('<rect paint-order="stroke" vector-effect="non-scaling-stroke"/>'),
  transformFunction: svg('<g transform="perspective(9)"><rect/></g>'),
  infiniteTransform: svg('<g transform="scale(1e999)"><rect/></g>'),
  pathLetter: svg('<path d="M0 0 X9 9"/>'),
  sheetAtRule: svg('<style>@media all{rect{fill:red}}</style><rect/>'),
  sheetEscape: svg('<style>.\\61{fill:red}</style><rect class="a"/>'),
  sheetCombinator: svg('<style>g rect{fill:red}</style><rect/>'),
  sheetAttributeSelector: svg('<style>[d]{fill:red}</style><rect/>'),
  sheetVar: svg('<style>svg{--m:red} rect{fill:var(--m)}</style><rect/>'),
  sheetNested: svg('<style>g{rect{fill:red}}</style><rect/>'),
  sheetOpen: svg('<style>rect{fill:red</style><rect/>'),
  sheetTextOutsideRules: svg('<style>rect{fill:red} x</style><rect/>'),
  sheetCommentSplit: svg('<style>rect{mar<!---->ker-mid:url(#m)}</style><rect/>'),
  sheetCdataSplit: svg('<style>rect{mar<![CDATA[ker-mid:url(#m)]]>}</style><rect/>'),
  xmlStylesheet: svg('', '<?xml-stylesheet type="text/css" href="data:text/css,a"?>'),
  internalDtd: svg('', '<!DOCTYPE svg [<!ENTITY a "a">]>'),
  twoDoctypes: svg('', '<!DOCTYPE svg><!DOCTYPE svg>'),
  svgNamespaceAttribute: svg('<rect s:fill="red" xmlns:s="http://www.w3.org/2000/svg"/>'),
  unboundPrefixAttribute: svg('<rect x:fill="red"/>'),
  xmlBase: svg('<rect xml:base="https://example.test/"/>'),
  metadataStyle: svg(
    '<metadata><style>line{stroke-dasharray:0.001;filter:blur(30px)}</style></metadata>' +
      '<line x1="0" y1="0" x2="96" y2="0" stroke="black"/>'.repeat(900)
  ),
  titleSet: svg(
    '<title><set href="#p" attributeName="fill" to="blue"/></title><path id="p" d="M0 0h9"/>'
  ),
  descFilter: svg(
    '<desc><filter id="f"><feGaussianBlur stdDeviation="50"/></filter></desc><rect/>'
  ),
  metadataXhtml: svg('<metadata><div xmlns="http://www.w3.org/1999/xhtml"/></metadata><rect/>'),
  descOtherNamespaceStyle: svg(
    '<desc><x:style xmlns:x="urn:x">rect{fill:red}</x:style></desc><rect/>'
  ),
  titleNestedSet: svg(
    '<title><x:a xmlns:x="urn:x"><set attributeName="fill"/></x:a></title><rect/>'
  ),
  hrefOnRect: svg('<rect href="#a"/><g id="a"/>'),
};

for (const [name, bytes] of Object.entries(REFUSED))
  test(`an SVG picture outside the icon subset draws its raster alternative: ${name}`, () => {
    expect(drawn(bytes)).toBe(false);
  });
