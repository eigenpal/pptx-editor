// An SVG picture draws as vector art only inside a small subset, within a cost budget. Each
// hostile picture below draws its raster alternative; each ordinary one draws as vector art.
// More cases, and the cost of the check itself: svg-picture-subset.test.ts.

import { expect, test } from 'bun:test';
import { svgDrawable } from '../svg-picture.ts';

const enc = (s: string) => new TextEncoder().encode(s);
const NS = 'xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"';
const svg = (body: string, prolog = '', attrs = '') =>
  enc(`${prolog}<svg ${NS} ${attrs}>${body}</svg>`);

/** The signature and header of a PNG of `w` x `h` pixels, in base64: no pixels follow. */
function pngHeader(w: number, h: number): string {
  const b = new Uint8Array(33);
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  new DataView(b.buffer).setUint32(16, w);
  new DataView(b.buffer).setUint32(20, h);
  b[24] = 8;
  b[25] = 6;
  return btoa(String.fromCharCode(...b));
}

/** Nine levels of ten `<use>` copies: 10^9 paths. `ref` spells the link to `a{i}`. */
function chain(ref = (i: number) => `#a${i}`, extra = ''): string {
  let s = '<g id="a0"><path d="M0 0h1"/></g>';
  for (let i = 1; i <= 9; i += 1)
    s += `<g id="a${i}">${`<use href="${ref(i - 1)}"/>`.repeat(10)}</g>${extra.replace('#', `${i}`)}`;
  return s;
}
const M = `<defs><marker id="m">${'<rect/>'.repeat(100)}</marker></defs>`;
const P = `d="M0 0${' l1 1'.repeat(5_000)}"`;
const at = (n: number) => '<rect width="9" height="9"/>'.repeat(n);
let masks = '<mask id="k0"><rect width="9" height="9" fill="white"/></mask>';
for (let i = 1; i <= 10; i += 1)
  masks += `<mask id="k${i}">${`<rect width="9" height="9" mask="url(#k${i - 1})"/>`.repeat(10)}</mask>`;

const HOSTILE: Record<string, Uint8Array> = {
  use: svg(`<defs>${chain()}</defs><use href="#a9"/>`),
  mask: svg(`<defs>${chain()}<mask id="m"><use href="#a9"/></mask></defs><rect mask="url(#m)"/>`),
  pattern: svg(
    `<defs>${chain()}<pattern id="p"><use href="#a9"/></pattern></defs><rect fill="url(#p)"/>`
  ),
  clipPath: svg(
    `<defs>${chain()}<clipPath id="c"><use href="#a9"/></clipPath></defs><rect clip-path="url(#c)"/>`
  ),
  feImage: svg(
    `<defs>${chain()}<filter id="f"><feImage href="#a9"/></filter></defs><rect filter="url(#f)"/>`
  ),
  maskManyLinks: svg(
    `<defs><mask id="m">${at(1_000)}</mask></defs>${'<rect mask="url(#m)"/>'.repeat(1_000)}`
  ),
  nestedMasks: svg(`<defs>${masks}</defs><rect mask="url(#k10)"/>`),
  nestedMasksOutsideDefs: svg(`${masks}<rect mask="url(#k10)"/>`),
  nestedSvgUse: svg(`<defs>${chain()}</defs><svg><use href="#a9"/></svg>`),
  duplicateIds: svg(`<defs>${chain(undefined, '<g id="a#"/>')}</defs><use href="#a9"/>`),
  duplicateMarker: svg(
    `<defs><marker id="m">${at(1_000)}</marker><marker id="m"/></defs><path marker-mid="url(#m)" ${P}/>`
  ),
  percentUse: svg(`<defs>${chain((i) => `#%61${i}`)}</defs><use href="#%619"/>`),
  percentMarker: svg(`${M}<path marker-mid="url(#%6d)" ${P}/>`),
  foreignObject: svg('<foreignObject width="9" height="9"/>'),
  srcset: svg(
    '<foreignObject><img xmlns="http://www.w3.org/1999/xhtml" srcset="data:image/svg+xml,%3Csvg/%3E 1x"/></foreignObject>'
  ),
  imageSet: svg('<style>g{background-image:image-set("data:image/svg+xml,%3Csvg/%3E" 1x)}</style>'),
  leadingSpace: svg(`<image href="${' '.repeat(300)}data:image/svg+xml,%3Csvg/%3E"/>`),
  leadingTabs: svg(`<image href="${'&#9;'.repeat(300)}data:image/svg+xml,%3Csvg/%3E"/>`),
  splitScheme: svg('<image href="da&#10;ta:image/svg+xml,%3Csvg/%3E"/>'),
  plainImage: svg('<image href="data:image/svg+xml,%3Csvg/%3E"/>'),
  xmlStylesheet: svg('', '<?xml-stylesheet type="text/css" href="data:text/css,a"?>'),
  markerAttributeSelector: svg(`<style>[d]{marker-mid:url(#m)}</style>${M}<path ${P}/>`),
  markerOnUse: svg(
    `<defs><marker id="m">${at(100)}</marker><path id="p" ${P}/></defs><use href="#p" style="marker-mid:url(#m)"/>`
  ),
  markerShorthand: svg(`${M}<path marker="url(#m)" ${P}/>`),
  markerImportant: svg(`<style>path{marker-mid:url(#m)!important}</style>${M}<path ${P}/>`),
  markerMedia: svg(`<style>@media all{path{marker-mid:url(#m)}}</style>${M}<path ${P}/>`),
  markerSupports: svg(
    `<style>@supports (display:block){path{marker-mid:url(#m)}}</style>${M}<path ${P}/>`
  ),
  markerSheetVar: svg(`<style>svg{--m:url(#m)} path{marker-mid:var(--m)}</style>${M}<path ${P}/>`),
  markerStyleVar: svg(`<g style="--m:url(#m)">${M}<path style="marker-mid:var(--m)" ${P}/></g>`),
  markerEscapedId: svg(`<style>path{marker-mid:url(#\\6d)}</style>${M}<path ${P}/>`),
  markerEscapedProperty: svg(`<style>path{m\\61rker-mid:url(#m)}</style>${M}<path ${P}/>`),
  markerUpperSelector: svg(`<style>PATH{marker-mid:url(#m)}</style>${M}<path ${P}/>`),
  markerNamespaceSelector: svg(`<style>svg|path{marker-mid:url(#m)}</style>${M}<path ${P}/>`),
  markerEscapedClass: svg(`<style>.\\61{marker-mid:url(#m)}</style>${M}<path class="a" ${P}/>`),
  markerCommentInDeclaration: svg(`<style>path{marker-mid/**/:url(#m)}</style>${M}<path ${P}/>`),
  markerPath: svg(`${M}<path marker-mid="url(#m)" ${P}/>`),
  markerPolyline: svg(`${M}<polyline marker-mid="url(#m)" points="${'1,1 '.repeat(5_000)}"/>`),
  markerImplicitLines: svg(`${M}<path marker-mid="url(#m)" d="M0 0${' 1 1'.repeat(5_000)}"/>`),
  markerPackedArcs: svg(
    `${M}<path marker-mid="url(#m)" d="M0 0a1 1 0 0 1 1 1${' 1 1 0 001 1'.repeat(5_000)}"/>`
  ),
  markerArcsNoSpaces: svg(
    `${M}<path marker-mid="url(#m)" d="M0 0${'a1,1,0,0,1,1,1'.repeat(5_000)}"/>`
  ),
  markerPackedDots: svg(`${M}<path marker-mid="url(#m)" d="M0 0l${'.5.5'.repeat(5_000)}"/>`),
  markerInSymbol: svg(
    `<defs><marker id="m">${at(100)}</marker><symbol id="s"><path marker-mid="url(#m)" d="M0 0${' l1 1'.repeat(200)}"/></symbol></defs>${'<use href="#s"/>'.repeat(20)}`
  ),
  internalDtd: svg('', '<!DOCTYPE svg [<!ENTITY ns_svg "http://www.w3.org/2000/svg">]>'),
  hugePicture: svg(
    `<image width="9" height="9" href="data:image/png;base64,${pngHeader(16_000, 16_000)}"/>`
  ),
  externalPicture: svg('<image width="9" height="9" href="https://example.test/a.png"/>'),
  fullRects: svg('<rect width="1600" height="900"/>'.repeat(5_000), ' ', 'viewBox="0 0 1600 900"'),
};

for (const [name, bytes] of Object.entries(HOSTILE))
  test(`a hostile SVG picture draws its raster alternative: ${name}`, () => {
    expect(svgDrawable(bytes)).toBe(false);
  });

const arrow = '<defs><marker id="a"><path d="M0,0 L10,5 L0,10 z"/></marker></defs>';
/** A valid 1 x 1 PNG. */
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const ORDINARY: Record<string, Uint8Array> = {
  labels: svg(
    '<style>.a{font:12px sans-serif}</style><text class="a" aria-label="Sales data: 2024">Hi</text>' +
      `<image href="data:image/png;base64,${PNG}"/><desc>url(data:text/plain,x)</desc>`
  ),
  fonts: svg(
    '<style>@font-face{font-family:a;src:url(data:font/woff2;base64,AAAA) format("woff2")}' +
      '@font-face{font-family:b;src:url(data:application/x-font-ttf;base64,AAAA)}' +
      '@font-face{font-family:c;src:url(data:application/vnd.ms-fontobject;base64,AAAA)}</style>'
  ),
  octetStreamFont: svg(
    '<style>@font-face{font-family:x;src:url(data:application/octet-stream;base64,AAAA)}</style>'
  ),
  upperCasePng: svg(`<image href="data:image/PNG;base64,${PNG}"/>`),
  longPath: svg(`<path d="M0 0${' l1 1'.repeat(250_000)}"/>`),
  plainDoctype: svg(
    '',
    '<?xml version="1.0"?><!DOCTYPE svg PUBLIC "-//W3C//DTD SVG 1.1//EN" "http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd">'
  ),
  gradients: svg(
    '<defs><linearGradient id="g"><stop offset="0"/><stop offset="1"/></linearGradient></defs>' +
      '<rect fill="url(#g)"/>'.repeat(2_000)
  ),
};

for (const [name, bytes] of Object.entries(ORDINARY))
  test(`an ordinary SVG picture draws: ${name}`, () => {
    expect(svgDrawable(bytes)).toBe(true);
  });

// Markers are outside the subset: these draw their raster alternative.
const MARKERS: Record<string, Uint8Array> = {
  arrows: svg(arrow + '<line x1="0" y1="0" x2="9" y2="9" marker-end="url(#a)"/>'.repeat(500)),
  descendantSelectorArrows: svg(
    `${arrow}<style>#chart .edge path{marker-end:url(#a)}</style>` +
      `<g id="chart">${'<g class="edge"><path d="M0 0L9 9"/></g>'.repeat(200)}</g>`
  ),
};

for (const [name, bytes] of Object.entries(MARKERS))
  test(`an SVG picture with markers draws its raster alternative: ${name}`, () => {
    expect(svgDrawable(bytes)).toBe(false);
  });
