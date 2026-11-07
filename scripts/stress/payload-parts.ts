// Building blocks of the adversarial payload decks in `payload-decks.ts`: picture, SVG, and
// metafile bytes, picture and chart slides, and package edits.

import { strToU8, unzipSync, zipSync, zlibSync, type Zippable } from 'fflate';
import { buildPptx, para, run, shape, solid } from '../fixtures/builder.ts';
import { IN, R_NS, box, one, patched, range } from './deck-parts.ts';

export const R_IMAGE = `${R_NS}/image`;
export const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
export const CHART_CT = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
export const SVG_EXT = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';
export const ASVG = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';

// --- Picture bytes -----------------------------------------------------------------------------

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(parts: readonly Uint8Array[]): number {
  let c = 0xffffffff;
  for (const p of parts) for (const b of p) c = CRC_TABLE[(c ^ b) & 0xff]! ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

export function concat(parts: readonly Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}

function u32be(v: number): Uint8Array {
  const b = new Uint8Array(4);
  new DataView(b.buffer).setUint32(0, v >>> 0);
  return b;
}

function pngChunk(type: string, data: Uint8Array): Uint8Array {
  const t = strToU8(type);
  return concat([u32be(data.length), t, data, u32be(crc32([t, data]))]);
}

const idatCache = new Map<string, Uint8Array>();

/**
 * A real PNG of `w`×`h` pixels, one bit per pixel in gray, in stripes: a few kilobytes that
 * decode to `w`×`h`×4 bytes. `seed` adds a text chunk, so two pictures never share bytes.
 * `idat` replaces the image data (a damaged picture).
 */
export function bitPng(w: number, h: number, seed = 0, idat?: Uint8Array): Uint8Array {
  const ihdr = new Uint8Array(13);
  const v = new DataView(ihdr.buffer);
  v.setUint32(0, w);
  v.setUint32(4, h);
  ihdr.set([1, 0, 0, 0, 0], 8); // bit depth 1, gray, deflate, filter 0, no interlace
  const key = `${w}x${h}`;
  let data = idat ?? idatCache.get(key);
  if (!data) {
    const row = Math.ceil(w / 8) + 1;
    const raw = new Uint8Array(row * h);
    for (let y = 0; y < h; y += 1) raw.fill(y % 64 < 32 ? 0xaa : 0x55, y * row + 1, (y + 1) * row);
    data = zlibSync(raw, { level: 6 });
    idatCache.set(key, data);
  }
  return concat([
    new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', ihdr),
    pngChunk('tEXt', strToU8(`Comment\0${seed}`)),
    pngChunk('IDAT', data),
    pngChunk('IEND', new Uint8Array(0)),
  ]);
}

/**
 * A real baseline JPEG of `w`×`h` gray pixels, all one value: every 8×8 block codes as two
 * bits with one-code Huffman tables, so 100 megapixels store in about 400 kilobytes.
 */
export function flatJpeg(w: number, h: number): Uint8Array {
  const seg = (marker: number, body: number[]): number[] => [
    0xff,
    marker,
    (body.length + 2) >> 8,
    (body.length + 2) & 0xff,
    ...body,
  ];
  const oneCodeTable = (cls: number) => [cls, 1, ...new Array(15).fill(0), 0];
  const blocks = Math.ceil(w / 8) * Math.ceil(h / 8);
  const scan = new Uint8Array(Math.ceil((blocks * 2) / 8)); // DC "0" + EOB "0" per block
  const head = [
    0xff,
    0xd8,
    ...seg(0xdb, [0, ...new Array(64).fill(1)]),
    ...seg(0xc0, [8, h >> 8, h & 0xff, w >> 8, w & 0xff, 1, 1, 0x11, 0]),
    ...seg(0xc4, oneCodeTable(0x00)),
    ...seg(0xc4, oneCodeTable(0x10)),
    ...seg(0xda, [1, 1, 0x00, 0, 63, 0]),
  ];
  return concat([new Uint8Array(head), scan, new Uint8Array([0xff, 0xd9])]);
}

/** A GIF of `w`×`h` with `frames` one-pixel frames that loop forever with no delay. */
export function manyFrameGif(w: number, h: number, frames: number): Uint8Array {
  const out: number[] = [...strToU8('GIF89a'), w & 0xff, w >> 8, h & 0xff, h >> 8, 0x80, 0, 0];
  out.push(0, 0, 0, 255, 255, 255); // two-color global table
  out.push(0x21, 0xff, 11, ...strToU8('NETSCAPE2.0'), 3, 1, 0, 0, 0);
  for (const i of range(frames)) {
    out.push(0x21, 0xf9, 4, 0, 0, 0, 0, 0); // graphic control: no delay
    const x = i % w;
    const y = Math.floor(i / w) % h;
    out.push(0x2c, x & 0xff, x >> 8, y & 0xff, y >> 8, 1, 0, 1, 0, 0);
    out.push(2, 2, 0x44, 0x01, 0); // LZW: clear, 0, end
  }
  out.push(0x3b);
  return new Uint8Array(out);
}

/** An EMF record: type, size in bytes, then 32-bit words. */
export const emfRecord = (type: number, words: readonly number[]): number[] => [
  type,
  8 + words.length * 4,
  ...words,
];

/** An EMF with `records` after its header, bounds `w`×`h` device pixels. */
export function emf(
  records: readonly number[][],
  w = 200,
  h = 100,
  extraBytes?: Uint8Array
): Uint8Array {
  const body = [...records, emfRecord(14, [0, 16, 20])];
  const words = body.flat();
  const right = Math.min(0x7fffffff, Math.round((w * 2540) / 96));
  const bottom = Math.min(0x7fffffff, Math.round((h * 2540) / 96));
  const header = [1, 88, 0, 0, w - 1, h - 1, 0, 0, right, bottom, 0x464d4520, 0x10000];
  header.push((22 + words.length) * 4 + (extraBytes?.length ?? 0), body.length + 1, 2, 0, 0, 0);
  header.push(1920, 1080, 508, 286);
  const out = new Uint8Array((header.length + words.length) * 4);
  const view = new DataView(out.buffer);
  [...header, ...words].forEach((x, i) => view.setInt32(i * 4, x | 0, true));
  return extraBytes
    ? concat([out.subarray(0, out.length - 20), extraBytes, out.subarray(-20)])
    : out;
}

/** `n` EMF rectangle records with a brush: a metafile that is only long. */
export function emfLong(n: number): Uint8Array {
  const records = [emfRecord(39, [1, 0, 0x0000ff, 0]), emfRecord(37, [1])];
  for (const i of range(n))
    records.push(emfRecord(43, [i % 190, i % 90, (i % 190) + 10, (i % 90) + 10]));
  return emf(records);
}

/**
 * An EMF whose one `EMR_STRETCHDIBITS` record carries a bitmap header of `w`×`h` 32-bit pixels
 * with a few bytes of pixel data.
 */
export function emfBigBitmap(w: number, h: number): Uint8Array {
  const bmi = [40, w, h, 1 | (32 << 16), 0, 0, 0, 0, 0, 0];
  const bits = [0, 0, 0, 0];
  // Bounds, dest x, y, src x, y, cx, cy, offBmi, cbBmi, offBits, cbBits, usage, rop, dest cx, cy.
  const fixed = 8 + 16 + 4 * 15;
  const words = [
    0,
    0,
    199,
    99,
    0,
    0,
    0,
    0,
    w,
    h,
    fixed,
    40,
    fixed + 40,
    16,
    0,
    0x00cc0020,
    200,
    100,
  ];
  return emf([emfRecord(81, [...words, ...bmi, ...bits])]);
}

// --- SVG pictures ------------------------------------------------------------------------------

export const svgDoc = (w: number, h: number, body: string, extra = ''): Uint8Array =>
  strToU8(
    `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" ` +
      `width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"${extra}>${body}</svg>`
  );

/** Ten levels of `<use>`, each ten copies of the level below: 10^10 rectangles. */
export const svgUseBomb = (): Uint8Array => {
  let body = '<defs><rect id="l0" width="10" height="10" fill="#4472c4"/>';
  for (let i = 1; i <= 10; i += 1)
    body += `<g id="l${i}">${range(10)
      .map((k) => `<use href="#l${i - 1}" xlink:href="#l${i - 1}" x="${k}"/>`)
      .join('')}</g>`;
  return svgDoc(400, 300, `${body}</defs><use href="#l10" xlink:href="#l10"/>`);
};

/** A `<use>` that refers to its own ancestor. */
export const svgUseCycle = (): Uint8Array =>
  svgDoc(
    400,
    300,
    '<g id="a"><rect width="10" height="10"/><use href="#a" xlink:href="#a" x="5"/></g>'
  );

export const svgBlurHuge = (): Uint8Array =>
  svgDoc(
    20000,
    20000,
    '<defs><filter id="f" x="-50%" y="-50%" width="200%" height="200%" filterUnits="userSpaceOnUse">' +
      '<feGaussianBlur stdDeviation="5000"/><feMorphology operator="dilate" radius="5000"/>' +
      '<feTurbulence baseFrequency="0.9" numOctaves="100"/></filter></defs>' +
      range(200)
        .map(
          (i) => `<rect x="${i * 90}" y="${i * 90}" width="4000" height="4000" filter="url(#f)"/>`
        )
        .join('')
  );

export const svgDeep = (depth: number): Uint8Array =>
  svgDoc(
    400,
    300,
    `${'<g transform="translate(0.001)">'.repeat(depth)}<rect width="9" height="9"/>${'</g>'.repeat(depth)}`
  );

export const svgExternal = (): Uint8Array =>
  svgDoc(
    400,
    300,
    '<image href="http://127.0.0.1:9/leak.png" xlink:href="http://127.0.0.1:9/leak.png" width="100" height="100"/>' +
      '<use href="http://127.0.0.1:9/leak.svg#x"/><style>@import url(http://127.0.0.1:9/leak.css);' +
      ' rect { fill: url(http://127.0.0.1:9/leak.svg#p) }</style><rect width="50" height="50"/>' +
      '<foreignObject width="100" height="100"><iframe xmlns="http://www.w3.org/1999/xhtml" src="http://127.0.0.1:9/x"/></foreignObject>' +
      '<script>for(;;){}</script>'
  );

export const svgHugeSize = (): Uint8Array =>
  svgDoc(
    1_000_000_000,
    1_000_000_000,
    '<rect width="1000000000" height="1000000000" fill="#ed7d31"/>'
  );

/** An entity expansion inside an SVG part: an XML reader that expands entities explodes. */
export const svgEntities = (): Uint8Array => {
  let dtd = '<!DOCTYPE svg [<!ENTITY a0 "lol">';
  for (let i = 1; i <= 10; i += 1) dtd += `<!ENTITY a${i} "${`&a${i - 1};`.repeat(10)}">`;
  return strToU8(
    `${dtd}]><svg xmlns="http://www.w3.org/2000/svg" width="400" height="300"><text y="20">&a10;</text></svg>`
  );
};

// --- Slides with pictures ----------------------------------------------------------------------

export const pic = (
  id: number,
  rel: string,
  x: number,
  y: number,
  cx: number,
  cy: number,
  blip = '',
  fill = '<a:stretch><a:fillRect/></a:stretch>'
) =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="${rel}">${blip}</a:blip>${fill}</p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="${Math.round(x)}" y="${Math.round(y)}"/><a:ext cx="${Math.round(cx)}" cy="${Math.round(cy)}"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';

interface Media {
  readonly name: string;
  readonly bytes: Uint8Array;
  readonly contentType: string;
}

/** One slide of `shapes` with picture relationships `rIdM<i>` to `media`. */
export function pictureDeck(
  shapes: string,
  media: readonly Media[],
  extraRels: [string, string, string][] = []
): Uint8Array {
  return buildPptx({
    slides: [
      {
        shapes,
        rels: [
          ...media.map((m, i): [string, string, string] => [
            `rIdM${i}`,
            R_IMAGE,
            `../media/${m.name}`,
          ]),
          ...extraRels,
        ],
      },
    ],
    media,
  });
}

export const png = (name: string, bytes: Uint8Array): Media => ({
  name,
  bytes,
  contentType: 'image/png',
});

/** One full-slide picture of `media`, with `blip` children (`a:duotone`, `a:alphaModFix`, …). */
export const fullPicture = (m: Media, blip = '', fill?: string) =>
  pictureDeck(pic(2, 'rIdM0', 0, 0, 12 * IN, 6.75 * IN, blip, fill), [m]);

/** A picture with an SVG alternative: the PNG `rIdM0` and the SVG `rIdSvg`. */
export function svgPictureDeck(svg: Uint8Array): Uint8Array {
  const blip = `<a:extLst><a:ext uri="${SVG_EXT}"><asvg:svgBlip xmlns:asvg="${ASVG}" r:embed="rIdSvg"/></a:ext></a:extLst>`;
  return buildPptx({
    slides: [
      {
        shapes: pic(2, 'rIdM0', IN, IN, 8 * IN, 5 * IN, blip),
        rels: [
          ['rIdM0', R_IMAGE, '../media/image0.png'],
          ['rIdSvg', R_IMAGE, '../media/image1.svg'],
        ],
      },
    ],
    media: [
      png('image0.png', bitPng(64, 48)),
      { name: 'image1.svg', bytes: svg, contentType: 'image/svg+xml' },
    ],
  });
}

export const metafileDeck = (bytes: Uint8Array): Uint8Array =>
  pictureDeck(pic(2, 'rIdM0', IN, IN, 8 * IN, 4 * IN), [
    { name: 'image0.emf', bytes, contentType: 'image/x-emf' },
  ]);

/** The deck's parts zipped with media stored, not deflated: the zip limits never refuse it. */
export function stored(bytes: Uint8Array): Uint8Array {
  const files = unzipSync(bytes);
  const z: Zippable = {};
  for (const [name, data] of Object.entries(files))
    z[name] = name.startsWith('ppt/media/') ? [data, { level: 0 }] : data;
  return zipSync(z);
}

// --- Effects -----------------------------------------------------------------------------------

export const effectShape = (
  id: number,
  effects: string,
  o: Partial<Parameters<typeof shape>[0]> = {}
) =>
  shape({
    id,
    x: 2 * IN,
    y: 1.5 * IN,
    cx: 8 * IN,
    cy: 4 * IN,
    fill: solid('4472C4'),
    effects: `<a:effectLst>${effects}</a:effectLst>`,
    ...o,
  });

export const MAX_RAD = 2_147_483_647;
export const clr = '<a:srgbClr val="FF0000"><a:alpha val="60000"/></a:srgbClr>';

// --- Charts ------------------------------------------------------------------------------------

export const numCache = (values: readonly (number | string)[]) =>
  `<c:numRef><c:f>Sheet1!$B$2</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values
    .map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)
    .join('')}</c:numCache></c:numRef>`;
export const strCache = (values: readonly string[]) =>
  `<c:strRef><c:f>Sheet1!$A$2</c:f><c:strCache><c:ptCount val="${values.length}"/>${values
    .map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)
    .join('')}</c:strCache></c:strRef>`;

export const ser = (
  i: number,
  values: readonly (number | string)[],
  cats?: readonly string[],
  extra = ''
) =>
  `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${extra}${cats ? `<c:cat>${strCache(cats)}</c:cat>` : ''}<c:val>${numCache(values)}</c:val></c:ser>`;

export const axes = (scaling = '<c:orientation val="minMax"/>') =>
  '<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
  `<c:valAx><c:axId val="2"/><c:scaling>${scaling}</c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx>`;

export const chartXml = (plot: string, legend = true) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${C_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R_NS}">` +
  `<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>${plot}</c:plotArea>${legend ? '<c:legend><c:legendPos val="r"/></c:legend>' : ''}<c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;

export const lineChart = (series: string, scaling?: string, kind = 'lineChart') =>
  `<c:${kind}>${kind === 'barChart' ? '<c:barDir val="col"/>' : ''}<c:grouping val="${kind === 'barChart' ? 'clustered' : 'standard'}"/><c:varyColors val="0"/>${series}<c:axId val="1"/><c:axId val="2"/></c:${kind}>${axes(scaling)}`;

/** A slide with `n` chart frames, each to its own part `charts[i % charts.length]`. */
export function chartDeck(
  charts: readonly string[],
  frames = 1,
  rels?: (i: number) => string
): Uint8Array {
  const shapes = range(frames)
    .map(
      (i) =>
        `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${i + 2}" name="Chart ${i}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>` +
        `<p:xfrm><a:off x="${IN + (i % 10) * 1000}" y="${IN}"/><a:ext cx="${10 * IN}" cy="${5 * IN}"/></p:xfrm>` +
        `<a:graphic><a:graphicData uri="${C_NS}"><c:chart xmlns:c="${C_NS}" r:id="rIdC${i % charts.length}"/></a:graphicData></a:graphic></p:graphicFrame>`
    )
    .join('');
  return buildPptx({
    slides: [
      {
        shapes,
        rels: charts.map((_, i): [string, string, string] => [
          `rIdC${i}`,
          `${R_NS}/chart`,
          rels?.(i) ?? `../charts/chart${i + 1}.xml`,
        ]),
      },
    ],
    extraParts: charts.map((xml, i): [string, string, string] => [
      `ppt/charts/chart${i + 1}.xml`,
      CHART_CT,
      xml,
    ]),
  });
}

// --- Text --------------------------------------------------------------------------------------

export const NOFIT = '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>';
export const ZWJ_FAMILY = '\u{1F468}\u200D\u{1F469}\u200D\u{1F467}\u200D\u{1F466}';
export const MIXED = `Hello مرحبا שלום 你好 नमस्ते ${ZWJ_FAMILY}\u{1F3F3}\uFE0F\u200D\u{1F308} สวัสดี `;

// --- Package -----------------------------------------------------------------------------------

/** A deck with `n` masters, each with its own theme of `themeBytes` filler. */
export function manyMastersDeck(n: number, filler: number): Uint8Array {
  return patched(one(box(2, para(run('masters')))), (f) => {
    const master = f.get('ppt/slideMasters/slideMaster1.xml');
    const masterRels = f.get('ppt/slideMasters/_rels/slideMaster1.xml.rels');
    const theme = f.get('ppt/theme/theme1.xml');
    const pad = `<a:extLst><a:ext uri="{PAD}"><x:pad xmlns:x="urn:pad">${'x'.repeat(filler)}</x:pad></a:ext></a:extLst>`;
    const bigTheme = theme.replace('</a:theme>', `${pad}</a:theme>`);
    let types = f.get('[Content_Types].xml');
    let presRels = f.get('ppt/_rels/presentation.xml.rels');
    let pres = f.get('ppt/presentation.xml');
    const ids: string[] = [];
    for (const i of range(n)) {
      const m = i + 2;
      f.set(
        `ppt/slideMasters/slideMaster${m}.xml`,
        master.replace(/<p:sldLayoutIdLst>[\s\S]*?<\/p:sldLayoutIdLst>/, '')
      );
      f.set(
        `ppt/slideMasters/_rels/slideMaster${m}.xml.rels`,
        masterRels
          .replace(/<Relationship [^>]*slideLayout"[^>]*\/>/g, '')
          .replace(/theme\/theme1\.xml/, `theme/theme${m + 1}.xml`)
      );
      f.set(`ppt/theme/theme${m + 1}.xml`, bigTheme);
      types = types.replace(
        '</Types>',
        `<Override PartName="/ppt/slideMasters/slideMaster${m}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>` +
          `<Override PartName="/ppt/theme/theme${m + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`
      );
      presRels = presRels.replace(
        '</Relationships>',
        `<Relationship Id="rIdMaster${m}" Type="${R_NS}/slideMaster" Target="slideMasters/slideMaster${m}.xml"/></Relationships>`
      );
      ids.push(`<p:sldMasterId id="${2147483648 + 100 * m}" r:id="rIdMaster${m}"/>`);
    }
    pres = pres.replace('</p:sldMasterIdLst>', `${ids.join('')}</p:sldMasterIdLst>`);
    f.set('[Content_Types].xml', types);
    f.set('ppt/_rels/presentation.xml.rels', presRels);
    f.set('ppt/presentation.xml', pres);
  });
}

/** `n` hyperlinks with external targets, and a picture linked (not embedded) from a URL. */
export function externalLinksDeck(n: number): Uint8Array {
  const runs = range(n)
    .map((i) => run(`link ${i} `, 'sz="600"', `<a:hlinkClick r:id="rIdL${i}"/>`))
    .join('');
  const linked =
    '<p:pic><p:nvPicPr><p:cNvPr id="3" name="Linked"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:link="rIdLinked"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  return buildPptx({
    slides: [
      {
        shapes: box(2, para(runs), { bodyPr: NOFIT }) + linked,
        rels: [
          ...range(n).map((i): [string, string, string, 'External'] => [
            `rIdL${i}`,
            `${R_NS}/hyperlink`,
            `http://127.0.0.1:9/l${i}`,
            'External',
          ]),
          ['rIdLinked', R_IMAGE, 'http://127.0.0.1:9/linked.png', 'External'],
        ],
      },
    ],
  });
}

/** Video and audio frames whose media are external URLs, and one embedded clip. */
export function mediaDeck(): Uint8Array {
  const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
  const frame = (id: number, kind: 'videoFile' | 'audioFile', link: string, media: string) =>
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Media ${id}"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr/>` +
    `<p:nvPr><a:${kind} r:link="${link}"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:p14="${P14}" r:embed="${media}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="rIdM0"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${id * IN}" y="${IN}"/><a:ext cx="${IN}" cy="${IN}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const timing =
    '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>' +
    '<p:video><p:cMediaNode vol="80000"><p:cTn id="2" fill="hold" display="0"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cMediaNode></p:video>' +
    '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>';
  return buildPptx({
    slides: [
      {
        shapes:
          frame(2, 'videoFile', 'rIdV', 'rIdV') +
          frame(3, 'audioFile', 'rIdA', 'rIdA') +
          frame(4, 'videoFile', 'rIdE', 'rIdE'),
        rels: [
          ['rIdM0', R_IMAGE, '../media/image0.png'],
          ['rIdV', `${R_NS}/video`, 'http://127.0.0.1:9/movie.mp4', 'External'],
          ['rIdA', `${R_NS}/audio`, 'http://127.0.0.1:9/sound.mp3', 'External'],
          ['rIdE', `${R_NS}/video`, '../media/media1.mp4'],
        ],
        ext: timing,
      },
    ],
    media: [
      png('image0.png', bitPng(16, 16)),
      {
        name: 'media1.mp4',
        bytes: new Uint8Array(5 * 1024 * 1024).fill(7),
        contentType: 'video/mp4',
      },
    ],
  });
}

/** `n` nested `mc:AlternateContent` levels around one shape. */
export function nestedAlternateContent(n: number): string {
  const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
  let inner = box(2, para(run('alternate')));
  for (let i = 0; i < n; i += 1)
    inner = `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice Requires="p">${inner}</mc:Choice><mc:Fallback>${inner.length < 2000 ? inner : ''}</mc:Fallback></mc:AlternateContent>`;
  return inner;
}
