// Namespace URIs and the qualified-name shorthand the engine uses to address elements.
//
// The tree stores EXPANDED names (namespace URI + local name). Prefixes are fidelity evidence
// only: a file may bind DrawingML to `a:` or to `drawing:`, and both mean the same element.
// Code addresses elements through `q('a:solidFill')`, which maps the conventional prefix to
// its URI once, so a lookup never depends on the prefix a file chose.

export const NS = Object.freeze({
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
  p15: 'http://schemas.microsoft.com/office/powerpoint/2012/main',
  a14: 'http://schemas.microsoft.com/office/drawing/2010/main',
  a16: 'http://schemas.microsoft.com/office/drawing/2014/main',
  asvg: 'http://schemas.microsoft.com/office/drawing/2016/SVG/main',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  dgm: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
  dsp: 'http://schemas.microsoft.com/office/drawing/2008/diagram',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  xml: 'http://www.w3.org/XML/1998/namespace',
  xmlns: 'http://www.w3.org/2000/xmlns/',
});

/** A conventional prefix the engine knows. */
export type KnownPrefix = keyof typeof NS;

/** An expanded element or attribute name. */
export interface QName {
  readonly ns: string;
  readonly local: string;
}

const cache = new Map<string, QName>();

/**
 * The expanded name for a conventional `prefix:local` spelling. Unprefixed means no namespace
 * (the usual case for attributes). Throws on an unknown prefix: this is for engine code, never
 * for file data.
 */
export function q(name: string): QName {
  const hit = cache.get(name);
  if (hit) return hit;
  const colon = name.indexOf(':');
  let out: QName;
  if (colon < 0) out = { ns: '', local: name };
  else {
    const prefix = name.slice(0, colon) as KnownPrefix;
    const ns = NS[prefix];
    if (ns === undefined) throw new Error(`unknown engine prefix ${prefix}`);
    out = { ns, local: name.slice(colon + 1) };
  }
  cache.set(name, out);
  return out;
}

/** The conventional prefix for a namespace URI, used when the engine creates nodes. */
export function conventionalPrefix(ns: string): string | undefined {
  for (const [prefix, uri] of Object.entries(NS)) if (uri === ns) return prefix;
  return undefined;
}

/** Relationship types the engine follows. */
export const REL = Object.freeze({
  officeDocument:
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  notesSlide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
  notesMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  font: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
  diagramDrawing: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing',
  tableStyles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles',
  presProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps',
  viewProps: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps',
});

/** Content types the engine writes. */
export const CONTENT_TYPE = Object.freeze({
  presentation:
    'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  notesSlide: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
  relationships: 'application/vnd.openxmlformats-package.relationships+xml',
});
