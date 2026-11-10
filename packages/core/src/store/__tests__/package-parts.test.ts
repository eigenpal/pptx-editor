// Regression tests for audit finding K2 and the content-type lookup: which
// parts the reader parses, which it carries as bytes, and which must parse for a deck to open.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  canonicalPackageFingerprint,
  contentTypeOf,
  equivalentPartName,
  presentationInfo,
  readPptx,
  writePptx,
  type PptxPackage,
} from '../index.ts';

const SVG_TYPE = 'image/svg+xml';

function deckFiles(): Record<string, Uint8Array> {
  return unzipSync(
    buildPptx({ slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100 }) }] })
  );
}

function withPart(
  files: Record<string, Uint8Array>,
  name: string,
  contentType: string,
  bytes: Uint8Array
): Uint8Array {
  if (contentType.startsWith('image/')) {
    files['ppt/slides/slide1.xml'] = strToU8(
      strFromU8(files['ppt/slides/slide1.xml']).replace(
        '</p:sld>',
        '<p:extLst><p:ext uri="urn:test" r:id="rIdMedia"/></p:extLst></p:sld>'
      )
    );
    files['ppt/slides/_rels/slide1.xml.rels'] = strToU8(
      strFromU8(files['ppt/slides/_rels/slide1.xml.rels']).replace(
        '</Relationships>',
        `<Relationship Id="rIdMedia" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/${name}"/></Relationships>`
      )
    );
  }
  const ct = strFromU8(files['[Content_Types].xml']).replace(
    '</Types>',
    `<Override PartName="/${name}" ContentType="${contentType}"/></Types>`
  );
  return zipSync({ ...files, '[Content_Types].xml': strToU8(ct), [name]: bytes });
}

function open(bytes: Uint8Array): PptxPackage {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error(JSON.stringify(r.error));
  return r.pkg;
}

describe('K2: media and uninterpreted XML are carried as bytes', () => {
  test('an SVG with a standard DOCTYPE opens, and saves byte-identical', () => {
    const svg = strToU8(
      '<?xml version="1.0"?>\n<!DOCTYPE svg PUBLIC "-//W3C//DTD SVG 1.1//EN" ' +
        '"http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd">\n' +
        '<svg xmlns="http://www.w3.org/2000/svg"><rect width="1" height="1"/></svg>'
    );
    const pkg = open(withPart(deckFiles(), 'ppt/media/image1.svg', SVG_TYPE, svg));
    const part = pkg.parts.get('/ppt/media/image1.svg');
    expect(part?.kind).toBe('binary');
    const saved = unzipSync(writePptx(pkg))['ppt/media/image1.svg'];
    expect([...saved]).toEqual([...svg]);
  });

  test('an SVG keeps the spaces between text spans and its comments', () => {
    const text =
      '<svg xmlns="http://www.w3.org/2000/svg"><text x="0" y="10"><tspan>A</tspan> ' +
      '<tspan>B</tspan></text><!-- c --></svg>';
    const pkg = open(withPart(deckFiles(), 'ppt/media/image1.svg', SVG_TYPE, strToU8(text)));
    expect(strFromU8(unzipSync(writePptx(pkg))['ppt/media/image1.svg'])).toBe(text);
  });

  test('an XML part the engine does not read is kept as it is, even when it does not parse', () => {
    const broken = strToU8('<c:chartSpace xmlns:c="urn:c"><unclosed></c:chartSpace>');
    const pkg = open(
      withPart(
        deckFiles(),
        'ppt/charts/chart1.xml',
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
        broken
      )
    );
    expect(pkg.parts.get('/ppt/charts/chart1.xml')?.kind).toBe('binary');
    expect([...unzipSync(writePptx(pkg))['ppt/charts/chart1.xml']]).toEqual([...broken]);
  });

  test('an optional part the engine reads is kept as bytes when it does not parse', () => {
    const files = deckFiles();
    files['ppt/tableStyles.xml'] = strToU8('<!DOCTYPE x><a:tblStyleLst/>');
    const pkg = open(zipSync(files));
    expect(pkg.parts.get('/ppt/tableStyles.xml')?.kind).toBe('binary');
    expect(presentationInfo(pkg).slides).toHaveLength(1);
  });

  test('a slide that does not parse still refuses the deck', () => {
    const files = deckFiles();
    files['ppt/slides/slide1.xml'] = strToU8('<!DOCTYPE x><p:sld/>');
    const r = readPptx(zipSync(files));
    expect(r).toMatchObject({
      ok: false,
      error: { reason: 'xml', partName: '/ppt/slides/slide1.xml', detail: 'dtd-forbidden' },
    });
  });

  test('a theme that does not parse still refuses the deck', () => {
    const files = deckFiles();
    files['ppt/theme/theme1.xml'] = strToU8('<a:theme><unclosed></a:theme>');
    expect(readPptx(zipSync(files)).ok).toBe(false);
  });

  test('slides with non-ASCII element names open', () => {
    const files = deckFiles();
    const slide = strFromU8(files['ppt/slides/slide1.xml']).replace(
      '</p:cSld>',
      '</p:cSld><p:extLst><p:ext uri="{x}"><Größe xmlns="urn:x"/></p:ext></p:extLst>'
    );
    files['ppt/slides/slide1.xml'] = strToU8(slide);
    const pkg = open(zipSync(files));
    expect(canonicalPackageFingerprint(open(writePptx(pkg)))).toBe(
      canonicalPackageFingerprint(pkg)
    );
  });
});

describe('content types', () => {
  test('an Override part name is percent-decoded once, as the ZIP item name is', () => {
    const files = deckFiles();
    const ct = strFromU8(files['[Content_Types].xml']).replace(
      'PartName="/ppt/presentation.xml"',
      'PartName="/ppt/%70resentation.xml"'
    );
    files['[Content_Types].xml'] = strToU8(ct);
    const pkg = open(zipSync(files));
    expect(contentTypeOf(pkg, '/ppt/presentation.xml')).toBe(
      'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
    );
  });

  test('equivalentPartName finds a part whatever the ASCII case', () => {
    const pkg = open(buildPptx({ slides: [{ shapes: '' }] }));
    expect(equivalentPartName(pkg, '/PPT/Presentation.XML')).toBe('/ppt/presentation.xml');
    expect(equivalentPartName(pkg, '/ppt/missing.xml')).toBeUndefined();
  });
});
