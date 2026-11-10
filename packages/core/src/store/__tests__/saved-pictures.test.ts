import { expect, test } from 'bun:test';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  readPptx,
  writePptx,
  canonicalPackageFingerprint,
  serializeXmlPart,
  xmlPartOf,
} from '../index.ts';

const image = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const relNs = 'http://schemas.openxmlformats.org/package/2006/relationships';
const bytes = new Uint8Array([1, 2, 3]);
const media = [{ name: 'unused.png', contentType: 'image/png', bytes }];
const owner = (extra = '') => shape({ id: 2, x: 0, y: 0, cx: 914400, cy: 914400 }) + extra;
function save(input: Uint8Array) {
  const opened = readPptx(input);
  if (!opened.ok) throw new Error('fixture refused');
  const before = canonicalPackageFingerprint(opened.pkg);
  const saved = writePptx(opened.pkg);
  expect(canonicalPackageFingerprint(opened.pkg)).toBe(before);
  const result = readPptx(saved);
  if (!result.ok) throw new Error('save refused');
  expect(writePptx(result.pkg)).toEqual(saved);
  return result.pkg;
}

test('save removes an unused image relationship, image bytes, and content type override', () => {
  const pkg = save(
    buildPptx({
      slides: [
        { layout: 'blank', shapes: owner(), rels: [['unused', image, '../media/unused.png']] },
      ],
      media,
      extraParts: [['ppt/media/unused.png', 'image/png', '']],
    })
  );
  expect(pkg.parts.has('/ppt/media/unused.png')).toBe(false);
  expect(serializeXmlPart(xmlPartOf(pkg, '/ppt/slides/_rels/slide1.xml.rels')!)).not.toContain(
    'unused'
  );
  expect(serializeXmlPart(xmlPartOf(pkg, '/[Content_Types].xml')!)).not.toContain(
    '/ppt/media/unused.png'
  );
});

for (const extra of [
  '<p:extLst><p:ext uri="urn:test"><x:asset xmlns:x="urn:test" asset="unused"/></p:ext></p:extLst>',
  '<p:extLst><p:ext uri="urn:test"><x:asset xmlns:x="urn:test" assets="other unused"/></p:ext></p:extLst>',
])
  test('save keeps image references in unknown extension attributes', () => {
    const pkg = save(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: owner(extra),
            rels: [['unused', image, '../media/unused.png']],
          },
        ],
        media,
      })
    );
    expect(pkg.parts.get('/ppt/media/unused.png')).toMatchObject({ bytes });
  });

for (const opaque of [false, true])
  test(`save preserves images through ${opaque ? 'opaque owners' : 'unknown relationship types'}`, () => {
    const pkg = save(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: owner(),
            rels: opaque ? [] : [['unused', 'urn:test:asset', '../media/unused.png']],
          },
        ],
        media,
        extraParts: opaque
          ? [
              ['ppt/custom.xml', 'application/xml', '<custom/>'],
              [
                'ppt/_rels/custom.xml.rels',
                'application/vnd.openxmlformats-package.relationships+xml',
                `<Relationships xmlns="${relNs}"><Relationship Id="opaque" Type="${image}" Target="media/unused.png"/></Relationships>`,
              ],
            ]
          : [],
      })
    );
    expect(pkg.parts.get('/ppt/media/unused.png')).toMatchObject({ bytes });
  });

test('an unresolved relationship disables image cleanup', () => {
  const pkg = save(
    buildPptx({
      slides: [
        { layout: 'blank', shapes: owner(), rels: [['bad', 'urn:test:asset', '../../../../bad']] },
      ],
      media,
    })
  );
  expect(pkg.parts.has('/ppt/media/unused.png')).toBe(true);
});

test('save keeps an image used by another slide', () => {
  const pkg = save(
    buildPptx({
      slides: [
        { layout: 'blank', shapes: owner(), rels: [['unused', image, '../media/unused.png']] },
        {
          layout: 'blank',
          shapes: owner('<p:extLst><p:ext uri="urn:test" r:id="shared"/></p:extLst>'),
          rels: [['shared', image, '../media/unused.png']],
        },
      ],
      media,
    })
  );
  expect(pkg.parts.has('/ppt/media/unused.png')).toBe(true);
});

for (const [folder, root] of [
  ['slideLayouts', 'sldLayout'],
  ['slideMasters', 'sldMaster'],
  ['notesSlides', 'notes'],
  ['notesMasters', 'notesMaster'],
] as const) {
  test(`save keeps an image used by ${root}`, () => {
    const name = `ppt/${folder}/custom.xml`;
    const pkg = save(
      buildPptx({
        slides: [
          { layout: 'blank', shapes: owner(), rels: [['unused', image, '../media/unused.png']] },
        ],
        media,
        extraParts: [
          [
            name,
            'application/xml',
            `<p:${root} xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:extLst><p:ext uri="urn:test" r:id="shared"/></p:extLst></p:${root}>`,
          ],
          [
            `ppt/${folder}/_rels/custom.xml.rels`,
            'application/vnd.openxmlformats-package.relationships+xml',
            `<Relationships xmlns="${relNs}"><Relationship Id="shared" Type="${image}" Target="../media/unused.png"/></Relationships>`,
          ],
        ],
      })
    );
    expect(pkg.parts.get('/ppt/media/unused.png')).toMatchObject({ bytes });
  });
}

test('an unreadable relationship list disables image cleanup', () => {
  const pkg = save(
    buildPptx({
      slides: [{ layout: 'blank', shapes: owner() }],
      media,
      extraParts: [
        [
          'ppt/_rels/custom.xml.rels',
          'application/vnd.openxmlformats-package.relationships+xml',
          '<broken',
        ],
      ],
    })
  );
  expect(pkg.parts.has('/ppt/media/unused.png')).toBe(true);
});

test('save removes percent-encoded content type overrides for unused images', () => {
  const files = unzipSync(buildPptx({ slides: [{ layout: 'blank', shapes: owner() }], media }));
  files['[Content_Types].xml'] = strToU8(
    strFromU8(files['[Content_Types].xml']).replace(
      '</Types>',
      '<Override PartName="/ppt/media/%75nused.png" ContentType="image/png"/></Types>'
    )
  );
  const pkg = save(zipSync(files));
  expect(pkg.parts.has('/ppt/media/unused.png')).toBe(false);
  expect(serializeXmlPart(xmlPartOf(pkg, '/[Content_Types].xml')!)).not.toContain('%75nused');
});
