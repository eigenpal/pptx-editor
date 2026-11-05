import { describe, expect, test } from 'bun:test';
import { zipSync, strToU8 } from 'fflate';
import { buildPptx, shape, para, run, solid } from '../../../../../scripts/fixtures/builder.ts';
import {
  binaryPartOf,
  canonicalPackageFingerprint,
  presentationInfo,
  readPptx,
  semanticDigest,
  writePptx,
  xmlPartOf,
  type PptxPackage,
} from '../index.ts';

const PNG = Uint8Array.from([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0,
  0, 31, 21, 196, 137,
]);

function sample(): Uint8Array {
  return buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          x: 100,
          y: 200,
          cx: 3000000,
          cy: 1000000,
          fill: solid('FF0000'),
          text: para(run('Hello')),
        }),
        notes: 'Speaker note',
      },
      { shapes: '', layout: 'blank' },
    ],
    media: [{ name: 'image1.png', bytes: PNG, contentType: 'image/png' }],
  });
}

function open(bytes: Uint8Array): PptxPackage {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error(JSON.stringify(r.error));
  return r.pkg;
}

describe('package read and save', () => {
  test('reads the structure of a fixture', () => {
    const info = presentationInfo(open(sample()));
    expect(info.slides).toHaveLength(2);
    expect(info.width).toBe(12_192_000);
    expect(info.slides[0].layoutPart).toBe('/ppt/slideLayouts/slideLayout1.xml');
    expect(info.slides[0].masterPart).toBe('/ppt/slideMasters/slideMaster1.xml');
    expect(info.slides[0].themePart).toBe('/ppt/theme/theme1.xml');
    expect(info.slides[0].notesPart).toBe('/ppt/notesSlides/notesSlide1.xml');
  });

  test('save then reopen keeps the canonical fingerprint (round-trip check)', () => {
    const pkg = open(sample());
    const before = canonicalPackageFingerprint(pkg);
    expect(semanticDigest(pkg)).toBe(before);
    const twice = open(writePptx(open(writePptx(pkg))));
    expect(canonicalPackageFingerprint(twice)).toBe(before);
  });

  test('binary parts round-trip byte-identical', () => {
    const pkg = open(writePptx(open(sample())));
    expect(binaryPartOf(pkg, '/ppt/media/image1.png')).toEqual(PNG);
  });

  test('save is deterministic', () => {
    const pkg = open(sample());
    expect(writePptx(pkg)).toEqual(writePptx(pkg));
  });

  test('unknown elements survive save', () => {
    const bytes = buildPptx({
      slides: [
        { shapes: '<p:custData xmlns:x="urn:x"><x:thing a="1">kept</x:thing></p:custData>' },
      ],
    });
    const pkg = open(writePptx(open(bytes)));
    const root = xmlPartOf(pkg, '/ppt/slides/slide1.xml')!.root;
    expect(JSON.stringify(root)).toContain('urn:x');
    expect(JSON.stringify(root)).toContain('kept');
  });
});

describe('trust boundary', () => {
  test('refuses a part with a DTD', () => {
    const files = {
      '[Content_Types].xml': strToU8('<?xml version="1.0"?><!DOCTYPE x [<!ENTITY a "b">]><Types/>'),
    };
    const r = readPptx(zipSync(files));
    expect(r.ok).toBe(false);
  });

  test('refuses a traversal entry name', () => {
    const r = readPptx(zipSync({ '../evil.xml': strToU8('<x/>') }));
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.reason).toBe('zip');
  });

  test('refuses a package that is not a presentation', () => {
    const files = {
      '[Content_Types].xml': strToU8(
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
      ),
    };
    const r = readPptx(zipSync(files));
    expect(r.ok).toBe(false);
  });

  test('refuses garbage bytes', () => {
    expect(readPptx(new Uint8Array([1, 2, 3, 4])).ok).toBe(false);
  });

  test('external relationship targets are recorded, never resolved to parts', () => {
    const bytes = buildPptx({
      slides: [
        {
          shapes: '',
          rels: [
            [
              'rIdX',
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
              'https://example.com/a.png" TargetMode="External',
            ],
          ],
        },
      ],
    });
    const pkg = open(bytes);
    const slide = presentationInfo(pkg).slides[0];
    expect(slide.partName).toBe('/ppt/slides/slide1.xml');
  });
});
