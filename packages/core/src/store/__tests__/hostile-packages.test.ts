// Packages whose content types disagree with their relationships still open: the main part of
// a slide show, a template, or a macro-enabled deck reads as a presentation; a slide with a
// generic content type is still a slide; and a part that nothing structural points to never
// refuses the deck, whatever its content type says.

import { describe, expect, test } from 'bun:test';
import { readZip, writeZip } from '@eigenpal/ooxml-core';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { presentationInfo, readPptx } from '../index.ts';

const CT = '/[Content_Types].xml';
const PML = 'application/vnd.openxmlformats-officedocument.presentationml.';

const deck = () =>
  buildPptx({
    slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 1, cy: 1, text: para(run('A')) }) }],
  });

/** `deck()` with its content types (and other entries) changed by `edit`. */
function patched(edit: (types: string, entries: Map<string, Uint8Array>) => string): Uint8Array {
  const zip = readZip(deck());
  if (!zip.ok) throw new Error('zip');
  const entries = new Map(zip.entries);
  const types = new TextDecoder().decode(entries.get(CT));
  entries.set(CT, new TextEncoder().encode(edit(types, entries)));
  return writeZip(entries);
}

const slides = (bytes: Uint8Array): number => {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error(`refused: ${r.error.reason}`);
  return presentationInfo(r.pkg).slides.length;
};

describe('hostile packages', () => {
  test.each([
    `${PML}slideshow.main+xml`,
    `${PML}template.main+xml`,
    'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml',
  ])('a main part of type %s opens', (type) => {
    expect(slides(patched((t) => t.replace(`${PML}presentation.main+xml`, type)))).toBe(1);
  });

  test('a slide with no content type override is still a slide', () => {
    const bytes = patched((t) =>
      t.replace(/<Override PartName="\/ppt\/slides\/slide1\.xml"[^>]*\/>/, '')
    );
    expect(slides(bytes)).toBe(1);
  });

  test('a part that nothing points to never refuses the deck', () => {
    const bytes = patched((t, entries) => {
      entries.set('/ppt/media/stray.xml', new Uint8Array([0x89, 0x50, 0x4e, 0x47]));
      return t.replace(
        '</Types>',
        '<Override PartName="/ppt/media/stray.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>'
      );
    });
    expect(slides(bytes)).toBe(1);
  });

  test('a slide entry that names the master or a layout is not a slide', () => {
    const zip = readZip(deck());
    if (!zip.ok) throw new Error('zip');
    const entries = new Map(zip.entries);
    const name = '/ppt/presentation.xml';
    const rels = new TextDecoder().decode(entries.get('/ppt/_rels/presentation.xml.rels'));
    const master = /Id="(rId\d+)"[^>]*slideMaster"/.exec(rels)?.[1];
    expect(master).toBeDefined();
    const xml = new TextDecoder().decode(entries.get(name));
    const slideRef = /<p:sldId [^>]*r:id="(rId\d+)"/.exec(xml)![1]!;
    entries.set(
      name,
      new TextEncoder().encode(xml.replace(`r:id="${slideRef}"`, `r:id="${master}"`))
    );
    expect(slides(writeZip(entries))).toBe(0);
  });

  test('a slide with a relationship to a missing part duplicates', async () => {
    const { createPptxEditor } = await import('../../editor/index.ts');
    const bytes = patched((t, entries) => {
      const rels = '/ppt/slides/_rels/slide1.xml.rels';
      const xml = new TextDecoder().decode(entries.get(rels));
      const missing =
        '<Relationship Id="rIdGone" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/gone.png"/>';
      entries.set(
        rels,
        new TextEncoder().encode(xml.replace('</Relationships>', `${missing}</Relationships>`))
      );
      return t;
    });
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(bytes)).ok).toBe(true);
    expect(editor.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    expect(editor.snapshot().slides.length).toBe(2);
    editor.destroy();
  });
});
