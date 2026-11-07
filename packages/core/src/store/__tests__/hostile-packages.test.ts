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

  test('a slide with a content type that is not XML is still a slide', () => {
    const bytes = patched((t) =>
      t.replace(
        /(<Override PartName="\/ppt\/slides\/slide1\.xml" ContentType=")[^"]*"/,
        '$1application/octet-stream"'
      )
    );
    expect(slides(bytes)).toBe(1);
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

  test.each(['presProps', 'viewProps', 'tableStyles', 'presentation'])(
    'a slide entry that names the %s part is not a slide',
    (kind) => {
      const zip = readZip(deck());
      if (!zip.ok) throw new Error('zip');
      const entries = new Map(zip.entries);
      const relsName = '/ppt/_rels/presentation.xml.rels';
      const rels = new TextDecoder().decode(entries.get(relsName));
      const target = kind === 'presentation' ? 'presentation.xml' : `${kind}.xml`;
      if (kind !== 'presentation' && !entries.has(`/ppt/${target}`))
        entries.set(
          `/ppt/${target}`,
          new TextEncoder().encode(
            `<p:${kind} xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`
          )
        );
      entries.set(
        relsName,
        new TextEncoder().encode(
          rels.replace(
            /(<Relationship [^>]*Type="[^"]*\/slide"[^>]*Target=")[^"]*"/,
            `$1${target}"`
          )
        )
      );
      expect(slides(writeZip(entries))).toBe(0);
    }
  );

  test('a slide relationship that names a picture never refuses the deck', () => {
    const zip = readZip(deck());
    if (!zip.ok) throw new Error('zip');
    const entries = new Map(zip.entries);
    const relsName = '/ppt/_rels/presentation.xml.rels';
    const rels = new TextDecoder().decode(entries.get(relsName));
    entries.set('/ppt/media/a.png', new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0, 1, 2]));
    const types = new TextDecoder().decode(entries.get(CT));
    entries.set(
      CT,
      new TextEncoder().encode(
        types.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>')
      )
    );
    entries.set(
      relsName,
      new TextEncoder().encode(
        rels.replace(
          '</Relationships>',
          `<Relationship Id="rIdPic" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="media/a.png"/></Relationships>`
        )
      )
    );
    expect(slides(writeZip(entries))).toBe(1);
  });

  test('a slide relationship whose target differs in case still defers and opens', () => {
    const bytes = patched((t, entries) => {
      const relsName = '/ppt/_rels/presentation.xml.rels';
      const rels = new TextDecoder().decode(entries.get(relsName));
      entries.set(
        relsName,
        new TextEncoder().encode(rels.replace('slides/slide1.xml', 'SLIDES/Slide1.XML'))
      );
      return t;
    });
    expect(slides(bytes)).toBe(1);
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

describe('damaged notes parts and notes themes', () => {
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NOTES = '/ppt/notesSlides/notesSlide1.xml';
  const DAMAGED = new TextEncoder().encode('<p:notes><broken');

  /** A deck with speaker notes, its entries changed by `edit`. */
  function notesDeck(edit: (entries: Map<string, Uint8Array>) => void): Uint8Array {
    const zip = readZip(
      buildPptx({ slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 1, cy: 1 }), notes: 'N' }] })
    );
    if (!zip.ok) throw new Error('zip');
    const entries = new Map(zip.entries);
    edit(entries);
    return writeZip(entries);
  }

  /** Add relationship XML to the `.rels` part `name`. */
  function addRel(entries: Map<string, Uint8Array>, name: string, rel: string): void {
    const xml = new TextDecoder().decode(entries.get(name));
    entries.set(
      name,
      new TextEncoder().encode(xml.replace('</Relationships>', `${rel}</Relationships>`))
    );
  }

  test.each(['slideLayout', 'slideMaster', 'theme', 'slide'])(
    'a damaged notes part that a slide relationship of type %s also names drops the notes',
    (type) => {
      const bytes = notesDeck((entries) => {
        entries.set(NOTES, DAMAGED);
        addRel(
          entries,
          '/ppt/slides/_rels/slide1.xml.rels',
          `<Relationship Id="rIdHostile" Type="${R}/${type}" Target="../notesSlides/notesSlide1.xml"/>`
        );
      });
      const r = readPptx(bytes);
      if (!r.ok) throw new Error(`refused: ${r.error.reason}`);
      expect(presentationInfo(r.pkg).slides).toHaveLength(1);
      // The damaged part stays as its bytes, so the save keeps it.
      const part = r.pkg.parts.get(NOTES);
      expect(part?.kind).toBe('binary');
    }
  );

  test('a damaged theme of the notes master drops nothing the slides need', () => {
    const bytes = notesDeck((entries) => {
      entries.set('/ppt/theme/theme2.xml', new TextEncoder().encode('<a:theme><broken'));
    });
    expect(slides(bytes)).toBe(1);
  });

  test('a damaged theme of the handout master drops nothing the slides need', () => {
    const PML_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
    const bytes = notesDeck((entries) => {
      entries.set(
        '/ppt/handoutMasters/handoutMaster1.xml',
        new TextEncoder().encode(`<p:handoutMaster xmlns:p="${PML_NS}"><p:cSld/></p:handoutMaster>`)
      );
      entries.set(
        '/ppt/handoutMasters/_rels/handoutMaster1.xml.rels',
        new TextEncoder().encode(
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            `<Relationship Id="rId1" Type="${R}/theme" Target="../theme/theme3.xml"/></Relationships>`
        )
      );
      entries.set('/ppt/theme/theme3.xml', new TextEncoder().encode('<a:theme><broken'));
      addRel(
        entries,
        '/ppt/_rels/presentation.xml.rels',
        `<Relationship Id="rIdHandout" Type="${R}/handoutMaster" Target="handoutMasters/handoutMaster1.xml"/>`
      );
      const types = new TextDecoder().decode(entries.get(CT));
      entries.set(
        CT,
        new TextEncoder().encode(
          types.replace(
            '</Types>',
            `<Override PartName="/ppt/handoutMasters/handoutMaster1.xml" ContentType="${PML}handoutMaster+xml"/>` +
              '<Override PartName="/ppt/theme/theme3.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>'
          )
        )
      );
    });
    expect(slides(bytes)).toBe(1);
  });

  test('a damaged theme of the slide master still refuses the deck', () => {
    const bytes = notesDeck((entries) => {
      entries.set('/ppt/theme/theme1.xml', new TextEncoder().encode('<a:theme><broken'));
    });
    expect(readPptx(bytes).ok).toBe(false);
  });

  test.each([
    ['slide', '/ppt/slides/slide1.xml', 'notesSlide'],
    ['slide master', '/ppt/slideMasters/slideMaster1.xml', 'notesMaster'],
    ['slide layout', '/ppt/slideLayouts/slideLayout1.xml', 'notesSlide'],
    ['slide master theme', '/ppt/theme/theme1.xml', 'notesSlide'],
  ])('a damaged %s with a notes content type still refuses the deck', (_, part, notesType) => {
    const bytes = notesDeck((entries) => {
      entries.set(part, new TextEncoder().encode('<p:sld><broken'));
      const types = new TextDecoder().decode(entries.get(CT));
      const retyped = types.replace(
        new RegExp(`(PartName="${part}" ContentType=")[^"]*"`),
        `$1${PML}${notesType}+xml"`
      );
      expect(retyped).not.toBe(types);
      entries.set(CT, new TextEncoder().encode(retyped));
    });
    expect(readPptx(bytes).ok).toBe(false);
  });
});
