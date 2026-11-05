import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { binaryPartOf, presentationInfo, readPptx } from '../../store/index.ts';
import { buildEot, buildMtx } from '../../store/fonts/__tests__/fixtures.ts';
import { embeddedFontSources } from '../document-session.ts';
import { createPptxEditor } from '../index.ts';
import { themedDeck } from './font-test-helpers.ts';

/** Only generated font bytes enter this fixture. No installed font permissions are changed. */
function deckWithFont(fsType: number) {
  const font = buildEot(buildMtx(), { flags: 4, fsType, family: 'Arial' });
  const files = unzipSync(themedDeck({ latin: 'Arial' }));
  const append = (part: string, end: string, xml: string) => {
    const text = strFromU8(files[part]!);
    expect(text).toContain(end);
    files[part] = strToU8(text.replace(end, xml + end));
  };
  append(
    'ppt/presentation.xml',
    '</p:presentation>',
    '<p:embeddedFontLst><p:embeddedFont><p:font typeface="Arial"/>' +
      '<p:regular r:id="rIdPermissionFont"/></p:embeddedFont></p:embeddedFontLst>'
  );
  append(
    'ppt/_rels/presentation.xml.rels',
    '</Relationships>',
    '<Relationship Id="rIdPermissionFont" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" ' +
      'Target="fonts/permission.fntdata"/>'
  );
  append(
    '[Content_Types].xml',
    '</Types>',
    '<Override PartName="/ppt/fonts/permission.fntdata" ContentType="application/x-fontdata"/>'
  );
  files['ppt/fonts/permission.fntdata'] = new Uint8Array(font);
  return { bytes: zipSync(files), font };
}

test.each([
  [0, 1], // installable
  [2, 0], // restricted
  [4, 0], // preview and print
  [8, 1], // editable
  [0x0208, 0], // bitmap-only, even with editable permission
])('embedded font permission %i admits %i faces', (fsType, count) => {
  const read = readPptx(deckWithFont(fsType).bytes);
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  const result = embeddedFontSources(read.pkg);
  expect(result.sources).toHaveLength(count);
  expect(result.refused).toHaveLength(1 - count);
});

test('preview-print fonts use substitutes and survive editing, save, and reopen unchanged', async () => {
  const { bytes, font } = deckWithFont(4);
  const editor = createPptxEditor();
  try {
    await editor.load(bytes);
    expect(editor.session!.fontReport.embeddedFaces).toBe(0);
    expect(editor.session!.fontReport.refusedEmbedded).toEqual([
      'Arial (regular): preview and print only',
    ]);
    expect(editor.session!.faces.some((face) => face.origin === 'embedded')).toBe(false);
    expect(editor.snapshot().substitutedFonts.map((entry) => entry.family)).toContain('Arial');
    expect(editor.exec({ type: 'addSlide' }).ok).toBe(true);
    const saved = editor.save()!;
    const read = readPptx(saved);
    if (!read.ok) throw new Error(JSON.stringify(read.error));
    expect(binaryPartOf(read.pkg, '/ppt/fonts/permission.fntdata')).toEqual(font);
    expect(presentationInfo(read.pkg).embeddedFonts[0]?.faces[0]?.partName).toBe(
      '/ppt/fonts/permission.fntdata'
    );
    expect(embeddedFontSources(read.pkg).sources).toHaveLength(0);
    await editor.load(saved);
    expect(editor.session!.fontReport.embeddedFaces).toBe(0);
    expect(editor.snapshot().substitutedFonts.map((entry) => entry.family)).toContain('Arial');
  } finally {
    editor.destroy();
  }
});
