import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { attr, canonicalForm, elements, find, textOf, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const part = '/ppt/slides/slide1.xml';
const root = (e: PptxEditorInstance) => xmlPartOf(e.session!.store.package, part)!.root;
const shapes = (e: PptxEditorInstance) =>
  elements(find(root(e), 'p:spTree')!).filter((el) => el.kind === 'sp');
async function open(bytes: Uint8Array) {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

test('edits undo, redo, and reopen without changing unrelated package content', async () => {
  // Extend an existing reference-backed deck in memory; this is a preservation check, not a fidelity fixture.
  const files = unzipSync(readFileSync('fixtures/text-runs.pptx'));
  files['ppt/slides/slide1.xml'] = strToU8(
    strFromU8(files['ppt/slides/slide1.xml']).replace(
      '</p:spTree>',
      '<p:extLst><p:ext uri="urn:package-preservation-test"><x:payload xmlns:x="urn:unknown" flag="keep">opaque</x:payload></p:ext></p:extLst></p:spTree>'
    )
  );
  files['opaque.bin'] = new Uint8Array([0, 255, 17, 32]);
  const e = await open(zipSync(files));
  const pristine = unzipSync(e.save()!);
  const [target, ...untouched] = shapes(e);
  expect(untouched.length).toBeGreaterThan(0);
  const originalShapes = untouched.map(canonicalForm);
  const originalExtension = canonicalForm(find(root(e), 'p:extLst')!);
  const originalText = textOf(find(target, 'p:txBody'));
  expect(originalText.length).toBeGreaterThan(5);
  const select = (editor: PptxEditorInstance, from: number, to: number) =>
    editor.setSelection({
      kind: 'text',
      nodeId: shapes(editor)[0].id,
      anchor: { paragraph: 0, offset: from },
      focus: { paragraph: 0, offset: to },
    });
  select(e, 0, 5);
  expect(e.input('text', 'Edited').ok).toBe(true);
  const editedText = textOf(find(shapes(e)[0], 'p:txBody'));
  expect(editedText).toBe('Edited' + originalText.slice(5));
  expect(e.exec({ type: 'undo' }).ok).toBe(true);
  expect(textOf(find(shapes(e)[0], 'p:txBody'))).toBe(originalText);
  expect(e.snapshot().selection).toMatchObject({
    kind: 'text',
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 5 },
  });
  expect(e.exec({ type: 'redo' }).ok).toBe(true);
  select(e, 0, 6);
  expect(e.exec({ type: 'setFontFormat', bold: true, color: '#C00000', fontSizePt: 24 }).ok).toBe(
    true
  );
  const formatted = canonicalForm(shapes(e)[0]);
  expect(e.exec({ type: 'undo' }).ok).toBe(true);
  expect(e.exec({ type: 'redo' }).ok).toBe(true);
  expect(canonicalForm(shapes(e)[0])).toBe(formatted);
  e.setSelection({ kind: 'shapes', nodeIds: [target.id] });
  const xfrm = find(shapes(e)[0], 'a:xfrm')!;
  const oldX = Number(attr(find(xfrm, 'a:off'), 'x'));
  const oldWidth = Number(attr(find(xfrm, 'a:ext'), 'cx'));
  const transform = {
    type: 'setTransform' as const,
    nodeId: target.id,
    x: oldX + 914400,
    y: Number(attr(find(xfrm, 'a:off'), 'y')),
    cx: oldWidth + 914400,
    cy: Number(attr(find(xfrm, 'a:ext'), 'cy')),
    rot: 0,
  };
  expect(e.exec(transform).ok).toBe(true);
  expect(e.exec({ type: 'undo' }).ok).toBe(true);
  expect(attr(find(shapes(e)[0], 'a:off'), 'x')).toBe(String(oldX));
  expect(e.exec({ type: 'redo' }).ok).toBe(true);
  const saved = e.save()!;
  const savedFiles = unzipSync(saved);
  expect(Object.keys(savedFiles).sort()).toEqual(Object.keys(pristine).sort());
  for (const name of Object.keys(pristine)) {
    if (name !== part.slice(1)) expect(savedFiles[name]).toEqual(pristine[name]);
  }
  const again = await open(saved);
  expect(textOf(find(shapes(again)[0], 'p:txBody'))).toBe(editedText);
  select(again, 0, 6);
  expect(again.snapshot().format).toMatchObject({ bold: true, fontSize: 24, color: '#c00000' });
  const runs = elements(find(shapes(again)[0], 'a:p')!).filter((el) => el.kind === 'run');
  expect(textOf(runs[0])).toBe('Edited');
  expect(attr(find(runs[0], 'a:srgbClr'), 'val')).toBe('C00000');
  expect(runs.slice(1).every((r) => attr(find(r, 'a:srgbClr'), 'val') !== 'C00000')).toBe(true);
  expect(attr(find(shapes(again)[0], 'a:off'), 'x')).toBe(String(transform.x));
  expect(attr(find(shapes(again)[0], 'a:ext'), 'cx')).toBe(String(transform.cx));
  expect(shapes(again).slice(1).map(canonicalForm)).toEqual(originalShapes);
  expect(canonicalForm(find(root(again), 'p:extLst')!)).toBe(originalExtension);
  e.destroy();
  again.destroy();
});
