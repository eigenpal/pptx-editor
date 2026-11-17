// The `replaceDocument` command, after DOCX Editor's replace-document-command.test.ts: the file
// replaces the presentation as one undo step, a handle guards it, and every refusal leaves the
// presentation and its history unchanged.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor, onEditorEvent, type DocumentChange } from '../index.ts';
import {
  canonicalForm,
  presentationInfo,
  readPptx,
  serializeXmlPart,
  xmlPartOf,
  type PptxPackage,
} from '../../store/index.ts';

const box = (id: number, text: string) =>
  shape({ id, x: 100_000, y: 100_000, cx: 3_000_000, cy: 600_000, text: para(run(text)) });
const deck = (...texts: string[]) =>
  buildPptx({ slides: texts.map((t) => ({ shapes: box(2, t) })) });

async function open(bytes = deck('One', 'Two')) {
  const editor = createPptxEditor({ fixedMeasurer: true });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

const textOf = (editor: Awaited<ReturnType<typeof open>>) => {
  const pkg = editor.session!.store.package;
  return presentationInfo(pkg).slides.map((s) => {
    const part = xmlPartOf(pkg, s.partName);
    return part ? serializeXmlPart(part) : '';
  });
};

/** Every part of `pkg` as canonical text, by name. */
function canonical(pkg: PptxPackage): Map<string, string> {
  const out = new Map<string, string>();
  for (const [name, part] of pkg.parts)
    out.set(name, part.kind === 'xml' ? canonicalForm(part.xml.root) : `${part.bytes.length}`);
  return out;
}

describe('replaceDocument', () => {
  test('replaces the presentation as one undo step above earlier history', async () => {
    const editor = await open();
    expect(editor.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    const beforeReplace = canonical(editor.session!.store.package);
    const result = editor.exec({ type: 'replaceDocument', pptx: deck('Alpha', 'Beta', 'Gamma') });
    expect(result).toEqual({ ok: true, replaceDocument: { ok: true, changed: true } });
    expect(editor.snapshot().slides.length).toBe(3);
    expect(textOf(editor).join('|')).toContain('Gamma');
    // One undo returns to the presentation before the replacement; the earlier edit stays.
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonical(editor.session!.store.package)).toEqual(beforeReplace);
    expect(editor.snapshot().slides.length).toBe(3);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.snapshot().slides.length).toBe(2);
    // Redo applies it again.
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(textOf(editor).join('|')).toContain('Gamma');
    editor.destroy();
  });

  test('the content equals a load of the file, through save and reopen', async () => {
    const editor = await open();
    const file = deck('Alpha', 'Beta', 'Gamma');
    editor.exec({ type: 'replaceDocument', pptx: file });
    const reference = await open(file);
    const saved = readPptx(editor.save()!);
    const expected = readPptx(reference.save()!);
    if (!saved.ok || !expected.ok) throw new Error('save did not reopen');
    expect(canonical(saved.pkg)).toEqual(canonical(expected.pkg));
    editor.destroy();
    reference.destroy();
  });

  test('accepts an ArrayBuffer, and an identical file changes nothing', async () => {
    const editor = await open();
    const same = deck('One', 'Two');
    const revision = editor.getDocumentHandle().revision;
    const result = editor.exec({ type: 'replaceDocument', pptx: same.slice().buffer });
    expect(result).toEqual({ ok: true, replaceDocument: { ok: true, changed: false } });
    expect(editor.getDocumentHandle().revision).toBe(revision);
    expect(editor.snapshot().canUndo).toBe(false);
    editor.destroy();
  });

  test('a matching handle applies; an edited document refuses and changes nothing', async () => {
    const editor = await open();
    const expected = editor.getDocumentHandle();
    expect(editor.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    const stale = editor.exec({ type: 'replaceDocument', pptx: deck('X'), expected });
    expect(stale.ok).toBe(false);
    expect(stale.replaceDocument).toMatchObject({ code: 'staleDocument', reason: 'edited' });
    expect(editor.snapshot().slides.length).toBe(3);
    const fresh = editor.exec({
      type: 'replaceDocument',
      pptx: deck('X'),
      expected: editor.getDocumentHandle(),
    });
    expect(fresh.ok).toBe(true);
    expect(editor.snapshot().slides.length).toBe(1);
    editor.destroy();
  });

  test('a handle from document A never replaces document B, even at the same revision', async () => {
    const editor = await open();
    const fromA = editor.getDocumentHandle();
    await editor.load(deck('B1', 'B2'));
    expect(editor.getDocumentHandle().revision).toBe(fromA.revision);
    const r = editor.exec({ type: 'replaceDocument', pptx: deck('X'), expected: fromA });
    expect(r.replaceDocument).toMatchObject({ code: 'staleDocument', reason: 'otherDocument' });
    editor.destroy();
  });

  test('refuses a history group, bad bytes, and a malformed handle with typed outcomes', async () => {
    const editor = await open();
    const group = editor.beginHistoryGroup();
    const grouped = editor.exec(
      { type: 'replaceDocument', pptx: deck('X') },
      { historyGroup: group }
    );
    group.end();
    expect(grouped.replaceDocument).toMatchObject({ code: 'invalidRequest' });
    const junk = editor.exec({ type: 'replaceDocument', pptx: new Uint8Array([1, 2, 3]) });
    expect(junk.replaceDocument).toMatchObject({ code: 'invalidDocument' });
    const wrong = editor.exec({
      type: 'replaceDocument',
      pptx: 'nope' as unknown as Uint8Array,
    });
    expect(wrong.replaceDocument).toMatchObject({ code: 'invalidRequest' });
    const handle = editor.exec({
      type: 'replaceDocument',
      pptx: deck('X'),
      expected: { loadId: 1 } as never,
    });
    expect(handle.replaceDocument).toMatchObject({ code: 'invalidRequest' });
    expect(editor.snapshot().slides.length).toBe(2);
    expect(editor.snapshot().canUndo).toBe(false);
    editor.destroy();
  });

  test('viewing mode refuses with a typed reason', async () => {
    const editor = await open();
    editor.setMode('view');
    const r = editor.exec({ type: 'replaceDocument', pptx: deck('X') });
    expect(r.replaceDocument).toMatchObject({ code: 'editNotPermitted', reason: 'viewing' });
    expect(editor.snapshot().slides.length).toBe(2);
    editor.destroy();
  });

  test('an editor with no document refuses with unavailable', () => {
    const editor = createPptxEditor({ fixedMeasurer: true });
    const r = editor.exec({ type: 'replaceDocument', pptx: deck('X') });
    expect(r.replaceDocument).toMatchObject({ code: 'unavailable' });
    editor.destroy();
  });

  test('the change event reports source refresh; edits and undo report none', async () => {
    const editor = await open();
    const seen: DocumentChange[] = [];
    onEditorEvent(editor, 'change', (_snapshot, change) => void seen.push(change));
    editor.exec({ type: 'replaceDocument', pptx: deck('Alpha') });
    editor.exec({ type: 'undo' });
    editor.exec({ type: 'duplicateSlide', index: 0 });
    await editor.load(deck('Loaded'));
    expect(seen.map((c) => c.source)).toEqual(['refresh', undefined, undefined, 'load']);
    expect(seen.every((c) => typeof c.revision === 'number')).toBe(true);
    editor.destroy();
  });

  test('the slide on screen and a surviving selection stay', async () => {
    const editor = await open(deck('One', 'Two', 'Three'));
    editor.goToSlide(1);
    const two = editor.snapshot().shapes?.[0]?.nodeId;
    if (!two) throw new Error('no shape on slide 2');
    editor.setSelection({ kind: 'shapes', nodeIds: [two] });
    editor.exec({ type: 'replaceDocument', pptx: deck('One', 'Two changed', 'Three') });
    expect(editor.snapshot().currentSlide).toBe(1);
    expect(editor.snapshot().selection).toEqual({ kind: 'shapes', nodeIds: [two] });
    editor.destroy();
  });
});
