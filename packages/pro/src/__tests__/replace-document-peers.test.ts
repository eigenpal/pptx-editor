// `replaceDocument` and document refresh `mode: 'edit'` in a collaboration room, after DOCX
// Editor's replace-document-peers.test.ts and document-refresh-edit-peers.test.ts: the update
// reaches the other participant as an ordinary edit and converges with concurrent edits, undo
// reverts only the update, and a reload refuses with `collaboration`.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createDocumentRefresh } from '@pptx-editor.dev/core';
import { serializeXmlPart, xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { differingParts, fingerprint, link, replica } from './convergence-harness.ts';

const box = (id: number, text: string) =>
  shape({ id, x: 400_000, y: 400_000, cx: 4_000_000, cy: 800_000, text: para(run(text)) });
const deck = (...texts: string[]) =>
  buildPptx({ slides: texts.map((t) => ({ shapes: box(2, t) })) });

async function room(bytes: Uint8Array) {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const wire = link(docA, docB);
  const a = await replica(docA, 'a', bytes);
  const b = await replica(docB, 'b', bytes);
  const converged = () => {
    const diff = differingParts(a.session!.store.package, b.session!.store.package);
    expect(diff).toEqual([]);
    expect(fingerprint(a)).toBe(fingerprint(b));
  };
  const done = () => {
    a.destroy();
    b.destroy();
  };
  return { a, b, wire, converged, done };
}

describe('replaceDocument in a collaboration room', () => {
  test('reaches the other participant and converges', async () => {
    const { a, b, converged, done } = await room(deck('One', 'Two'));
    const r = a.exec({ type: 'replaceDocument', pptx: deck('One changed', 'Two', 'Three') });
    expect(r.ok).toBe(true);
    expect(b.snapshot().slides.length).toBe(3);
    converged();
    done();
  });

  test('converges with a concurrent edit of another slide, and undo reverts only the update', async () => {
    const { a, b, wire, converged, done } = await room(deck('One', 'Two', 'Three'));
    wire.pause();
    expect(a.exec({ type: 'replaceDocument', pptx: deck('One changed', 'Two', 'Three') }).ok).toBe(
      true
    );
    b.goToSlide(2);
    expect(b.exec({ type: 'duplicateSlide', index: 2 }).ok).toBe(true);
    wire.flush();
    converged();
    expect(a.snapshot().slides.length).toBe(4);
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    converged();
    // The peer's duplicate stays; the update is gone.
    expect(b.snapshot().slides.length).toBe(4);
    const first = (e: typeof a) =>
      serializeXmlPart(xmlPartOf(e.session!.store.package, '/ppt/slides/slide1.xml')!);
    expect(first(a)).not.toContain('One changed');
    expect(first(b)).not.toContain('One changed');
    done();
  });
});

describe('document refresh in a collaboration room', () => {
  test("mode 'edit' applies for every participant; a reload refuses with collaboration", async () => {
    const { a, b, converged, done } = await room(deck('One'));
    const refresh = createDocumentRefresh(a);
    const submission = await refresh.capture();
    const reload = await refresh.applyUpdate({ submission, sequence: 1, bytes: deck('Two') });
    expect(reload).toMatchObject({ ok: false, code: 'collaboration' });
    const edit = await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck('Two', 'Three'),
      mode: 'edit',
    });
    expect(edit).toMatchObject({ ok: true, mode: 'edit' });
    expect(b.snapshot().slides.length).toBe(2);
    converged();
    done();
  });

  test("another participant's edit after capture refuses with local-edits", async () => {
    const { a, b, done } = await room(deck('One'));
    const refresh = createDocumentRefresh(a);
    const submission = await refresh.capture();
    expect(b.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    const r = await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck('X'),
      mode: 'edit',
    });
    expect(r).toMatchObject({ ok: false, code: 'local-edits' });
    expect(a.snapshot().slides.length).toBe(2);
    done();
  });
});
