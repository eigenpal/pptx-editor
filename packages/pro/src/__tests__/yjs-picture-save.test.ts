/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import { canonicalPackageFingerprint, readPptx } from '@pptx-editor.dev/core/store';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { peers } from './lineage-support.ts';
import { coverageDeck, png } from './convergence-harness.ts';

function savedImages(editor: PptxEditorInstance): string[] {
  const result = readPptx(editor.save()!);
  if (!result.ok) throw new Error('save refused');
  return [...result.pkg.parts.keys()].filter((name) => name.startsWith('/ppt/media/')).sort();
}
function selectPicture(editor: PptxEditorInstance) {
  const picture = editor.snapshot().shapes!.find((shape) => shape.name === 'Picture')!;
  editor.setSelection({ kind: 'shapes', nodeIds: [picture.nodeId] });
}

test('saving a replacement does not remove media needed by a disconnected peer or undo', async () => {
  const { editors, sync, docs } = await peers(coverageDeck(), [1001, 2002]);
  const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
  try {
    selectPicture(a);
    selectPicture(b);
    sync.pause();
    expect(a.exec({ type: 'replaceImage', bytes: png(6, 6, 40), mime: 'image/png' }).ok).toBe(true);
    expect(b.exec({ type: 'duplicate' }).ok).toBe(true);
    const state = Y.encodeStateAsUpdate(docs[0]!);
    const before = canonicalPackageFingerprint(a.session!.store.package);
    expect(savedImages(a)).toHaveLength(1);
    expect(Y.encodeStateAsUpdate(docs[0]!)).toEqual(state);
    expect(canonicalPackageFingerprint(a.session!.store.package)).toBe(before);
    expect(a.session!.store.package.parts.has('/ppt/media/image1.png')).toBe(true);
    sync.flush();
    expect(savedImages(a)).toHaveLength(2);
    expect(savedImages(b)).toEqual(savedImages(a));
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expect(savedImages(a)).toEqual(['/ppt/media/image1.png']);
    expect(savedImages(b)).toEqual(savedImages(a));
    expect(a.exec({ type: 'redo' }).ok).toBe(true);
    expect(savedImages(a)).toHaveLength(2);
    expect(savedImages(b)).toEqual(savedImages(a));
  } finally {
    for (const editor of editors) editor.destroy();
    for (const doc of docs) doc.destroy();
  }
});

test('deleting a picture omits its bytes from save but undo restores them on every peer', async () => {
  const { editors, docs } = await peers(coverageDeck(), [1001, 2002]);
  const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
  try {
    selectPicture(a);
    expect(a.exec({ type: 'delete' }).ok).toBe(true);
    expect(savedImages(a)).toEqual([]);
    expect(savedImages(b)).toEqual([]);
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expect(savedImages(a)).toEqual(['/ppt/media/image1.png']);
    expect(savedImages(b)).toEqual(savedImages(a));
  } finally {
    for (const editor of editors) editor.destroy();
    for (const doc of docs) doc.destroy();
  }
});
