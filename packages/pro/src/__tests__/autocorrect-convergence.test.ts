// AutoCorrect edits in a collaboration room: a list a typed marker starts, and a capital a
// space makes, reach the other peer and save the same on both.

import { describe, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core';
import { convergeInput } from './convergence-harness.ts';

/** A caret at the end of shape `A` ("Hello world"), in a new paragraph holding `typed`. */
const newParagraph =
  (typed: string) =>
  (e: PptxEditorInstance): void => {
    const nodeId = e.snapshot().shapes?.find((s) => s.name === 'A')?.nodeId;
    if (!nodeId) throw new Error('no shape A');
    const end = { paragraph: 0, offset: 11 };
    e.setSelection({ kind: 'text', nodeId, anchor: end, focus: end });
    e.input('paragraph');
    // More than one character at once is never corrected, so the marker stays as typed.
    e.input('text', typed);
  };

describe('AutoCorrect in a collaboration room', () => {
  test('a typed list marker and a space start a list on both peers', async () => {
    await convergeInput('text', { arrange: newParagraph('1.'), data: ' ' });
  });

  test('a space after a first word capitalizes it on both peers', async () => {
    await convergeInput('text', { arrange: newParagraph('next'), data: ' ' });
  });
});
