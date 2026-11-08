// AutoCorrect edits in a collaboration room: a list a typed marker starts, and a fraction sign
// a space makes, reach the other peer and save the same on both.

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

  test('a space after a typed fraction makes the fraction sign on both peers', async () => {
    await convergeInput('text', { arrange: newParagraph('1/2'), data: ' ' });
  });

  test('a system replacement of a typed word reaches the other peer', async () => {
    // A replacement selects its target in the model and types the text over it.
    const arrange = (e: PptxEditorInstance): void => {
      newParagraph('teh')(e);
      const sel = e.snapshot().selection;
      if (sel.kind !== 'text') throw new Error('no caret');
      const anchor = { paragraph: 1, offset: 0 };
      e.setSelection({ ...sel, anchor, focus: { paragraph: 1, offset: 3 } });
    };
    await convergeInput('text', { arrange, data: 'the' });
  });
});
