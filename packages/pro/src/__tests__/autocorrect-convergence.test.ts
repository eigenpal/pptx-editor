// AutoCorrect edits in a collaboration room: a list a typed marker starts, a fraction sign a
// space makes, and a system replacement reach the other peer and save the same on both.

import { afterAll, describe, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import type { CommandResult, PptxEditorInstance } from '@pptx-editor.dev/core';
import { convergeEdit, convergeInput } from './convergence-harness.ts';

const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

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
    // Peer A's surface gets `insertReplacementText` over `teh` in its input, as a system
    // correction sends it: the replacement path (`replaceText`) writes the edit.
    const g = globalThis as { ResizeObserver?: unknown };
    const had = 'ResizeObserver' in g;
    g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
    try {
      await convergeEdit(
        'insertReplacementText',
        { arrange: newParagraph('teh ') },
        (a): CommandResult => {
          const host = doc.createElement('div') as unknown as HTMLElement;
          doc.body.append(host as unknown as Node);
          a.attach(host);
          const input = host.querySelector('textarea') as HTMLTextAreaElement;
          if (input.value !== 'teh ') throw new Error(`the input holds "${input.value}"`);
          input.setSelectionRange(0, 3);
          const event = new win.InputEvent('beforeinput', {
            inputType: 'insertReplacementText',
            data: 'the',
            bubbles: true,
            cancelable: true,
          });
          input.dispatchEvent(event as unknown as Event);
          a.detach();
          return a.snapshot().lastRejection ? { ok: false, reason: 'refused' } : { ok: true };
        }
      );
    } finally {
      if (!had) delete g.ResizeObserver;
    }
  });
});
