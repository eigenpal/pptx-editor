/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Synthetic decks and runtimes for the object model tests.

import { buildPptx, para, run, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { readPptx } from '@pptx-editor.dev/core/store';
import { PptxEditor, type PptxEditorServerRuntime } from '../index.ts';

export const IN = 914_400;

/** Two slides: a filled box with two paragraphs and a plain box, then an empty slide. */
export function deck(): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({
            id: 2,
            name: 'Box',
            x: IN,
            y: IN,
            cx: 3 * IN,
            cy: IN,
            fill: solid('4472C4'),
            text: para(run('Hello', 'b="1" sz="2400"')) + para(run('World')),
          }) + shape({ id: 3, name: 'Other', x: 5 * IN, y: IN, cx: IN, cy: IN }),
      },
      { layout: 'title', shapes: '' },
    ],
  });
}

/** A server runtime over `bytes`, measuring with no substitute fonts so tests stay fast. */
export function open(bytes: Uint8Array = deck()): Promise<PptxEditorServerRuntime> {
  return PptxEditor.createServer(bytes, { substitutes: false });
}

/** Save and reopen, checking the saved bytes read back. */
export async function reopen(runtime: PptxEditorServerRuntime): Promise<PptxEditorServerRuntime> {
  const bytes = await runtime.save();
  if (!readPptx(bytes).ok) throw new Error('the saved presentation does not read back');
  return open(bytes);
}

/** The error a promise rejects with. */
export async function rejection(promise: Promise<unknown>): Promise<unknown> {
  try {
    await promise;
  } catch (error) {
    return error;
  }
  throw new Error('expected a rejection');
}

/** The error a call throws. */
export function thrown(call: () => unknown): unknown {
  try {
    call();
  } catch (error) {
    return error;
  }
  throw new Error('expected a throw');
}

/**
 * The refusal of a call whose argument this API does not accept: the call does not throw, and
 * the next sync refuses with it, as in Office.js.
 */
export async function refusedAtSync(
  context: { sync(): Promise<void> },
  call: () => unknown
): Promise<unknown> {
  call();
  return rejection(context.sync());
}
