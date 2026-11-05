/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Every argument refusal says what to pass: client checks through `failArgument`, and engine
// checks through `argument-hints.ts`. None falls back to the generic text of the code.

import { describe, expect, test } from 'bun:test';
import { isPptxEditorError, type RequestContext } from '../index.ts';
import { CODE_DIAGNOSTICS } from '../runtime/edit-diagnostics.ts';
import { open } from './support.ts';

const GENERIC = CODE_DIAGNOSTICS['invalid-argument']![1];

/** The hint of the refusal that `call` leads to, at the call or at the sync. */
async function hintOf(context: RequestContext, call: () => unknown): Promise<string | undefined> {
  try {
    call();
    await context.sync();
  } catch (error) {
    if (!isPptxEditorError(error) || error.code !== 'InvalidArgument') throw error;
    return error.debugInfo.innerError?.message;
  }
  throw new Error('expected a refusal');
}

describe('argument refusals', () => {
  test('each names what to pass', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      const box = () => slide.shapes.addTextBox('x');
      const cases: [() => unknown, string][] = [
        [() => slide.shapes.getItemAt(-1), 'a whole number of 0 or more'],
        [() => context.presentation.slides.getItemAt(1.5), 'a whole number of 0 or more'],
        [() => slide.shapes.addTable(2, 2, { values: [['a']] }), 'an array of 2 entries'],
        [() => slide.shapes.load({ top: -1 }), 'a whole number of 0 or more'],
        [() => box().fill.setSolidColor('notacolor'), "'#RRGGBB'"],
        [() => (box().textFrame.textRange.font.size = -4), 'from 1 to 4000 points'],
        [() => (box().textFrame.textRange.font.size = Number.NaN), 'a finite number'],
        [() => (box().lineFormat.weight = -1), 'from 0 to 1584 points'],
        [() => slide.shapes.addTable(0, 2), 'of 1 or more'],
        [() => slide.shapes.addTextBox('x', { left: 1e9 }), 'points from -169 093 to 169 093'],
        [() => (box().textFrame.textRange.paragraphFormat.indentLevel = 12), '1 to 9'],
      ];
      for (const [call, expected] of cases) {
        const hint = await hintOf(context, call);
        expect(hint).toBeDefined();
        expect(hint).not.toBe(GENERIC);
        expect(hint).toContain(expected);
      }
    });
  });
});
