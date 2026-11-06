/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Overlapping syncs: a `sync()` called before the previous one settles waits for it, keeps
// the calls queued before it, and does not stop when an earlier sync is refused.

import { describe, expect, spyOn, test } from 'bun:test';
import { createStoreAutomationHost, type AutomationHost } from '@pptx-editor.dev/core/automation';
import { PptxDocumentStore, readPptx } from '@pptx-editor.dev/core/store';
import { isPptxEditorError } from '../index.ts';
import { resetDevWarnings } from '../runtime/dev-warnings.ts';
import { createRuntime } from '../runtime/runtime.ts';
import { deck } from './support.ts';

const code = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.code : String(error);

/** A server runtime whose host records the operations of each batch, in order. */
function counting() {
  const read = readPptx(deck());
  if (!read.ok) throw new Error('fixture');
  const inner = createStoreAutomationHost(new PptxDocumentStore(read.pkg));
  const batches: string[][] = [];
  const host: AutomationHost = {
    ...inner,
    execute: (request) => {
      batches.push(request.operations.map((operation) => operation.op));
      return inner.execute(request);
    },
  };
  return { runtime: createRuntime({ host, save: true }), batches };
}

describe('overlapping syncs', () => {
  test('two unawaited syncs apply in call order, each with its own calls', async () => {
    const { runtime, batches } = counting();
    const name = await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      shape.name = 'First';
      shape.load('name');
      const first = context.sync();
      shape.name = 'Second';
      const second = context.sync();
      await Promise.all([first, second]);
      shape.load('name');
      await context.sync();
      return shape.name;
    });
    // The second sync waited for the first and wrote last.
    expect(name).toBe('Second');
    const writes = batches.filter((ops) => ops.includes('setShape'));
    expect(writes).toHaveLength(2);
    expect(batches.indexOf(writes[0]!)).toBeLessThan(batches.indexOf(writes[1]!));
  });

  test('the second sync sees what the first loaded', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      shape.load('name');
      const order: string[] = [];
      const first = context.sync().then(() => order.push(`first:${shape.name}`));
      const second = context.sync().then(() => order.push(`second:${shape.name}`));
      await Promise.all([first, second]);
      expect(order).toEqual(['first:Box', 'second:Box']);
    });
  });

  test('an empty sync called during another resolves after it', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      shape.load('name');
      void context.sync();
      await context.sync();
      expect(shape.name).toBe('Box');
    });
  });

  test('a refused sync in the middle does not stop or reorder the others', async () => {
    const { runtime } = counting();
    const outcome = await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      const box = shapes.getItem('2');
      box.name = 'One';
      const first = context.sync();
      shapes.getItem('999').name = 'Missing';
      const second = context.sync();
      box.name = 'Three';
      box.load('name');
      const third = context.sync();
      const settled = await Promise.allSettled([first, second, third]);
      return { settled: settled.map((s) => s.status), name: box.name, error: settled[1] };
    });
    expect(outcome.settled).toEqual(['fulfilled', 'rejected', 'fulfilled']);
    expect(code((outcome.error as PromiseRejectedResult).reason)).toBe('ItemNotFound');
    expect(outcome.name).toBe('Three');
  });

  test('a refused sync applies none of its own calls', async () => {
    const { runtime } = counting();
    const name = await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      const box = shapes.getItem('2');
      const first = context.sync();
      box.name = 'Lost';
      shapes.getItem('999').name = 'Missing';
      const second = context.sync().catch((error: unknown) => code(error));
      box.load('name');
      const third = context.sync();
      await first;
      expect(await second).toBe('ItemNotFound');
      await third;
      return box.name;
    });
    expect(name).toBe('Box');
  });

  test('a sync that the run did not await applies before the run ends', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const box = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      box.load('name');
      await context.sync();
      box.name = 'Unawaited';
      void context.sync();
    });
    const name = await runtime.run(async (context) => {
      const box = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      box.load('name');
      await context.sync();
      return box.name;
    });
    expect(name).toBe('Unawaited');
  });

  test('a refused sync that the caller ignores is an unhandled rejection, as in Office.js', () => {
    // The test runner fails a test on an unhandled rejection, so a separate process shows it.
    const support = JSON.stringify(`${import.meta.dir}/support.ts`);
    const script = `
      import { open } from ${support};
      process.on('unhandledRejection', (reason) => console.log('unhandled', reason.code));
      const runtime = await open();
      await runtime.run(async (context) => {
        context.presentation.slides.getItemAt(0).shapes.getItem('999').name = 'Missing';
        void context.sync();
      });
      await new Promise((resolve) => setTimeout(resolve, 10));
      console.log('run ended');`;
    const child = Bun.spawnSync([process.execPath, '-e', script]);
    expect(child.stdout.toString().trim().split('\n')).toEqual([
      'unhandled ItemNotFound',
      'run ended',
    ]);
  });

  test('a run that returns with writes no sync sent warns once in development', async () => {
    const { runtime } = counting();
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      for (let i = 0; i < 2; i += 1)
        await runtime.run(async (context) => {
          context.presentation.slides.getItemAt(0).shapes.getItem('2').name = 'Unsent';
        });
      // A callback that throws leaves its calls without a warning: the error says enough.
      await runtime
        .run(async (context) => {
          context.presentation.slides.getItemAt(0).shapes.getItem('2').name = 'Unsent';
          throw new Error('stop');
        })
        .catch(() => undefined);
      // A load without a sync is not a lost edit: no warning.
      await runtime.run(async (context) => {
        context.presentation.slides.getItemAt(0).shapes.load('items/name');
      });
      expect(warn.mock.calls.map((call) => String(call[0]))).toEqual([
        expect.stringContaining('ended with 1 queued write that no sync() sent'),
      ]);
    } finally {
      warn.mockRestore();
      resetDevWarnings();
    }
  });

  test('the warning names the argument refusals that the missing sync would have given', async () => {
    const { runtime } = counting();
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      await runtime.run(async (context) => {
        const shapes = context.presentation.slides.getItemAt(0).shapes;
        shapes.getItem('2').name = 'Unsent';
        shapes.addGeometricShape('Blob' as never);
      });
      const [message] = warn.mock.calls.map((call) => String(call[0]));
      expect(message).toContain('ended with 2 queued writes that no sync() sent');
      expect(message).toContain('That sync would have refused with this error:');
      expect(message).toContain('InvalidArgument at ');
      expect(message).toContain('addGeometricShape');
    } finally {
      warn.mockRestore();
      resetDevWarnings();
    }
  });
});
