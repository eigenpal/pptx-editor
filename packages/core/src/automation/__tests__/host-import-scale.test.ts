// Slide import at scale: the work grows with the number of slides, not with its square.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { presentationInfo, readPptx } from '../../store/index.ts';
import { toBase64 } from '../commands/base64.ts';
import { OpBatch } from '../commands/op-batch.ts';
import { createServerAutomationHost, type AutomationHost } from '../index.ts';

/** `count` slides, each with a text box and notes. */
function deck(count: number): Uint8Array {
  return buildPptx({
    slides: Array.from({ length: count }, (_, i) => ({
      layout: 'blank',
      shapes: shape({
        id: 2,
        name: 'B',
        x: 0,
        y: 0,
        cx: 100_000,
        cy: 100_000,
        text: para(run(`S${i}`)),
      }),
      notes: `n${i}`,
    })),
  } as Parameters<typeof buildPptx>[0]);
}

function open(bytes: Uint8Array): AutomationHost {
  const opened = createServerAutomationHost(bytes);
  if (!opened.ok) throw new Error(JSON.stringify(opened.error));
  return opened.host;
}

/**
 * Import a deck of `count` slides into itself, counting the package updates the import makes:
 * each `OpBatch.push` or `pushAll` applies its ops to the whole working package once.
 */
function importWork(count: number): { updates: number; host: AutomationHost } {
  const bytes = deck(count);
  const host = open(bytes);
  const push = OpBatch.prototype.push;
  const pushAll = OpBatch.prototype.pushAll;
  let updates = 0;
  OpBatch.prototype.push = function (this: OpBatch, ...args: Parameters<OpBatch['push']>) {
    updates += 1;
    return push.apply(this, args);
  };
  OpBatch.prototype.pushAll = function (this: OpBatch, ...args: Parameters<OpBatch['pushAll']>) {
    updates += 1;
    return pushAll.apply(this, args);
  };
  try {
    const response = host.execute({
      operations: [{ op: 'insertSlides', base64: toBase64(bytes) }],
    });
    if (!response.ok) throw new Error(JSON.stringify(response.results));
  } finally {
    OpBatch.prototype.push = push;
    OpBatch.prototype.pushAll = pushAll;
  }
  return { updates, host };
}

describe('automation host: slide import at scale', () => {
  test('an import makes the same number of package updates for 200 slides as for 800', () => {
    // A package update costs time in proportion to the package, so an import that made one
    // for each slide grew with the square of the slide count (14.5 s for 800 slides).
    const small = importWork(200);
    const large = importWork(800);
    expect(large.updates).toBe(small.updates);
    expect(large.updates).toBeLessThan(50);
    const saved = large.host.save();
    if (!saved.ok) throw new Error(saved.error.code);
    const read = readPptx(saved.bytes);
    if (!read.ok) throw new Error('the saved package does not read');
    const slides = presentationInfo(read.pkg).slides;
    expect(slides).toHaveLength(1_600);
    expect(new Set(slides.map((s) => s.sldId)).size).toBe(1_600);
    expect(new Set(slides.map((s) => s.partName)).size).toBe(1_600);
  }, 60_000);
});
