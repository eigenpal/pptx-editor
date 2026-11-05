import { afterEach, describe, expect, mock, test } from 'bun:test';
import * as session from '../../packages/core/src/editor/document-session.ts';
import * as store from '../../packages/core/src/store/index.ts';
import { buildPptx, para, run, shape } from '../fixtures/builder.ts';

// An engine fault at each step after the first open. Each stand-in calls the real function
// unless its step is set to fail, so the mock is harmless to the other test files.
const fail = new Set<string>();
const realStore = { ...store };
const realSession = { ...session };
// The census step in progress: a store function can serve more than one step.
let current = '';
// The error class of the fault: an engine error, or a fault of the machine.
let Fault: new (message: string) => Error = TypeError;
const fault = (step: string): void => {
  if (fail.has(step) && current === step) throw new Fault(`${step} fault`);
};
let opens = 0;
mock.module('../../packages/core/src/store/index.ts', () => ({
  ...realStore,
  writePptx: (...args: Parameters<typeof store.writePptx>) => {
    fault('write');
    return realStore.writePptx(...args);
  },
  semanticDigest: (...args: Parameters<typeof store.semanticDigest>) => {
    fault('digest');
    return realStore.semanticDigest(...args);
  },
  packageFingerprints: (...args: Parameters<typeof store.packageFingerprints>) => {
    fault('compare');
    return realStore.packageFingerprints(...args);
  },
}));
mock.module('../../packages/core/src/editor/document-session.ts', () => ({
  ...realSession,
  openPresentation: (...args: Parameters<typeof session.openPresentation>) => {
    opens += 1;
    if (opens === 2) fault('reopen');
    return realSession.openPresentation(...args);
  },
}));
const { censusDeck, MachineFault } = await import('./roundtrip-census.ts');

const deck = buildPptx({
  slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 9, cy: 9, text: para(run('x')) }) }],
});

const census = (keep?: (kind: string) => void) =>
  censusDeck(deck, {
    keep,
    onStage: (stage) => {
      current = stage;
    },
  });

afterEach(() => {
  fail.clear();
  opens = 0;
  current = '';
  Fault = TypeError;
});

describe('round-trip census faults', () => {
  test('a write fault is an error of the write step, with the counts of the open', async () => {
    fail.add('write');
    const r = await census();
    expect(r.error).toEqual({ stage: 'write', reason: 'exception', detail: 'TypeError' });
    expect(r.slides).toBe(1);
    expect(r.roundTrip).toBeNull();
    expect(r.edit).toBeNull();
  });

  test('a reopen fault is an error of the reopen step, and the saved copy is still kept', async () => {
    fail.add('reopen');
    const kept: string[] = [];
    const r = await census((kind) => kept.push(kind));
    expect(r.error).toEqual({ stage: 'reopen', reason: 'exception', detail: 'TypeError' });
    expect(kept).toEqual(['saved']);
    expect(r.roundTrip).toBeNull();
  });

  test('a digest fault leaves the comparison and pass 2 to run', async () => {
    fail.add('digest');
    const r = await census();
    expect(r.error).toEqual({ stage: 'digest', reason: 'exception', detail: 'TypeError' });
    expect(r.roundTrip?.digestStable).toBeNull();
    expect(r.roundTrip?.fingerprintEqual).toBe(true);
    expect(r.edit).not.toBeNull();
  });

  test('a compare fault records no round-trip result and no pass 2', async () => {
    fail.add('compare');
    const r = await census();
    expect(r.error).toEqual({ stage: 'compare', reason: 'exception', detail: 'TypeError' });
    expect(r.roundTrip).toBeNull();
    expect(r.edit).toBeNull();
  });

  test('a machine fault in any step is never a deck error: it passes up to the caller', async () => {
    Fault = MachineFault;
    for (const step of ['write', 'reopen', 'digest', 'compare']) {
      fail.clear();
      fail.add(step);
      opens = 0;
      await expect(census()).rejects.toBeInstanceOf(MachineFault);
    }
  });
});
