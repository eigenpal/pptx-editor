// The automation `readTransition` and `setTransition` operations over `p:transition`
// (`CT_SlideTransition`, ECMA-376 Part 1, §19.3.1.50).

import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import {
  createServerAutomationHost,
  type AutomationBatchResponse,
  type AutomationHandle,
  type AutomationHost,
  type AutomationOperation,
  type AutomationValue,
} from '../index.ts';

function open(): AutomationHost {
  const opened = createServerAutomationHost(
    buildPptx({ slides: [{ layout: 'blank', shapes: '' }] })
  );
  if (!opened.ok) throw new Error(JSON.stringify(opened.error));
  return opened.host;
}

function exec(host: AutomationHost, operations: AutomationOperation[]): AutomationBatchResponse {
  return host.execute({ operations });
}

/** Run a batch that must succeed and return its values. */
function ok(host: AutomationHost, operations: AutomationOperation[]): AutomationValue[] {
  const response = exec(host, operations);
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results.map((r) => (r.status === 'ok' ? r.value : { kind: 'null' as const }));
}

function firstSlide(host: AutomationHost): AutomationHandle {
  const [value] = ok(host, [{ op: 'listSlides' }]);
  if (value?.kind !== 'handles') throw new Error('no slides');
  return value.handles[0]!;
}

describe('automation host: slide transitions', () => {
  test('reads, sets, and clears a transition', () => {
    const host = open();
    const slide = firstSlide(host);
    const read = () =>
      (ok(host, [{ op: 'readTransition', slide }])[0] as { transition: unknown }).transition;
    expect(read()).toEqual({
      kind: 'none',
      durationMs: 750,
      advanceOnClick: true,
      advanceAfterMs: null,
    });
    ok(host, [{ op: 'setTransition', slide, transition: { kind: 'fade', durationMs: 1200 } }]);
    ok(host, [{ op: 'setTransition', slide, transition: { advanceAfterMs: 3000 } }]);
    expect(read()).toEqual({
      kind: 'fade',
      durationMs: 1200,
      advanceOnClick: true,
      advanceAfterMs: 3000,
    });
    // The same settings again write nothing, and answer applied.
    const revision = exec(host, [{ op: 'readTransition', slide }]).revision;
    ok(host, [{ op: 'setTransition', slide, transition: { kind: 'fade' } }]);
    expect(exec(host, [{ op: 'readTransition', slide }]).revision).toBe(revision);
    ok(host, [
      {
        op: 'setTransition',
        slide,
        transition: { kind: 'none', advanceOnClick: true, advanceAfterMs: null },
      },
    ]);
    expect(read()).toMatchObject({ kind: 'none', advanceAfterMs: null });
  });

  test('refuses bad arguments, and the batch writes nothing', () => {
    const host = open();
    const slide = firstSlide(host);
    const refusals: [unknown, string][] = [
      [{ kind: 'dissolve' }, 'transitionKind'],
      [{ kind: 'other' }, 'transitionKind'],
      [{ durationMs: 9 }, 'durationMs'],
      [{ durationMs: 59_001 }, 'durationMs'],
      [{ durationMs: Number.POSITIVE_INFINITY }, 'durationMs'],
      [{ advanceOnClick: 1 }, 'advanceOnClick'],
      [{ advanceAfterMs: -1 }, 'advanceAfterMs'],
      [{ advanceAfterMs: 3_600_001 }, 'advanceAfterMs'],
      [{ spd: 'fast' }, 'transition'],
      [null, 'transition'],
    ];
    for (const [transition, detail] of refusals) {
      const refused = exec(host, [
        { op: 'setTransition', slide, transition: { kind: 'cut' } },
        { op: 'setTransition', slide, transition } as unknown as AutomationOperation,
      ]);
      expect(refused.results[1]).toMatchObject({ error: { code: 'invalid-argument', detail } });
    }
    expect(ok(host, [{ op: 'readTransition', slide }])[0]).toMatchObject({
      transition: { kind: 'none' },
    });
    // No effect: a duration has nothing to keep it.
    const noEffect = exec(host, [{ op: 'setTransition', slide, transition: { durationMs: 500 } }]);
    expect(noEffect.results[0]).toMatchObject({
      error: { code: 'unsupported-content', detail: 'transition-no-effect' },
    });
  });
});
