import { expect, test } from 'bun:test';
import { composeFontOrigins } from '../../layout/font-resolver.ts';
import { MAX_DEFERRED_STEPS } from '../../layout/font-deferred.ts';
import type { FontConfigurationFragment } from '../../layout/font-composition.ts';

async function exhaust(answer: FontConfigurationFragment | undefined): Promise<number> {
  let steps = 0;
  while (answer?.deferred) {
    expect(steps).toBeLessThanOrEqual(MAX_DEFERRED_STEPS);
    answer = await answer.deferred;
    steps += 1;
  }
  return steps;
}

test('cyclic deferred font answers stop and let event-loop timers run', async () => {
  const errors: unknown[] = [];
  const answer: { sources: []; deferred?: Promise<FontConfigurationFragment> } = { sources: [] };
  answer.deferred = Promise.resolve(answer);
  const composed = await composeFontOrigins(
    [answer],
    { families: [], defaultFamily: 'Arial' },
    {
      onOriginFailure: ({ cause }) => errors.push(cause),
    }
  );
  expect(await exhaust(composed)).toBe(MAX_DEFERRED_STEPS);
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(errors).toHaveLength(1);
  expect(String(errors[0])).toContain('256 deferred steps');
});

test('fresh deferred answers cannot bypass the background composition bound', async () => {
  const errors: unknown[] = [];
  const next = (): FontConfigurationFragment => ({
    sources: [],
    deferred: {
      then: (resolve: (answer: FontConfigurationFragment) => unknown) => resolve(next()),
    } as unknown as Promise<FontConfigurationFragment>,
  });
  const composed = await composeFontOrigins(
    [next()],
    { families: [], defaultFamily: 'Arial' },
    {
      onOriginFailure: ({ cause }) => errors.push(cause),
    }
  );
  expect(await exhaust(composed)).toBe(MAX_DEFERRED_STEPS);
  expect(errors).toHaveLength(1);
});
