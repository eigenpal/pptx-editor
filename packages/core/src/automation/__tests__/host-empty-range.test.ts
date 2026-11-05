// The font of an empty text range: the format that text typed at that point takes.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  createServerAutomationHost,
  type AutomationHandle,
  type AutomationHost,
  type AutomationValue,
} from '../index.ts';

const IN = 914_400;

/** The default text: see the comment in it. */
const TEXT =
  // "Bold" (0-4), "Plain" (4-9), a line break (9-10), "Big" (10-13); then "" and "Tail".
  para(run('Bold', 'b="1"') + run('Plain') + '<a:br/>' + run('Big', 'sz="4000"')) +
  para('') +
  para(run('Tail', 'i="1"'));

function setup(text = TEXT): { host: AutomationHost; box: AutomationHandle } {
  const opened = createServerAutomationHost(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            name: 'Box',
            x: IN,
            y: IN,
            cx: 3 * IN,
            cy: IN,
            text,
          }),
        },
      ],
    })
  );
  if (!opened.ok) throw new Error('fixture');
  const host = opened.host;
  const response = host.execute({
    operations: [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'getShape', container: { result: 0 }, key: { index: 0 } },
    ],
  });
  const value = response.results[1];
  if (value?.status !== 'ok' || value.value.kind !== 'handle') throw new Error('no shape');
  return { host, box: value.value.handle };
}

function fontAt(host: AutomationHost, box: AutomationHandle, start: number, length = 0) {
  const response = host.execute({
    operations: [{ op: 'readFont', shape: box, span: { start, length } }],
  });
  const result = response.results[0];
  if (result?.status !== 'ok') throw new Error(JSON.stringify(result));
  return (result.value as Extract<AutomationValue, { kind: 'font' }>).font;
}

describe('automation host: the font of an empty range', () => {
  test('follows the run before the point, or the run that starts a paragraph', () => {
    const { host, box } = setup();
    // At a boundary between two runs, the run before the point wins: no mixed null.
    expect(fontAt(host, box, 4).bold).toBe(true);
    expect(fontAt(host, box, 4, 1).bold).toBe(false);
    expect(fontAt(host, box, 0).bold).toBe(true);
    expect(fontAt(host, box, 6).bold).toBe(false);
    // Right after a line break, the run after it.
    expect(fontAt(host, box, 10).size).toBe(40);
    // An empty paragraph reads its end properties.
    expect(fontAt(host, box, 14).italic).toBe(false);
    // The start of the last paragraph reads its first run.
    expect(fontAt(host, box, 15).italic).toBe(true);
    // A range that is not empty still reads mixed values as null.
    expect(fontAt(host, box, 0, 9).bold).toBeNull();
  });

  test('a range that ends at the start of a later paragraph reads only the text it holds', () => {
    const { host, box } = setup(para(run('Hello', 'b="1"')) + para(run('World')));
    // "Hello\r": the paragraph break ends the range, so "World" is not in it.
    expect(fontAt(host, box, 0, 6).bold).toBe(true);
    // Only the break: the end of "Hello".
    expect(fontAt(host, box, 5, 1).bold).toBe(true);
    expect(fontAt(host, box, 0, 7).bold).toBeNull();
    // The same rule after an empty paragraph: "Big" through the empty paragraph's start.
    const other = setup();
    expect(fontAt(other.host, other.box, 10, 4).size).toBe(40);
  });
});
