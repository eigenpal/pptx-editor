// `readSlideNotes` and `setSlideNotes`: the speaker notes of a slide through the automation
// host, over the editor's notes commands (`commands/notes-commands.ts`). The notes are the body
// placeholder of a notes slide (`p:notes`, ECMA-376 Part 1, 19.3.1.26) in its Notes Slide part
// (13.3.5).

import { describe, expect, test } from 'bun:test';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import {
  createServerAutomationHost,
  type AutomationHandle,
  type AutomationHost,
  type AutomationOperation,
  type AutomationValue,
} from '../index.ts';

/** Two slides; the first has notes, the second has none. */
function open(bytes?: Uint8Array): AutomationHost {
  const opened = createServerAutomationHost(
    bytes ??
      buildPptx({
        slides: [
          { layout: 'blank', shapes: '', notes: 'First notes' },
          { layout: 'blank', shapes: '' },
        ],
      })
  );
  if (!opened.ok) throw new Error(JSON.stringify(opened.error));
  return opened.host;
}

function ok(host: AutomationHost, operations: AutomationOperation[]): AutomationValue[] {
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results.map((r) => (r.status === 'ok' ? r.value : { kind: 'null' as const }));
}

function slides(host: AutomationHost): readonly AutomationHandle[] {
  const [value] = ok(host, [{ op: 'listSlides' }]);
  if (value?.kind !== 'handles') throw new Error('not handles');
  return value.handles;
}

const notes = (host: AutomationHost, slide: AutomationHandle) =>
  ok(host, [{ op: 'readSlideNotes', slide }])[0];

describe('automation host: speaker notes', () => {
  test('reads, edits, and creates the notes slide, and the notes survive save and reopen', () => {
    const host = open();
    const [first, second] = slides(host);
    expect(notes(host, first!)).toEqual({ kind: 'text', text: 'First notes' });
    expect(notes(host, second!)).toEqual({ kind: 'text', text: '' });
    ok(host, [
      { op: 'setSlideNotes', slide: first!, text: 'First notes, edited' },
      { op: 'setSlideNotes', slide: second!, text: 'Line one\nLine two' },
    ]);
    expect(notes(host, second!)).toEqual({ kind: 'text', text: 'Line one\nLine two' });
    const saved = host.save();
    if (!saved.ok) throw new Error(saved.error.code);
    expect(readPptx(saved.bytes).ok).toBe(true);
    const again = open(saved.bytes);
    const [a, b] = slides(again);
    expect([notes(again, a!), notes(again, b!)]).toEqual([
      { kind: 'text', text: 'First notes, edited' },
      { kind: 'text', text: 'Line one\nLine two' },
    ]);
  });

  test('refuses text that is not a string, too long, or that XML cannot hold, and changes nothing', () => {
    const host = open();
    const [slide] = slides(host);
    for (const text of [7, 'a\u0000b', 'x'.repeat(1_000_001)]) {
      const refused = host.execute({
        operations: [
          { op: 'setSlideNotes', slide: slide!, text } as unknown as AutomationOperation,
        ],
      });
      expect(refused.results[0]).toMatchObject({
        error: { code: 'invalid-argument', detail: 'text' },
      });
    }
    expect(notes(host, slide!)).toEqual({ kind: 'text', text: 'First notes' });
  });
});
