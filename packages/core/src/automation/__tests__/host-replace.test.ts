// `replaceText`: Replace All through the automation host, over the editor's find
// (`commands/find.ts`). The answer is the count; a refused argument names itself; no match
// changes nothing.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { PptxDocumentStore, readPptx } from '../../store/index.ts';
import { findMatches } from '../commands/find.ts';
import { createStoreAutomationHost, type AutomationOperation } from '../index.ts';

const IN = 914_400;

function setup() {
  const read = readPptx(
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
            text: para(run('Acme', 'b="1"') + run(' and acme')),
          }),
          notes: 'Acme notes',
        },
      ],
    })
  );
  if (!read.ok) throw new Error('fixture');
  const store = new PptxDocumentStore(read.pkg);
  return { store, host: createStoreAutomationHost(store) };
}

const replace = (query: unknown, replacement: unknown, options?: unknown): AutomationOperation =>
  ({
    op: 'replaceText',
    query,
    replacement,
    ...(options === undefined ? {} : { options }),
  }) as AutomationOperation;

describe('replaceText', () => {
  test('answers the count and writes every match in one commit', () => {
    const { host, store } = setup();
    const before = store.revision;
    const response = host.execute({ operations: [replace('acme', 'Globex')] });
    expect(response.ok).toBe(true);
    expect(response.results[0]).toEqual({ status: 'ok', value: { kind: 'number', value: 3 } });
    expect(store.revision).toBe(before + 1);
    expect(findMatches(store.package, 'Globex', { matchCase: true })).toHaveLength(3);
    expect(findMatches(store.package, 'acme')).toHaveLength(0);
  });

  test('passes matchCase and wholeWord to the search', () => {
    const { host } = setup();
    const response = host.execute({
      operations: [replace('acme', 'x', { matchCase: true, wholeWord: true })],
    });
    expect(response.results[0]).toEqual({ status: 'ok', value: { kind: 'number', value: 1 } });
  });

  test('no match answers 0 and changes nothing', () => {
    const { host, store } = setup();
    const before = store.package;
    const response = host.execute({ operations: [replace('Nowhere', 'x')] });
    expect(response.results[0]).toEqual({ status: 'ok', value: { kind: 'number', value: 0 } });
    expect(response.changed).toBe(false);
    expect(store.package).toBe(before);
  });

  test('refuses an argument it cannot honor, and names it', () => {
    const cases: [AutomationOperation, string][] = [
      [replace('', 'x'), 'query'],
      [replace('a\nb', 'x'), 'query'],
      [replace('a\rb', 'x'), 'query'],
      [replace('a'.repeat(1001), 'x'), 'query'],
      [replace('a\u0001', 'x'), 'query'],
      [replace(7, 'x'), 'query'],
      [replace('a', 'x'.repeat(1001)), 'replacement'],
      [replace('a', 'x\u0001'), 'replacement'],
      [replace('a', null), 'replacement'],
      [replace('a', 'x', null), 'find-options'],
      [replace('a', 'x', { matchCase: 'yes' }), 'find-options'],
      [replace('a', 'x', { wholeWord: 1 }), 'find-options'],
    ];
    for (const [op, detail] of cases) {
      const { host, store } = setup();
      const before = store.package;
      const response = host.execute({ operations: [op] });
      expect(response.ok).toBe(false);
      const result = response.results[0];
      expect(result?.status === 'error' && [result.error.code, result.error.detail]).toEqual([
        'invalid-argument',
        detail,
      ]);
      expect(store.package).toBe(before);
    }
  });
});
