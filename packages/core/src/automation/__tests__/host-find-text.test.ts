// `findText`: the matches of a query in one shape's text, through the automation host. The
// spans are in the offsets of `readText`, also after a field, whose shown text can differ in
// length from its stored text. A refused argument names itself.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { PptxDocumentStore, readPptx } from '../../store/index.ts';
import { createStoreAutomationHost, type AutomationOperation } from '../index.ts';

const IN = 914_400;

/** "page <slide number> acme" with a stored field text of two digits, then "Acme". */
function setup() {
  const read = readPptx(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 3 * IN,
            cy: IN,
            text:
              para(
                run('page ') +
                  '<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>12</a:t></a:fld>' +
                  run(' acme')
              ) + para(run('Acme')),
          }),
        },
      ],
    })
  );
  if (!read.ok) throw new Error('fixture');
  return createStoreAutomationHost(new PptxDocumentStore(read.pkg));
}

const shapeRef: AutomationOperation = {
  op: 'getShape',
  container: { result: 0 },
  key: { id: '2' },
};

function withShape(host: ReturnType<typeof setup>, op: Record<string, unknown>) {
  return host.execute({
    operations: [
      { op: 'getSlide', key: { index: 0 } },
      shapeRef,
      { op: 'readText', shape: { result: 1 } },
      { shape: { result: 1 }, ...op } as AutomationOperation,
    ],
  });
}

describe('findText', () => {
  test('answers spans that slice the readText text, also after a field', () => {
    const host = setup();
    const response = withShape(host, { op: 'findText', query: 'acme' });
    const text = response.results[2];
    const found = response.results[3];
    if (text?.status !== 'ok' || text.value.kind !== 'text') throw new Error('readText');
    if (found?.status !== 'ok' || found.value.kind !== 'textSpans') throw new Error('findText');
    const shown = text.value.text;
    expect(found.value.spans).toHaveLength(2);
    expect(found.value.spans.map((s) => shown.slice(s.start, s.start + s.length))).toEqual([
      'acme',
      'Acme',
    ]);
    expect(response.changed).toBe(false);
  });

  test('a span limits the matches to those wholly inside it', () => {
    const host = setup();
    const all = withShape(host, { op: 'findText', query: 'acme' }).results[2];
    if (all?.status !== 'ok' || all.value.kind !== 'text') throw new Error('readText');
    const second = all.value.text.lastIndexOf('Acme');
    const response = withShape(host, {
      op: 'findText',
      query: 'acme',
      span: { start: second - 2, length: 5 },
    });
    expect(response.results[3]).toEqual({
      status: 'ok',
      value: { kind: 'textSpans', spans: [] },
    });
    const whole = withShape(host, {
      op: 'findText',
      query: 'acme',
      span: { start: second, length: 4 },
    });
    expect(whole.results[3]).toEqual({
      status: 'ok',
      value: { kind: 'textSpans', spans: [{ start: second, length: 4 }] },
    });
  });

  test('passes matchCase and wholeWord', () => {
    const host = setup();
    const count = (query: string, options: Record<string, boolean>) => {
      const found = withShape(host, { op: 'findText', query, options }).results[3];
      if (found?.status !== 'ok' || found.value.kind !== 'textSpans') throw new Error('findText');
      return found.value.spans.length;
    };
    // "page 1 acme" and "Acme": "acm" is inside a word, and "acme" is lower case once.
    expect(count('acm', {})).toBe(2);
    expect(count('acm', { wholeWord: true })).toBe(0);
    expect(count('acme', { matchCase: true })).toBe(1);
    expect(count('acme', { wholeWord: true })).toBe(2);
  });

  test('never matches a field', () => {
    const host = setup();
    for (const query of ['page 1', 'page 12', 'e 1 a']) {
      const found = withShape(host, { op: 'findText', query }).results[3];
      expect(found).toEqual({ status: 'ok', value: { kind: 'textSpans', spans: [] } });
    }
    // U+FFFC stands in for a field in the search text, so the query refuses.
    const marker = withShape(host, { op: 'findText', query: '\uFFFC' }).results[3];
    expect(marker?.status === 'error' && marker.error.code).toBe('invalid-argument');
  });

  test('refuses an argument it cannot honor, and names it', () => {
    const host = setup();
    const cases: [Record<string, unknown>, string][] = [
      [{ op: 'findText', query: '' }, 'query'],
      [{ op: 'findText', query: 'a\rb' }, 'query'],
      [{ op: 'findText', query: 'a', options: { matchWildcards: true } }, 'find-options'],
      [{ op: 'findText', query: 'a', span: { start: 0, length: 999 } }, 'length'],
      [{ op: 'findText', query: 'a', span: { start: 0.5, length: 1 } }, 'span'],
    ];
    for (const [op, detail] of cases) {
      const result = withShape(host, op).results[3];
      expect(result?.status === 'error' && [result.error.code, result.error.detail]).toEqual([
        'invalid-argument',
        detail,
      ]);
    }
  });
});
