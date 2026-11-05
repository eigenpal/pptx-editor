// Hyperlink relationships: a link that is removed or replaced takes its relationship with it,
// unless something else in the slide still names it.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { PptxDocumentStore, readPptx, REL, relationshipsOf } from '../../store/index.ts';
import {
  createStoreAutomationHost,
  type AutomationHandle,
  type AutomationHost,
  type AutomationOperation,
  type AutomationValue,
} from '../index.ts';

const IN = 914_400;
const SLIDE = '/ppt/slides/slide1.xml';

function setup(): { host: AutomationHost; store: PptxDocumentStore; box: AutomationHandle } {
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
            text: para(run('Hello World')),
          }),
        },
      ],
    })
  );
  if (!read.ok) throw new Error('fixture');
  const store = new PptxDocumentStore(read.pkg);
  const host = createStoreAutomationHost(store);
  const [, box] = ok(host, [
    { op: 'getSlide', key: { index: 0 } },
    { op: 'getShape', container: { result: 0 }, key: { index: 0 } },
  ]);
  if (box?.kind !== 'handle') throw new Error('no shape');
  return { host, store, box: box.handle };
}

function ok(host: AutomationHost, operations: AutomationOperation[]): AutomationValue[] {
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results.map((r) => (r.status === 'ok' ? r.value : { kind: 'null' as const }));
}

/** The targets of the slide's hyperlink relationships. */
function links(store: PptxDocumentStore): string[] {
  return relationshipsOf(store.package, SLIDE)
    .filter((rel) => rel.type === REL.hyperlink)
    .map((rel) => rel.externalTarget ?? '')
    .sort();
}

describe('automation host: hyperlink relationships', () => {
  test('a text link that is replaced or removed leaves no unused relationship', () => {
    const { host, store, box } = setup();
    const span = { start: 0, length: 5 };
    ok(host, [{ op: 'setHyperlink', shape: box, span, address: 'https://a.example/' }]);
    expect(links(store)).toEqual(['https://a.example/']);
    ok(host, [{ op: 'setHyperlink', shape: box, span, address: 'https://b.example/' }]);
    expect(links(store)).toEqual(['https://b.example/']);
    ok(host, [{ op: 'setHyperlink', shape: box, span, address: null }]);
    expect(links(store)).toEqual([]);
  });

  test('a relationship that another run still names stays', () => {
    const { host, store, box } = setup();
    ok(host, [{ op: 'setHyperlink', shape: box, address: 'https://a.example/' }]);
    // Remove the link from the first word only: the second word still names it.
    ok(host, [{ op: 'setHyperlink', shape: box, span: { start: 0, length: 5 }, address: null }]);
    expect(links(store)).toEqual(['https://a.example/']);
    ok(host, [{ op: 'setHyperlink', shape: box, address: null }]);
    expect(links(store)).toEqual([]);
  });

  test('a shape link that is replaced or removed leaves no unused relationship', () => {
    const { host, store, box } = setup();
    ok(host, [{ op: 'setShapeHyperlink', shape: box, address: 'https://a.example/' }]);
    ok(host, [{ op: 'setShapeHyperlink', shape: box, address: 'https://b.example/' }]);
    expect(links(store)).toEqual(['https://b.example/']);
    ok(host, [{ op: 'setShapeHyperlink', shape: box, address: null }]);
    expect(links(store)).toEqual([]);
  });
});

describe('automation host: the text links of a range', () => {
  /** "Hello World" with "Hello" and "World" linked apart, and a shape link on the box. */
  function linked(): { host: AutomationHost; box: AutomationHandle } {
    const { host, box } = setup();
    ok(host, [
      { op: 'setHyperlink', shape: box, span: { start: 0, length: 5 }, address: 'https://a/' },
      { op: 'setHyperlink', shape: box, span: { start: 6, length: 5 }, address: 'https://b/' },
      { op: 'setShapeHyperlink', shape: box, address: 'https://shape/' },
    ]);
    return { host, box };
  }

  function spans(host: AutomationHost, op: AutomationOperation): [string | null, number | null][] {
    const [value] = ok(host, [op]);
    if (value?.kind !== 'hyperlinks') throw new Error('not hyperlinks');
    return value.hyperlinks.map((l) => [l.address, l.start]);
  }

  test('lists the text links of the whole text, without the shape link', () => {
    const { host, box } = linked();
    expect(spans(host, { op: 'listTextHyperlinks', shape: box })).toEqual([
      ['https://a/', 0],
      ['https://b/', 6],
    ]);
  });

  test('lists only the links that overlap a span, each with its whole span', () => {
    const { host, box } = linked();
    const [value] = ok(host, [
      { op: 'listTextHyperlinks', shape: box, span: { start: 3, length: 2 } },
    ]);
    if (value?.kind !== 'hyperlinks') throw new Error('not hyperlinks');
    expect(value.hyperlinks.map((l) => [l.type, l.start, l.length])).toEqual([['textRange', 0, 5]]);
    expect(value.hyperlinks[0]!.shape).toEqual(box as AutomationHandle<'shape'>);
    // The space between the links touches neither.
    const space = { start: 5, length: 1 };
    expect(spans(host, { op: 'listTextHyperlinks', shape: box, span: space })).toEqual([]);
    // An empty span holds the link around its point.
    const point = { start: 7, length: 0 };
    expect(spans(host, { op: 'listTextHyperlinks', shape: box, span: point })).toEqual([
      ['https://b/', 6],
    ]);
  });

  test('refuses a span outside the text', () => {
    const { host, box } = linked();
    const response = host.execute({
      operations: [{ op: 'listTextHyperlinks', shape: box, span: { start: 9, length: 5 } }],
    });
    expect(response.ok).toBe(false);
    expect(response.results[0]).toMatchObject({
      status: 'error',
      error: { code: 'invalid-argument' },
    });
  });
});
