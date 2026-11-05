// Untrusted values through the automation host: shape ids that the file repeats, lookup keys that
// name prototype members, text that XML 1.0 cannot carry, and picture data over the size cap.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  createServerAutomationHost,
  type AutomationHandle,
  type AutomationHost,
  type AutomationOperation,
  type AutomationValue,
} from '../index.ts';

function open(bytes: Uint8Array): AutomationHost {
  const opened = createServerAutomationHost(bytes);
  if (!opened.ok) throw new Error(JSON.stringify(opened.error));
  return opened.host;
}

function ok(host: AutomationHost, operations: AutomationOperation[]): AutomationValue[] {
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results.map((r) => (r.status === 'ok' ? r.value : { kind: 'null' as const }));
}

/** The error of the refused operation: its code and detail. */
function refused(host: AutomationHost, operations: AutomationOperation[]) {
  const before = host.revision();
  const response = host.execute({ operations });
  expect(response.ok).toBe(false);
  expect(host.revision()).toBe(before);
  const failed = response.results.find((r) => r.status === 'error');
  return failed?.status === 'error' ? failed.error : undefined;
}

function handles(value: AutomationValue | undefined): readonly AutomationHandle[] {
  if (value?.kind !== 'handles') throw new Error(`not handles: ${JSON.stringify(value)}`);
  return value.handles;
}

const box = (id: number, name: string, text: string): string =>
  shape({ id, name, x: 0, y: 0, cx: 914_400, cy: 914_400, text: para(run(text)) });

describe('automation host: untrusted values', () => {
  test('a write to a shape whose id another shape repeats refuses, and changes nothing', () => {
    const host = open(
      buildPptx({
        slides: [
          { layout: 'blank', shapes: box(2, 'A', 'one') + box(2, 'B', 'two') + box(3, 'C', 'x') },
        ],
      })
    );
    const [, list] = ok(host, [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'listShapes', container: { result: 0 } },
    ]);
    const [a, b, c] = handles(list);
    expect(refused(host, [{ op: 'setText', shape: b!, text: 'changed' }])).toMatchObject({
      code: 'unsupported-content',
      detail: 'duplicate-id',
    });
    expect(refused(host, [{ op: 'setTag', owner: a!, key: 'K', value: 'v' }])?.detail).toBe(
      'duplicate-id'
    );
    expect(ok(host, [{ op: 'readText', shape: a! }])[0]).toEqual({ kind: 'text', text: 'one' });
    // A shape with its own id still takes writes.
    ok(host, [{ op: 'setText', shape: c!, text: 'y' }]);
    expect(ok(host, [{ op: 'readText', shape: c! }])[0]).toEqual({ kind: 'text', text: 'y' });
  });

  test('an autoSize that names a prototype member is an invalid argument', () => {
    const host = open(buildPptx({ slides: [{ layout: 'blank', shapes: box(2, 'A', 'one') }] }));
    const [, list] = ok(host, [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'listShapes', container: { result: 0 } },
    ]);
    for (const autoSize of ['toString', 'constructor', '__proto__'])
      expect(
        refused(host, [
          {
            op: 'setTextFrame',
            shape: handles(list)[0]!,
            props: { autoSize: autoSize as 'none' },
          },
        ])
      ).toMatchObject({ code: 'invalid-argument', detail: 'autoSize' });
  });

  test('a graphic frame uri that names a prototype member reads as unsupported', () => {
    const frame =
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="F"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
      '<p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></p:xfrm>' +
      '<a:graphic><a:graphicData uri="constructor"/></a:graphic></p:graphicFrame>';
    const host = open(buildPptx({ slides: [{ layout: 'blank', shapes: frame }] }));
    const [, list] = ok(host, [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'listShapes', container: { result: 0 } },
    ]);
    const [read] = ok(host, [{ op: 'readShape', shape: handles(list)[0]! }]);
    expect(read).toMatchObject({ kind: 'shape', shape: { type: 'unsupported' } });
  });

  test('text with a lone surrogate is an invalid argument', () => {
    const host = open(buildPptx({ slides: [{ layout: 'blank', shapes: box(2, 'A', 'one') }] }));
    const [, list] = ok(host, [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'listShapes', container: { result: 0 } },
    ]);
    expect(
      refused(host, [{ op: 'setText', shape: handles(list)[0]!, text: 'a\uD800b' }])
    ).toMatchObject({ code: 'invalid-argument', detail: 'text' });
  });

  test('picture data longer than the media cap refuses before it decodes', () => {
    const host = open(buildPptx({ slides: [{ layout: 'blank', shapes: box(2, 'A', 'one') }] }));
    const [, list] = ok(host, [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'listShapes', container: { result: 0 } },
    ]);
    // Not valid base64 either: a decode would refuse it too, but only after it allocates.
    const huge = '!'.repeat(80 * 1024 * 1024);
    expect(
      refused(host, [
        { op: 'setFill', shape: handles(list)[0]!, fill: { kind: 'image', base64: huge } },
      ])
    ).toMatchObject({ code: 'invalid-argument', detail: 'base64' });
  });
});
