// The section operations of the automation host: the section list is a `p14:sectionLst` in a
// `p:ext` (`CT_Extension`, ECMA-376 Part 1, 19.2.1.11) of the presentation's `p:extLst`
// (`CT_ExtensionList`, 19.2.1.12). Each section names its slides by `p:sldId/@id` (19.7.13).

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

const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const URI = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';
const A = '{6E1B33A1-4C0F-4B53-9C2B-0D4A4B3E8A11}';
const B = '{0B7F5C2D-1E4A-4F3B-8C6D-2A9E7B1C5D22}';

/** Three slides: section A holds 256 and 257, section B holds 258. */
function deck(): Uint8Array {
  return buildPptx({
    slides: [1, 2, 3].map(() => ({ layout: 'blank' as const, shapes: '' })),
    presentationExt:
      `<p:extLst><p:ext uri="${URI}"><p14:sectionLst xmlns:p14="${P14}">` +
      `<p14:section name="A" id="${A}"><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="257"/></p14:sldIdLst></p14:section>` +
      `<p14:section name="B" id="${B}"><p14:sldIdLst><p14:sldId id="258"/></p14:sldIdLst></p14:section>` +
      '</p14:sectionLst></p:ext></p:extLst>',
  });
}

function open(bytes = deck()): AutomationHost {
  const opened = createServerAutomationHost(bytes);
  if (!opened.ok) throw new Error(JSON.stringify(opened.error));
  return opened.host;
}

function ok(host: AutomationHost, operations: AutomationOperation[]): AutomationValue[] {
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results.map((r) => (r.status === 'ok' ? r.value : { kind: 'null' as const }));
}

/** The error of the one failing operation of a refused batch. */
function refusal(host: AutomationHost, operations: AutomationOperation[]) {
  const response = host.execute({ operations });
  expect(response.ok).toBe(false);
  const failed = response.results.find((r) => r.status === 'error');
  return failed?.status === 'error' ? failed.error : undefined;
}

function handle(value: AutomationValue | undefined): AutomationHandle {
  if (value?.kind !== 'handle') throw new Error(`not a handle: ${JSON.stringify(value)}`);
  return value.handle;
}

/** Each section as `name:slide ids`. */
function sections(host: AutomationHost): string[] {
  const [list] = ok(host, [{ op: 'listSections' }]);
  if (list?.kind !== 'handles') throw new Error('not handles');
  return ok(
    host,
    list.handles.map((section) => ({ op: 'readSection', section }))
  ).map((v) => {
    if (v.kind !== 'section') throw new Error('not a section');
    return `${v.section.name}:${v.section.slideIds.join(',')}`;
  });
}

describe('automation host: sections', () => {
  test('a section reads its id, name, index, and slides, and keeps its handle', () => {
    const host = open();
    const [byIndex, byId, read] = ok(host, [
      { op: 'getSection', key: { index: 1 } },
      { op: 'getSection', key: { id: B.toLowerCase() } },
      { op: 'readSection', section: { result: 0 } },
    ]);
    expect(handle(byIndex)).toBe(handle(byId));
    expect(read).toEqual({
      kind: 'section',
      section: { id: B, name: 'B', index: 1, slideIds: ['258'] },
    });
    expect(ok(host, [{ op: 'getSection', key: { id: 'nope' }, orNull: true }])[0]).toEqual({
      kind: 'null',
    });
  });

  test('add, rename, move, and delete each act on the section list, and survive a save', () => {
    const host = open();
    const [, added] = ok(host, [
      { op: 'getSlide', key: { id: '257' } },
      { op: 'addSection', name: 'C', slide: { result: 0 } },
      { op: 'renameSection', section: { result: 1 }, name: 'Middle' },
      { op: 'moveSection', section: { result: 1 }, index: 2 },
    ]);
    expect(sections(host)).toEqual(['A:256', 'B:258', 'Middle:257']);
    const [slides] = ok(host, [{ op: 'listSlides' }]);
    expect(slides?.kind === 'handles' && slides.handles.length).toBe(3);
    ok(host, [{ op: 'deleteSection', section: handle(added), deleteSlides: true }]);
    expect(sections(host)).toEqual(['A:256', 'B:258']);
    const saved = host.save();
    if (!saved.ok) throw new Error(saved.error.code);
    expect(readPptx(saved.bytes).ok).toBe(true);
    expect(sections(open(saved.bytes))).toEqual(['A:256', 'B:258']);
    // The handle of a deleted section is stale.
    expect(refusal(host, [{ op: 'readSection', section: handle(added) }])).toMatchObject({
      code: 'invalid-handle',
    });
  });

  test('bad arguments refuse with invalid-argument, and the batch writes nothing', () => {
    const host = open();
    const cases: AutomationOperation[] = [
      { op: 'renameSection', section: { result: 0 }, name: '' },
      { op: 'renameSection', section: { result: 0 }, name: 7 as unknown as string },
      { op: 'moveSection', section: { result: 0 }, index: 2 },
      { op: 'moveSection', section: { result: 0 }, index: 0.5 },
      { op: 'deleteSection', section: { result: 0 }, deleteSlides: 'yes' as unknown as boolean },
    ];
    for (const op of cases) {
      const error = refusal(host, [
        { op: 'getSection', key: { index: 0 } },
        { op: 'renameSection', section: { result: 0 }, name: 'Z' },
        op,
      ]);
      expect(error).toMatchObject({ code: 'invalid-argument' });
      expect(sections(host)).toEqual(['A:256,257', 'B:258']);
    }
  });

  test('a section handle is not a slide handle', () => {
    const host = open();
    const [section] = ok(host, [{ op: 'getSection', key: { index: 0 } }]);
    expect(refusal(host, [{ op: 'deleteSlide', slide: handle(section) }])).toMatchObject({
      code: 'invalid-handle',
    });
  });
});
