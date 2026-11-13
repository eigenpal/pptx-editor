import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, find, readPptx, xmlPartOf } from '../../store/index.ts';
import {
  createServerAutomationHost,
  type AutomationBatchResponse,
  type AutomationHandle,
  type AutomationHost,
  type AutomationOperation,
  type AutomationValue,
} from '../index.ts';

const IN = 914_400;

function deck(): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({
            id: 2,
            name: 'Box',
            x: IN,
            y: IN,
            cx: 2 * IN,
            cy: IN,
            fill: solid('4472C4'),
            text: para(run('Hello', 'b="1" sz="2400"')) + para(run('World')),
          }) + shape({ id: 3, name: 'Other', x: 4 * IN, y: IN, cx: IN, cy: IN }),
      },
      { layout: 'title', shapes: '' },
    ],
  });
}

function open(bytes = deck()): AutomationHost {
  const opened = createServerAutomationHost(bytes);
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

function handle(value: AutomationValue | undefined): AutomationHandle {
  if (value?.kind !== 'handle') throw new Error(`not a handle: ${JSON.stringify(value)}`);
  return value.handle;
}

function handles(value: AutomationValue | undefined): readonly AutomationHandle[] {
  if (value?.kind !== 'handles') throw new Error(`not handles: ${JSON.stringify(value)}`);
  return value.handles;
}

/** The first shape of the first slide. */
function firstShape(host: AutomationHost): AutomationHandle {
  const [, box] = ok(host, [
    { op: 'getSlide', key: { index: 0 } },
    { op: 'getShape', container: { result: 0 }, key: { index: 0 } },
  ]);
  return handle(box);
}

function reopen(host: AutomationHost): AutomationHost {
  const saved = host.save();
  if (!saved.ok) throw new Error(saved.error.code);
  expect(readPptx(saved.bytes).ok).toBe(true);
  return open(saved.bytes);
}

describe('automation host: batches', () => {
  test('a batch is one transaction: a refused operation commits nothing', () => {
    const host = open();
    const before = host.revision();
    const response = exec(host, [{ op: 'addSlide' }, { op: 'getSlide', key: { index: 99 } }]);
    expect(response.ok).toBe(false);
    expect(response.results[0]).toEqual({ status: 'skipped' });
    expect(response.results[1]).toMatchObject({
      status: 'error',
      error: { code: 'item-not-found' },
    });
    expect(host.revision()).toBe(before);
    expect(handles(ok(host, [{ op: 'listSlides' }])[0])).toHaveLength(2);
  });

  test('result references address what an earlier operation answered', () => {
    const host = open();
    const values = ok(host, [
      { op: 'addSlide', index: 0 },
      { op: 'addShape', container: { result: 0 }, shape: { kind: 'textBox', text: 'New' } },
      { op: 'readText', shape: { result: 1 } },
      { op: 'readSlide', slide: { result: 0 } },
    ]);
    expect(values[2]).toEqual({ kind: 'text', text: 'New' });
    expect(values[3]).toMatchObject({ kind: 'slide', slide: { index: 0 } });
  });

  test('a null lookup answers null to reads and refuses writes', () => {
    const host = open();
    const reads = ok(host, [
      { op: 'getSlide', key: { id: '9999' }, orNull: true },
      { op: 'readSlide', slide: { result: 0 } },
    ]);
    expect(reads).toEqual([{ kind: 'null' }, { kind: 'null' }]);
    const write = exec(host, [
      { op: 'getSlide', key: { id: '9999' }, orNull: true },
      { op: 'deleteSlide', slide: { result: 0 } },
    ]);
    expect(write.ok).toBe(false);
    expect(write.results[1]).toMatchObject({ error: { code: 'invalid-handle' } });
  });

  test('handles are stable and scoped to one host', () => {
    const a = open();
    const b = open();
    const first = handle(ok(a, [{ op: 'getSlide', key: { index: 0 } }])[0]);
    expect(handle(ok(a, [{ op: 'getSlide', key: { index: 0 } }])[0])).toEqual(first);
    const foreign = exec(b, [{ op: 'readSlide', slide: first }]);
    expect(foreign.results[0]).toMatchObject({ error: { code: 'invalid-handle' } });
  });

  test('expectedRevision refuses a batch planned against an older read', () => {
    const host = open();
    const at = host.revision();
    ok(host, [{ op: 'addSlide' }]);
    const stale = host.execute({ operations: [{ op: 'listSlides' }], expectedRevision: at });
    expect(stale.results[0]).toMatchObject({ error: { code: 'stale-revision' } });
  });

  test('a server host has no selection', () => {
    const host = open();
    expect(host.capabilities.selection).toBe(false);
    expect(exec(host, [{ op: 'readSelection' }]).results[0]).toMatchObject({
      error: { code: 'unsupported-capability' },
    });
  });

  test('an unknown operation is refused by name', () => {
    const host = open();
    const response = host.execute({
      operations: [{ op: 'nope' } as unknown as AutomationOperation],
    });
    expect(response.results[0]).toMatchObject({ error: { code: 'unknown-operation' } });
  });

  test('a disposed host refuses everything', () => {
    const host = open();
    host.dispose();
    expect(exec(host, [{ op: 'listSlides' }]).results[0]).toMatchObject({
      error: { code: 'disposed' },
    });
    expect(host.save().ok).toBe(false);
  });

  test('events fire once per committed batch', () => {
    const host = open();
    const seen: number[] = [];
    const off = host.subscribe((e) => seen.push(e.revision));
    ok(host, [{ op: 'addSlide' }, { op: 'addSlide' }]);
    ok(host, [{ op: 'listSlides' }]);
    off();
    expect(seen).toHaveLength(1);
  });
});

describe('automation host: slides, layouts, and masters', () => {
  test('add, move, and delete slides, and read their ids and layouts', () => {
    const host = open();
    const [added] = ok(host, [{ op: 'addSlide', index: 1 }]);
    const [read] = ok(host, [{ op: 'readSlide', slide: handle(added) }]);
    expect(read).toMatchObject({ kind: 'slide', slide: { index: 1 } });
    ok(host, [{ op: 'moveSlide', slide: handle(added), index: 2 }]);
    expect(ok(host, [{ op: 'readSlide', slide: handle(added) }])[0]).toMatchObject({
      slide: { index: 2 },
    });
    ok(host, [{ op: 'deleteSlide', slide: handle(added) }]);
    expect(handles(ok(host, [{ op: 'listSlides' }])[0])).toHaveLength(2);
    expect(exec(host, [{ op: 'readSlide', slide: handle(added) }]).results[0]).toMatchObject({
      error: { code: 'invalid-handle' },
    });
  });

  test('masters and layouts list in file order, and a slide takes a chosen layout', () => {
    const host = open();
    const [masters] = ok(host, [{ op: 'listMasters' }]);
    const master = handles(masters)[0]!;
    const [layouts, info] = ok(host, [
      { op: 'listLayouts', master },
      { op: 'readMaster', master },
    ]);
    expect(handles(layouts).length).toBeGreaterThanOrEqual(2);
    expect(info).toMatchObject({ kind: 'layout', layout: { type: null } });
    const layout = handles(layouts)[1]!;
    const [slide, read] = ok(host, [
      { op: 'addSlide', layout },
      { op: 'readSlide', slide: { result: 0 } },
    ]);
    expect(handle(slide).kind).toBe('slide');
    expect(read).toMatchObject({ slide: { layout } });
    const [layoutRead] = ok(host, [{ op: 'readLayout', layout }]);
    expect(layoutRead).toMatchObject({ kind: 'layout', layout: { master } });
  });

  test('applyLayout changes a slide layout', () => {
    const host = open();
    const [masters] = ok(host, [{ op: 'listMasters' }]);
    const [layouts] = ok(host, [{ op: 'listLayouts', master: handles(masters)[0]! }]);
    const target = handles(layouts)[0]!;
    const [slide] = ok(host, [{ op: 'getSlide', key: { index: 0 } }]);
    ok(host, [{ op: 'applyLayout', slide: handle(slide), layout: target }]);
    expect(ok(host, [{ op: 'readSlide', slide: handle(slide) }])[0]).toMatchObject({
      slide: { layout: target },
    });
  });

  test('exportSlides writes a presentation of the listed slides only', () => {
    const host = open();
    const [slide] = ok(host, [{ op: 'getSlide', key: { index: 1 } }]);
    const [exported] = ok(host, [{ op: 'exportSlides', slides: [handle(slide)] }]);
    if (exported?.kind !== 'base64') throw new Error('not base64');
    const bytes = Uint8Array.from(atob(exported.base64), (c) => c.charCodeAt(0));
    const copy = open(bytes);
    expect(handles(ok(copy, [{ op: 'listSlides' }])[0])).toHaveLength(1);
    // The source is untouched.
    expect(handles(ok(host, [{ op: 'listSlides' }])[0])).toHaveLength(2);
  });

  test('page setup reads and writes the slide size in points', () => {
    const host = open();
    expect(ok(host, [{ op: 'readPresentation' }])[0]).toMatchObject({
      presentation: { slideWidth: 960, slideHeight: 540 },
    });
    ok(host, [{ op: 'setPageSetup', slideWidth: 720 }]);
    expect(ok(reopen(host), [{ op: 'readPresentation' }])[0]).toMatchObject({
      presentation: { slideWidth: 720 },
    });
  });
});

describe('automation host: shapes', () => {
  test('reads a shape in points with its effective values', () => {
    const host = open();
    const [read] = ok(host, [{ op: 'readShape', shape: firstShape(host) }]);
    expect(read).toMatchObject({
      kind: 'shape',
      shape: {
        id: '2',
        name: 'Box',
        type: 'geometricShape',
        left: 72,
        top: 72,
        width: 144,
        height: 72,
        visible: true,
        zOrderPosition: 0,
        level: 0,
        geometry: 'rect',
        parentGroup: null,
      },
    });
  });

  test('adds every kind of shape with a box in points', () => {
    const host = open();
    const [slide] = ok(host, [{ op: 'getSlide', key: { index: 1 } }]);
    const box = { left: 10, top: 20, width: 100, height: 50 };
    const values = ok(host, [
      {
        op: 'addShape',
        container: handle(slide),
        shape: { kind: 'geometric', geometry: 'ellipse' },
        box,
      },
      {
        op: 'addShape',
        container: handle(slide),
        shape: { kind: 'line', connector: 'elbow' },
        box,
      },
      { op: 'addShape', container: handle(slide), shape: { kind: 'textBox', text: 'A\rB' }, box },
      {
        op: 'addShape',
        container: handle(slide),
        shape: { kind: 'table', rows: 2, columns: 3, values: [['a', 'b', 'c']] },
        box,
      },
      { op: 'readShape', shape: { result: 0 } },
      { op: 'readShape', shape: { result: 1 } },
      { op: 'readShape', shape: { result: 2 } },
      { op: 'readShape', shape: { result: 3 } },
      { op: 'readTable', shape: { result: 3 } },
    ]);
    expect(values[4]).toMatchObject({
      shape: {
        type: 'geometricShape',
        geometry: 'ellipse',
        left: 10,
        top: 20,
        width: 100,
        height: 50,
      },
    });
    expect(values[5]).toMatchObject({ shape: { type: 'line' } });
    expect(values[6]).toMatchObject({ shape: { type: 'textBox', left: 10 } });
    expect(values[7]).toMatchObject({ shape: { type: 'table' } });
    expect(values[8]).toMatchObject({
      table: {
        rowCount: 2,
        columnCount: 3,
        values: [
          ['a', 'b', 'c'],
          ['', '', ''],
        ],
      },
    });
    const saved = reopen(host);
    const [list] = ok(saved, [
      { op: 'getSlide', key: { index: 1 } },
      { op: 'listShapes', container: { result: 0 } },
    ]).slice(1);
    expect(handles(list)).toHaveLength(4);
  });

  test('refuses a geometry that is not an ECMA-376 preset', () => {
    const host = open();
    const [slide] = ok(host, [{ op: 'getSlide', key: { index: 0 } }]);
    const response = exec(host, [
      {
        op: 'addShape',
        container: handle(slide),
        shape: { kind: 'geometric', geometry: 'notAShape' },
      },
    ]);
    expect(response.ok).toBe(false);
  });

  test('moves, resizes, renames, hides, rotates, and describes a shape', () => {
    const host = open();
    const box = firstShape(host);
    ok(host, [
      {
        op: 'setShape',
        shape: box,
        props: {
          left: 36,
          top: 18,
          width: 200,
          name: 'Renamed',
          visible: false,
          rotation: 45,
          altTextTitle: 'Title',
          altTextDescription: 'Description',
          isDecorative: true,
        },
      },
    ]);
    const [read] = ok(reopen(host), [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'getShape', container: { result: 0 }, key: { id: '2' } },
      { op: 'readShape', shape: { result: 1 } },
    ]).slice(2);
    expect(read).toMatchObject({
      shape: {
        name: 'Renamed',
        left: 36,
        top: 18,
        width: 200,
        height: 72,
        visible: false,
        rotation: 45,
        altTextTitle: 'Title',
        altTextDescription: 'Description',
        isDecorative: true,
      },
    });
  });

  test('changes the z-order and deletes a shape', () => {
    const host = open();
    const box = firstShape(host);
    ok(host, [{ op: 'setZOrder', shape: box, position: 'bringToFront' }]);
    expect(ok(host, [{ op: 'readShape', shape: box }])[0]).toMatchObject({
      shape: { zOrderPosition: 1 },
    });
    // Already in front: not an error.
    ok(host, [{ op: 'setZOrder', shape: box, position: 'bringForward' }]);
    ok(host, [{ op: 'deleteShape', shape: box }]);
    expect(exec(host, [{ op: 'readShape', shape: box }]).results[0]).toMatchObject({
      error: { code: 'invalid-handle' },
    });
  });

  test('groups and ungroups shapes', () => {
    const host = open();
    const [, all] = ok(host, [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'listShapes', container: { result: 0 } },
    ]);
    const [group, members, first] = ok(host, [
      { op: 'group', shapes: [...handles(all)] },
      { op: 'listShapes', container: { result: 0 } },
      { op: 'readShape', shape: handles(all)[0]! },
    ]);
    expect(handles(members)).toHaveLength(2);
    expect(first).toMatchObject({ shape: { level: 1, parentGroup: handle(group), left: 72 } });
    const [ungrouped] = ok(host, [{ op: 'ungroup', group: handle(group) }]);
    expect(handles(ungrouped)).toHaveLength(2);
  });

  test('parent lookups answer the owner or null', () => {
    const host = open();
    const box = firstShape(host);
    const values = ok(host, [
      { op: 'parentOf', shape: box, kind: 'slide' },
      { op: 'parentOf', shape: box, kind: 'slideLayout', orNull: true },
    ]);
    expect(handle(values[0]).kind).toBe('slide');
    expect(values[1]).toEqual({ kind: 'null' });
  });
});

describe('automation host: fill and line', () => {
  test('reads, sets, and clears a fill', () => {
    const host = open();
    const box = firstShape(host);
    expect(ok(host, [{ op: 'readFill', shape: box }])[0]).toEqual({
      kind: 'fill',
      fill: { type: 'solid', foregroundColor: '#4472C4', transparency: 0 },
    });
    ok(host, [
      { op: 'setFill', shape: box, fill: { kind: 'solid', color: '#FF0000', transparency: 0.25 } },
    ]);
    expect(ok(host, [{ op: 'readFill', shape: box }])[0]).toMatchObject({
      fill: { type: 'solid', foregroundColor: '#FF0000', transparency: 0.25 },
    });
    ok(host, [{ op: 'setFill', shape: box, fill: { kind: 'transparency', transparency: 0.5 } }]);
    expect(ok(host, [{ op: 'readFill', shape: box }])[0]).toMatchObject({
      fill: { foregroundColor: '#FF0000', transparency: 0.5 },
    });
    ok(host, [{ op: 'setFill', shape: box, fill: { kind: 'clear' } }]);
    expect(
      ok(reopen(host), [
        { op: 'getSlide', key: { index: 0 } },
        { op: 'getShape', container: { result: 0 }, key: { id: '2' } },
        { op: 'readFill', shape: { result: 1 } },
      ])[2]
    ).toMatchObject({ fill: { type: 'noFill' } });
  });

  test('a picture fill adds a media part and reads as a picture', () => {
    const host = open();
    const box = firstShape(host);
    // A 1 by 1 PNG.
    const png =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
    ok(host, [{ op: 'setFill', shape: box, fill: { kind: 'image', base64: png } }]);
    const saved = host.save();
    if (!saved.ok) throw new Error(saved.error.code);
    const pkg = readPptx(saved.bytes);
    if (!pkg.ok) throw new Error('the saved presentation does not read back');
    expect([...pkg.pkg.parts.keys()].some((p) => p.startsWith('/ppt/media/image'))).toBe(true);
    expect(
      ok(open(saved.bytes), [
        { op: 'getSlide', key: { index: 0 } },
        { op: 'getShape', container: { result: 0 }, key: { id: '2' } },
        { op: 'readFill', shape: { result: 1 } },
      ])[2]
    ).toMatchObject({ fill: { type: 'pictureAndTexture' } });
    for (const base64 of ['', 'not base64!', 'R0lG']) {
      const response = exec(host, [{ op: 'setFill', shape: box, fill: { kind: 'image', base64 } }]);
      expect(response.results[0]).toMatchObject({ error: { code: 'invalid-argument' } });
    }
  });

  test('refuses a color that is not hex or a known name', () => {
    const host = open();
    const response = exec(host, [
      { op: 'setFill', shape: firstShape(host), fill: { kind: 'solid', color: 'url(x)' } },
    ]);
    expect(response.results[0]).toMatchObject({ error: { code: 'invalid-argument' } });
  });

  test('sets the outline color, weight, dash, and visibility', () => {
    const host = open();
    const box = firstShape(host);
    ok(host, [
      {
        op: 'setLine',
        shape: box,
        line: { color: '#00FF00', weight: 3, dashStyle: 'dash', style: 'dbl' },
      },
    ]);
    expect(ok(host, [{ op: 'readLine', shape: box }])[0]).toMatchObject({
      line: { visible: true, color: '#00FF00', weight: 3, dashStyle: 'dash', style: 'dbl' },
    });
    ok(host, [{ op: 'setLine', shape: box, line: { visible: false } }]);
    expect(ok(host, [{ op: 'readLine', shape: box }])[0]).toMatchObject({
      line: { visible: false },
    });
    expect(reopen(host)).toBeDefined();
  });
});

describe('automation host: text', () => {
  test('reads text with \\r between paragraphs, and replaces a span', () => {
    const host = open();
    const box = firstShape(host);
    expect(ok(host, [{ op: 'readText', shape: box }])[0]).toEqual({
      kind: 'text',
      text: 'Hello\rWorld',
    });
    expect(ok(host, [{ op: 'readText', shape: box, span: { start: 6, length: 5 } }])[0]).toEqual({
      kind: 'text',
      text: 'World',
    });
    ok(host, [{ op: 'setText', shape: box, span: { start: 1, length: 4 }, text: 'i' }]);
    expect(ok(host, [{ op: 'readText', shape: box }])[0]).toEqual({
      kind: 'text',
      text: 'Hi\rWorld',
    });
  });

  test('writes paragraphs and line breaks, and they survive a save', () => {
    const host = open();
    const box = firstShape(host);
    ok(host, [{ op: 'setText', shape: box, text: 'One\nTwo\vthree\r\nFour' }]);
    const [text] = ok(reopen(host), [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'getShape', container: { result: 0 }, key: { id: '2' } },
      { op: 'readText', shape: { result: 1 } },
    ]).slice(2);
    expect(text).toEqual({ kind: 'text', text: 'One\rTwo\vthree\rFour' });
  });

  test('refuses a span outside the text and control characters', () => {
    const host = open();
    const box = firstShape(host);
    expect(
      exec(host, [{ op: 'readText', shape: box, span: { start: 3, length: 100 } }]).results[0]
    ).toMatchObject({
      error: { code: 'invalid-argument' },
    });
    expect(exec(host, [{ op: 'setText', shape: box, text: 'a\u0001b' }]).results[0]).toMatchObject({
      error: { code: 'invalid-argument' },
    });
  });

  test('reads the effective font, null where runs differ, and writes font properties', () => {
    const host = open();
    const box = firstShape(host);
    const [first, all] = ok(host, [
      { op: 'readFont', shape: box, span: { start: 0, length: 5 } },
      { op: 'readFont', shape: box },
    ]);
    expect(first).toMatchObject({ font: { bold: true, size: 24 } });
    expect(all).toMatchObject({ font: { bold: null, size: null } });
    ok(host, [
      {
        op: 'setFont',
        shape: box,
        span: { start: 6, length: 5 },
        font: {
          bold: true,
          italic: true,
          size: 30,
          color: '#112233',
          name: 'Georgia',
          underline: 'sng',
          allCaps: true,
        },
      },
    ]);
    expect(
      ok(host, [{ op: 'readFont', shape: box, span: { start: 6, length: 5 } }])[0]
    ).toMatchObject({
      font: {
        bold: true,
        italic: true,
        size: 30,
        color: '#112233',
        name: 'Georgia',
        underline: 'sng',
        allCaps: true,
      },
    });
    ok(host, [{ op: 'setFont', shape: box, font: { superscript: true, strikethrough: true } }]);
    expect(ok(host, [{ op: 'readFont', shape: box }])[0]).toMatchObject({
      font: { superscript: true, subscript: false, strikethrough: true },
    });
  });

  test('a font name applies to the Latin, East Asian, and complex script faces', () => {
    const host = open();
    const box = firstShape(host);
    ok(host, [{ op: 'setFont', shape: box, font: { name: 'Noto Sans TC' } }]);
    const saved = host.save();
    if (!saved.ok) throw new Error(saved.error.code);
    const pkg = readPptx(saved.bytes);
    if (!pkg.ok) throw new Error('unreadable');
    const slide = xmlPartOf(pkg.pkg, '/ppt/slides/slide1.xml')!.root;
    const rPr = child(find(slide, 'a:r')!, 'a:rPr');
    for (const face of ['a:latin', 'a:ea', 'a:cs'])
      expect(attr(child(rPr, face), 'typeface')).toBe('Noto Sans TC');
  });

  test('a font name maps a theme font to each face, and refuses other references', () => {
    const host = open();
    const box = firstShape(host);
    for (const name of ['+mj-ea', '+foo', '   '])
      expect(exec(host, [{ op: 'setFont', shape: box, font: { name } }]).ok).toBe(false);
    ok(host, [{ op: 'setFont', shape: box, font: { name: '+mj-lt' } }]);
    const saved = host.save();
    if (!saved.ok) throw new Error(saved.error.code);
    const pkg = readPptx(saved.bytes);
    if (!pkg.ok) throw new Error('unreadable');
    const slide = xmlPartOf(pkg.pkg, '/ppt/slides/slide1.xml')!.root;
    const rPr = child(find(slide, 'a:r')!, 'a:rPr');
    expect(['a:latin', 'a:ea', 'a:cs'].map((f) => attr(child(rPr, f), 'typeface'))).toEqual([
      '+mj-lt',
      '+mj-ea',
      '+mj-cs',
    ]);
  });

  test('reads and writes alignment, bullets, and the list level', () => {
    const host = open();
    const box = firstShape(host);
    ok(host, [
      {
        op: 'setParagraph',
        shape: box,
        paragraph: { alignment: 'ctr', bulletVisible: true, indentLevel: 2 },
      },
    ]);
    expect(ok(host, [{ op: 'readParagraph', shape: box }])[0]).toMatchObject({
      paragraph: { alignment: 'ctr', bulletVisible: true, bulletType: 'char', indentLevel: 2 },
    });
    ok(host, [
      {
        op: 'setParagraph',
        shape: box,
        span: { start: 0, length: 0 },
        paragraph: { bulletStyle: 'romanUcPeriod' },
      },
    ]);
    expect(
      ok(host, [{ op: 'readParagraph', shape: box, span: { start: 0, length: 1 } }])[0]
    ).toMatchObject({
      paragraph: { bulletType: 'autonum', bulletStyle: 'romanUcPeriod' },
    });
    ok(host, [{ op: 'setParagraph', shape: box, paragraph: { bulletVisible: false } }]);
    expect(ok(host, [{ op: 'readParagraph', shape: box }])[0]).toMatchObject({
      paragraph: { bulletVisible: false },
    });
  });

  test('reads and writes the text frame', () => {
    const host = open();
    const box = firstShape(host);
    ok(host, [
      {
        op: 'setTextFrame',
        shape: box,
        props: {
          leftMargin: 10,
          topMargin: 5,
          verticalAlignment: 'ctrCentered',
          wordWrap: false,
          autoSize: 'none',
        },
      },
    ]);
    expect(ok(host, [{ op: 'readTextFrame', shape: box }])[0]).toMatchObject({
      textFrame: {
        leftMargin: 10,
        topMargin: 5,
        rightMargin: 7.2,
        verticalAlignment: 'ctrCentered',
        wordWrap: false,
        autoSize: 'none',
        hasText: true,
      },
    });
  });

  test('a shape with no text frame is refused for text', () => {
    const host = open();
    const [slide] = ok(host, [{ op: 'getSlide', key: { index: 0 } }]);
    const [line] = ok(host, [
      { op: 'addShape', container: handle(slide), shape: { kind: 'line', connector: 'straight' } },
    ]);
    expect(exec(host, [{ op: 'readText', shape: handle(line) }]).results[0]).toMatchObject({
      error: { code: 'unsupported-content' },
    });
  });

  test('sets a text hyperlink and lists it on the slide', () => {
    const host = open();
    const box = firstShape(host);
    ok(host, [
      {
        op: 'setHyperlink',
        shape: box,
        span: { start: 0, length: 5 },
        address: 'https://example.com/',
        screenTip: 'Tip',
      },
    ]);
    const [slide] = ok(host, [{ op: 'getSlide', key: { index: 0 } }]);
    expect(
      ok(reopen(host), [
        { op: 'getSlide', key: { index: 0 } },
        { op: 'listHyperlinks', slide: { result: 0 } },
      ])[1]
    ).toMatchObject({
      hyperlinks: [
        {
          address: 'https://example.com/',
          screenTip: 'Tip',
          type: 'textRange',
          start: 0,
          length: 5,
        },
      ],
    });
    void slide;
    expect(
      exec(host, [{ op: 'setHyperlink', shape: box, address: 'javascript:alert(1)' }]).results[0]
    ).toMatchObject({ error: { code: 'invalid-argument' } });
  });
});

describe('automation host: tags, properties, and tables', () => {
  test('tags on the presentation, a slide, and a shape survive a save', () => {
    const host = open();
    const box = firstShape(host);
    const [slide, presentation] = ok(host, [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'getPresentation' },
    ]);
    ok(host, [
      { op: 'setTag', owner: handle(presentation), key: 'Deck', value: 'Q3' },
      { op: 'setTag', owner: handle(slide), key: 'Status', value: 'Draft' },
      { op: 'setTag', owner: box, key: 'role', value: 'title' },
      { op: 'setTag', owner: box, key: 'ROLE', value: 'body' },
    ]);
    const again = reopen(host);
    const [s, shape0] = ok(again, [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'getShape', container: { result: 0 }, key: { id: '2' } },
    ]);
    const [deckTags, slideTags, shapeTags] = ok(again, [
      { op: 'getPresentation' },
      { op: 'readTags', owner: { result: 0 } },
      { op: 'readTags', owner: handle(s) },
      { op: 'readTags', owner: handle(shape0) },
    ]).slice(1);
    expect(deckTags).toEqual({ kind: 'tags', tags: [{ key: 'DECK', value: 'Q3' }] });
    expect(slideTags).toEqual({ kind: 'tags', tags: [{ key: 'STATUS', value: 'Draft' }] });
    expect(shapeTags).toEqual({ kind: 'tags', tags: [{ key: 'ROLE', value: 'body' }] });
    ok(again, [{ op: 'deleteTag', owner: handle(shape0), key: 'role' }]);
    expect(ok(again, [{ op: 'readTags', owner: handle(shape0) }])[0]).toEqual({
      kind: 'tags',
      tags: [],
    });
  });

  test('document properties read and write through the core and extended parts', () => {
    const host = open();
    ok(host, [
      {
        op: 'setProperties',
        properties: {
          title: 'Plan',
          author: 'Ada',
          keywords: 'q3',
          company: 'Example Ltd',
          revisionNumber: 4,
        },
      },
    ]);
    const again = reopen(host);
    expect(ok(again, [{ op: 'getProperties' }])[0]).toMatchObject({
      properties: {
        title: 'Plan',
        author: 'Ada',
        keywords: 'q3',
        company: 'Example Ltd',
        revisionNumber: 4,
      },
    });
    expect(ok(again, [{ op: 'readPresentation' }])[0]).toMatchObject({
      presentation: { title: 'Plan' },
    });
    expect(
      exec(again, [{ op: 'setProperties', properties: { lastAuthor: 'x' } as never }]).ok
    ).toBe(false);
  });

  test('table cells, rows, columns, merges, and sizes', () => {
    const host = open();
    const [slide] = ok(host, [{ op: 'getSlide', key: { index: 0 } }]);
    const [table] = ok(host, [
      { op: 'addShape', container: handle(slide), shape: { kind: 'table', rows: 2, columns: 2 } },
    ]);
    const t = handle(table);
    ok(host, [
      { op: 'setTableCell', shape: t, row: 0, column: 1, text: 'B1' },
      { op: 'insertTableRows', shape: t, count: 1 },
      { op: 'insertTableColumns', shape: t, index: 0, count: 1 },
    ]);
    expect(ok(host, [{ op: 'readTable', shape: t }])[0]).toMatchObject({
      table: {
        rowCount: 3,
        columnCount: 3,
        values: [
          ['', '', 'B1'],
          ['', '', ''],
          ['', '', ''],
        ],
      },
    });
    ok(host, [
      { op: 'mergeTableCells', shape: t, row: 1, column: 0, rowCount: 2, columnCount: 2 },
      { op: 'setTableSize', shape: t, target: 'column', index: 2, size: 100 },
    ]);
    const [read, cell, outside] = ok(host, [
      { op: 'readTable', shape: t },
      { op: 'readTableCell', shape: t, row: 2, column: 1 },
      { op: 'readTableCell', shape: t, row: 9, column: 0 },
    ]);
    expect(read).toMatchObject({
      table: { mergedAreas: [{ rowIndex: 1, columnIndex: 0, rowCount: 2, columnCount: 2 }] },
    });
    expect((read as unknown as { table: { columnWidths: number[] } }).table.columnWidths[2]).toBe(
      100
    );
    expect(cell).toMatchObject({
      cell: { rowIndex: 1, columnIndex: 0, rowCount: 2, columnCount: 2 },
    });
    expect(outside).toEqual({ kind: 'null' });
    ok(host, [
      { op: 'deleteTableRows', shape: t, index: 0, count: 1 },
      { op: 'deleteTableColumns', shape: t, index: 0, count: 1 },
    ]);
    expect(
      ok(reopen(host), [
        { op: 'getSlide', key: { index: 0 } },
        { op: 'listShapes', container: { result: 0 } },
      ])
    ).toBeDefined();
  });
});

describe('automation host: navigation and shape links', () => {
  test('answers the layout and master of a slide as handles in the same batch', () => {
    const host = open();
    const [, layout, master, read] = ok(host, [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'getSlideLayout', slide: { result: 0 } },
      { op: 'getSlideMaster', owner: { result: 0 } },
      { op: 'readLayout', layout: { result: 1 } },
    ]);
    expect(handle(layout).kind).toBe('slideLayout');
    expect(handle(master).kind).toBe('slideMaster');
    const [layoutMaster] = ok(host, [{ op: 'getSlideMaster', owner: handle(layout) }]);
    expect(layoutMaster).toEqual(master);
    expect(read).toMatchObject({ kind: 'layout', layout: { master: handle(master) } });
  });

  test('adds a slide by layout id and refuses an unknown id', () => {
    const host = open();
    const [, layouts] = ok(host, [
      { op: 'getMaster', key: { index: 0 } },
      { op: 'listLayouts', master: { result: 0 } },
    ]);
    const second = handles(layouts)[1]!;
    const [read] = ok(host, [{ op: 'readLayout', layout: second }]);
    const id = read?.kind === 'layout' ? read.layout.id : '';
    const [slide] = ok(host, [{ op: 'addSlide', layoutId: id }]);
    const [layout] = ok(host, [{ op: 'getSlideLayout', slide: handle(slide) }]);
    expect(layout).toEqual({ kind: 'handle', handle: second });
    const refused = exec(host, [{ op: 'addSlide', layoutId: 'nope' }]);
    expect(refused.results[0]).toMatchObject({ error: { code: 'item-not-found' } });
  });

  test('answers the parent group, or null at the top level', () => {
    const host = open();
    const box = firstShape(host);
    const [top] = ok(host, [{ op: 'getParentGroup', shape: box, orNull: true }]);
    expect(top).toEqual({ kind: 'null' });
    const refused = exec(host, [{ op: 'getParentGroup', shape: box }]);
    expect(refused.results[0]).toMatchObject({ error: { code: 'item-not-found' } });
    const [slide] = ok(host, [{ op: 'getSlide', key: { index: 0 } }]);
    const [other] = ok(host, [{ op: 'getShape', container: handle(slide), key: { id: '3' } }]);
    const [group] = ok(host, [{ op: 'group', shapes: [box, handle(other)] }]);
    expect(ok(host, [{ op: 'getParentGroup', shape: box }])).toEqual([group]);
  });

  test('sets and removes the click link of a whole shape', () => {
    const host = open();
    const box = firstShape(host);
    const [slide] = ok(host, [{ op: 'getSlide', key: { index: 0 } }]);
    ok(host, [
      { op: 'setShapeHyperlink', shape: box, address: 'https://example.com/', screenTip: 'Go' },
    ]);
    const again = reopen(host);
    const [s] = ok(again, [{ op: 'getSlide', key: { index: 0 } }]);
    const [links] = ok(again, [{ op: 'listHyperlinks', slide: handle(s) }]);
    expect(links).toMatchObject({
      hyperlinks: [{ type: 'shape', address: 'https://example.com/', screenTip: 'Go' }],
    });
    ok(host, [{ op: 'setShapeHyperlink', shape: box, address: null }]);
    expect(ok(host, [{ op: 'listHyperlinks', slide: handle(slide) }])[0]).toEqual({
      kind: 'hyperlinks',
      hyperlinks: [],
    });
    const bad = exec(host, [{ op: 'setShapeHyperlink', shape: box, address: 'javascript:x' }]);
    expect(bad.results[0]).toMatchObject({ error: { code: 'invalid-argument' } });
  });
});
