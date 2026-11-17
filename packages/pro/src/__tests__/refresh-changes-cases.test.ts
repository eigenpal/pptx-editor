// Computed document refresh changes (review/refresh-changes.ts) in the cases that review found:
// large text bodies, related parts, repeated ids, writers that number slides again, hidden
// slides, removed notes, wrapped shapes, paragraph insertions and removals, word ranges, the
// change limit, and package changes that slide changes explain. Every computed location must
// pass the editor's own check.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { readPptx, type PptxPackage } from '@pptx-editor.dev/core/store';
import type { ComputedRefreshChange } from '@pptx-editor.dev/core/editor';
import { resolveRefreshChanges } from '../../../core/src/editor/document-refresh-changes.ts';
import {
  buildPptx,
  group,
  para,
  run,
  shape,
  type FixtureOptions,
  type FixtureSlide,
} from '../../../../scripts/fixtures/builder.ts';
import { reviewModule } from '../index.ts';
import { computeRefreshChanges } from '../review/refresh-changes.ts';
import { MAX_CHANGES } from '../review/refresh-common.ts';

const box = (id: number, text: string, x = 400_000, name?: string) =>
  shape({
    id,
    ...(name ? { name } : {}),
    x,
    y: 400_000,
    cx: 4_000_000,
    cy: 800_000,
    text: para(run(text)),
  });
const multi = (id: number, paras: readonly string[]) =>
  shape({ id, x: 1, y: 1, cx: 10, cy: 10, text: paras.map((p) => para(run(p))).join('') });
const pkg = (bytes: Uint8Array): PptxPackage => {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('fixture did not open');
  return r.pkg;
};
const deck = (slides: readonly FixtureSlide[], more: Partial<FixtureOptions> = {}) =>
  buildPptx({ slides, ...more });
const creation = (v: number) =>
  `<p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="${v}"/></p:ext></p:extLst>`;
const edit = (bytes: Uint8Array, f: (files: Record<string, Uint8Array>) => void) => {
  const files = unzipSync(bytes);
  f(files);
  return zipSync(files);
};
const replaceIn = (
  files: Record<string, Uint8Array>,
  path: string,
  from: string | RegExp,
  to: string
) => {
  files[path] = strToU8(strFromU8(files[path]!).replace(from, to));
};

/** The changes, after a check that the editor accepts every location and that ids are unique. */
function changesOf(a: Uint8Array, b: Uint8Array): ComputedRefreshChange[] {
  const returned = pkg(b);
  const { changes } = computeRefreshChanges(pkg(a), returned);
  expectValid(returned, changes);
  return changes;
}

function expectValid(returned: PptxPackage, changes: readonly ComputedRefreshChange[]): void {
  expect(new Set(changes.map((c) => c.id)).size).toBe(changes.length);
  const located = resolveRefreshChanges(returned, 'r', changes, new Map(), () => true);
  for (const l of located) {
    expect([l.change.id, l.change.status]).not.toEqual([l.change.id, 'invalid']);
    const want = changes.find((c) => c.id === l.change.id)!;
    if (want.location) expect([l.change.id, l.change.status]).toEqual([l.change.id, 'available']);
  }
}

const summary = (changes: readonly ComputedRefreshChange[]) =>
  changes.map((c) => [c.computed.scope, c.computed.action, c.computed.aspects ?? []]);

describe('large text bodies', () => {
  test('12000 changed paragraphs give one shape text change without a large table', () => {
    const n = 12_000;
    const make = (tag: string) =>
      deck([
        {
          shapes: multi(
            2,
            Array.from({ length: n }, (_, i) => `${tag}${i}`)
          ),
        },
      ]);
    const started = performance.now();
    const changes = changesOf(make('a'), make('b'));
    expect(performance.now() - started).toBeLessThan(5000);
    expect(summary(changes)).toEqual([['shape', 'changed', ['text']]]);
  });

  test('one changed paragraph among 12000 equal ones is one text change', () => {
    const paras = Array.from({ length: 12_000 }, (_, i) => `p${i}`);
    const after = [...paras];
    after[6000] = 'changed';
    const changes = changesOf(
      deck([{ shapes: multi(2, paras) }]),
      deck([{ shapes: multi(2, after) }])
    );
    expect(changes).toHaveLength(1);
    expect(changes[0]!.location).toMatchObject({ paragraphIndex: 6000, text: 'changed' });
  });
});

const IMG = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const PACKAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package';
const HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const pic = (id: number) =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="P${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
const png = (v: number) => [
  { name: 'image1.png', bytes: new Uint8Array([137, 80, 78, 71, v]), contentType: 'image/png' },
];
const picSlide: FixtureSlide = { shapes: pic(3), rels: [['rIdImg', IMG, '../media/image1.png']] };
const chartFrame = (id: number) =>
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Chart ${id}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame>`;
const chartDeck = (value: string, workbook: number) =>
  deck([{ shapes: chartFrame(4), rels: [['rIdChart', CHART, '../charts/chart1.xml']] }], {
    extraParts: [
      [
        'ppt/charts/chart1.xml',
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
        `<?xml version="1.0" encoding="UTF-8"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>${value}</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart><c:externalData r:id="rId1"/></c:chartSpace>`,
      ],
      [
        'ppt/charts/_rels/chart1.xml.rels',
        'application/vnd.openxmlformats-package.relationships+xml',
        `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${PACKAGE}" Target="../embeddings/book1.xlsx"/></Relationships>`,
      ],
    ],
    binaryParts: [
      {
        name: 'ppt/embeddings/book1.xlsx',
        bytes: new Uint8Array([80, 75, workbook]),
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
    ],
  });

describe('related parts', () => {
  test('picture bytes replaced under the same relationship give a picture change', () => {
    const changes = changesOf(
      deck([picSlide], { media: png(1) }),
      deck([picSlide], { media: png(2) })
    );
    expect(changes).toEqual([
      {
        id: expect.any(String),
        location: { slideIndex: 0, shapeId: 3 },
        computed: { scope: 'shape', action: 'changed', aspects: ['picture'] },
      },
    ]);
  });

  test('a changed chart part or chart workbook gives a chart change and no package change', () => {
    expect(summary(changesOf(chartDeck('1', 1), chartDeck('2', 1)))).toEqual([
      ['shape', 'changed', ['chart']],
    ]);
    expect(summary(changesOf(chartDeck('1', 1), chartDeck('1', 2)))).toEqual([
      ['shape', 'changed', ['chart']],
    ]);
    expect(changesOf(chartDeck('1', 1), chartDeck('1', 1))).toEqual([]);
  });

  test('an added picture or chart reports the shape and no package change', () => {
    expect(
      summary(
        changesOf(
          deck([{ shapes: box(2, 'A') }]),
          deck([{ shapes: box(2, 'A') + pic(3), rels: picSlide.rels }], { media: png(1) })
        )
      )
    ).toEqual([['shape', 'added', []]]);
    const empty = deck([{ shapes: '' }]);
    expect(summary(changesOf(empty, chartDeck('1', 1)))).toEqual([['shape', 'added', []]]);
  });

  test('media named again with the same bytes gives no change', () => {
    const renamed = edit(deck([picSlide], { media: png(1) }), (files) => {
      files['ppt/media/image7.png'] = files['ppt/media/image1.png']!;
      delete files['ppt/media/image1.png'];
      replaceIn(files, 'ppt/slides/_rels/slide1.xml.rels', 'image1.png', 'image7.png');
    });
    expect(changesOf(deck([picSlide], { media: png(1) }), renamed)).toEqual([]);
  });

  test('a changed hyperlink target is a shape change; a removed part is a package change', () => {
    const linked = (url: string) =>
      deck([
        {
          shapes: shape({
            id: 2,
            x: 1,
            y: 1,
            cx: 10,
            cy: 10,
            text: para(run('Link', '', '<a:hlinkClick r:id="rIdLink"/>')),
          }),
          rels: [['rIdLink', HYPERLINK, url, 'External']],
        },
      ]);
    expect(summary(changesOf(linked('https://a.example/'), linked('https://b.example/')))).toEqual([
      ['shape', 'changed', ['other']],
    ]);
    const extra = deck([{ shapes: box(2, 'A') }], {
      extraParts: [['ppt/custom/data1.xml', 'application/xml', '<?xml version="1.0"?><d/>']],
    });
    expect(summary(changesOf(extra, deck([{ shapes: box(2, 'A') }])))).toEqual([
      ['package', 'changed', []],
    ]);
  });

  test('a section list that follows an added slide gives no package change', () => {
    const sections = (ids: readonly number[]) =>
      `<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p14:section name="Main" id="{6D5B0C0E-3E5D-4C59-9E0B-6B8F0B4E1F00}"><p14:sldIdLst>${ids.map((id) => `<p14:sldId id="${id}"/>`).join('')}</p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst>`;
    const a = deck([{ shapes: box(2, 'One') }], { presentationExt: sections([256]) });
    const b = deck([{ shapes: box(2, 'One') }, { shapes: box(2, 'Two') }], {
      presentationExt: sections([256, 257]),
    });
    expect(summary(changesOf(a, b))).toEqual([['slide', 'added', []]]);
    const renamed = edit(b, (files) =>
      replaceIn(files, 'ppt/presentation.xml', 'name="Main"', 'name="Other"')
    );
    expect(summary(changesOf(a, renamed))).toContainEqual(['package', 'changed', []]);
  });
});

describe('repeated ids', () => {
  test('a repeated shape id gives slide locations and unique change ids', () => {
    const changes = changesOf(
      deck([{ shapes: box(2, 'A', 1, 'S') + box(2, 'B', 2, 'S') }]),
      deck([{ shapes: box(2, 'A2', 1, 'S') + box(2, 'B2', 2, 'S') }])
    );
    expect(changes).toHaveLength(2);
    for (const c of changes) expect(c.location).toEqual({ slideIndex: 0 });
    const added = changesOf(
      deck([{ shapes: box(2, 'A') }]),
      deck([{ shapes: box(2, 'A') + box(5, 'x', 1, 'N') + box(5, 'y', 2, 'N') }])
    );
    expect(summary(added)).toEqual([
      ['shape', 'added', []],
      ['shape', 'added', []],
    ]);
  });

  test('a top-level id that a group repeats gives a slide location', () => {
    const g = (t: string) => group({ id: 10, x: 0, y: 0, cx: 100, cy: 100 }, box(11, t));
    const changes = changesOf(
      deck([{ shapes: box(11, 'top') + g('in') }]),
      deck([{ shapes: box(11, 'top2') + g('in') }])
    );
    expect(changes).toEqual([
      {
        id: expect.any(String),
        location: { slideIndex: 0 },
        computed: { scope: 'shape', action: 'changed', aspects: ['text'] },
      },
    ]);
  });

  test('a duplicated slide that keeps its creation id gets its own change id', () => {
    const ext = creation(111);
    const changes = changesOf(
      deck([{ shapes: box(2, 'A'), ext }]),
      deck([
        { shapes: box(2, 'A2'), ext },
        { shapes: box(2, 'A'), ext },
      ])
    );
    expect(changes.map((c) => c.id)).toEqual(['text:111:2:0', 'slide:111#2']);
  });
});

describe('writers that number slide ids again', () => {
  test('an insertion before an edited slide keeps the text edit', () => {
    const changes = changesOf(
      deck([{ shapes: box(2, 'One') }, { shapes: box(2, 'Hello there') }]),
      deck([
        { shapes: box(2, 'Intro') },
        { shapes: box(2, 'One') },
        { shapes: box(2, 'Hello world') },
      ])
    );
    expect(summary(changes)).toEqual([
      ['slide', 'added', []],
      ['text', 'changed', ['text']],
    ]);
    expect(changes[1]!.location).toMatchObject({ slideIndex: 2, text: 'world' });
  });

  test('a deletion before an edited slide reports the removal and the edit', () => {
    const changes = changesOf(
      deck([{ shapes: box(2, 'One') }, { shapes: box(2, 'Two') }, { shapes: box(2, 'Three') }]),
      deck([{ shapes: box(2, 'One') }, { shapes: box(2, 'Three!') }])
    );
    expect(changes).toHaveLength(2);
    expect(changes[0]).toMatchObject({
      computed: { scope: 'slide', action: 'removed', previousSlideIndex: 1 },
    });
    expect(changes[1]).toMatchObject({
      computed: { scope: 'text', action: 'changed' },
      location: { slideIndex: 1, shapeId: 2, text: 'Three!' },
    });
  });

  test('two slides that trade similar text lose no edit', () => {
    const changes = changesOf(
      deck([{ shapes: box(2, 'Revenue 2024') }, { shapes: box(2, 'Revenue 2025') }]),
      deck([{ shapes: box(2, 'Revenue 2025') }, { shapes: box(2, 'Revenue 2026') }])
    );
    // Equal content pairs the first new slide with the second old one.
    expect(summary(changes)).toEqual([
      ['slide', 'moved', []],
      ['text', 'changed', ['text']],
    ]);
    expect(changes[1]!.location).toMatchObject({ slideIndex: 1, text: '2026' });
  });
});

describe('changes that were silent', () => {
  test('a hidden slide toggle is a visibility change', () => {
    expect(
      changesOf(deck([{ shapes: box(2, 'A') }]), deck([{ shapes: box(2, 'A'), hidden: true }]))
    ).toEqual([
      {
        id: 'slide:id256',
        location: { slideIndex: 0 },
        computed: { scope: 'slide', action: 'changed', aspects: ['visibility'] },
      },
    ]);
  });

  test('a slide on another layout is a layout change', () => {
    expect(
      changesOf(
        deck([{ shapes: box(2, 'A'), layout: 'title' }]),
        deck([{ shapes: box(2, 'A'), layout: 'blank' }])
      )
    ).toEqual([
      {
        id: 'slide:id256',
        location: { slideIndex: 0 },
        computed: { scope: 'slide', action: 'changed', aspects: ['layout'] },
      },
    ]);
  });

  test('removed notes are a notes removal located at the slide', () => {
    expect(
      changesOf(deck([{ shapes: box(2, 'A'), notes: 'N' }]), deck([{ shapes: box(2, 'A') }]))
    ).toEqual([
      {
        id: 'notes:id256',
        location: { slideIndex: 0 },
        computed: { scope: 'notes', action: 'removed' },
      },
    ]);
  });

  test('a shape in mc:AlternateContent reports its change', () => {
    const ac = (t: string) =>
      `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="p14">${box(7, t)}</mc:Choice><mc:Fallback>${box(7, t)}</mc:Fallback></mc:AlternateContent>`;
    expect(summary(changesOf(deck([{ shapes: ac('a') }]), deck([{ shapes: ac('b') }])))).toEqual([
      ['shape', 'changed', ['text']],
    ]);
    // A difference in the fallback only is a slide change.
    const fallback = (t: string) =>
      `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="p14">${box(7, 'a')}</mc:Choice><mc:Fallback>${box(7, t)}</mc:Fallback></mc:AlternateContent>`;
    expect(
      summary(changesOf(deck([{ shapes: fallback('a') }]), deck([{ shapes: fallback('b') }])))
    ).toEqual([['slide', 'changed', ['other']]]);
  });

  test('inserted and removed paragraphs give text changes', () => {
    const inserted = changesOf(
      deck([{ shapes: multi(2, ['a', 'c']) }]),
      deck([{ shapes: multi(2, ['a', 'b', 'c']) }])
    );
    expect(inserted).toEqual([
      {
        id: 'text:id256:2:1',
        location: { slideIndex: 0, shapeId: 2, paragraphIndex: 1, start: 0, end: 1, text: 'b' },
        computed: { scope: 'text', action: 'changed', aspects: ['text'] },
      },
    ]);
    const removed = changesOf(
      deck([{ shapes: multi(2, ['a', 'b', 'c']) }]),
      deck([{ shapes: multi(2, ['a', 'c']) }])
    );
    expect(removed).toEqual([
      {
        id: 'text:id256:2:removed',
        location: { slideIndex: 0, shapeId: 2 },
        computed: { scope: 'text', action: 'changed', aspects: ['text'] },
      },
    ]);
  });
});

describe('text ranges', () => {
  const textOf = (before: string, after: string) => {
    const changes = changesOf(
      deck([{ shapes: box(2, before) }]),
      deck([{ shapes: box(2, after) }])
    );
    expect(changes).toHaveLength(1);
    return changes[0]!.location?.text;
  };

  test('a range covers whole words with the punctuation they touch', () => {
    expect(textOf('Revenue $4.0M today', 'Revenue $4.2M today')).toBe('$4.2M');
    expect(textOf('Hello there', 'Hello world')).toBe('world');
    expect(textOf('a colour', 'a color')).toBe('color');
    expect(textOf('one two three', 'one twenty three')).toBe('twenty');
  });

  test('a range never splits a surrogate pair', () => {
    expect(textOf('😀', '😁')).toBe('😁');
    expect(textOf('x 😀 y', 'x 😁 y')).toBe('😁');
    expect(textOf('x😀y', 'x😁y')).toBe('😁');
  });
});

describe('the change limit', () => {
  test('the review module ends a truncated list with a package change', () => {
    const shapes = (n: number) => Array.from({ length: n }, (_, k) => box(k + 2, `s${k}`)).join('');
    const a = pkg(deck(Array.from({ length: 11 }, () => ({ shapes: '' }))));
    const b = pkg(deck(Array.from({ length: 11 }, () => ({ shapes: shapes(1000) }))));
    const listed = reviewModule().review!.refreshChanges!(a, b);
    expect(listed).toHaveLength(MAX_CHANGES);
    expect(listed[listed.length - 1]).toEqual({
      id: 'truncated',
      computed: { scope: 'package', action: 'changed' },
    });
  });
});
