// The commit checks of the parts a transaction touches (`store/part-checks.ts`). Each rule has
// a bad transaction that the store committed before the checks (plain ops, so nothing above
// the store stopped it) and a good one that still commits. A refusal commits nothing: no
// revision, no history entry, no change event, and `dryRun` reports the same refusal.

import { describe, expect, test } from 'bun:test';
import { buildPptx, cell, shape, table } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  children,
  CONTENT_TYPES_PART,
  elements,
  find,
  NS,
  PptxDocumentStore,
  readPptx,
  relsPartName,
  writePptx,
  xmlPartOf,
  type Element,
  type OpRejection,
  type PptxOp,
  type StoreChange,
} from '../index.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const SLIDE_RELS = '/ppt/slides/_rels/slide1.xml.rels';
const PRES = '/ppt/presentation.xml';
const PRES_RELS = relsPartName(PRES);

const row = (...cells: string[]): string => `<a:tr h="370840">${cells.join('')}</a:tr>`;

function tableDeck(rows: string, columns = [1_000_000, 1_000_000]): Uint8Array {
  return buildPptx({
    slides: [
      {
        shapes:
          shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100 }) +
          table({
            id: 3,
            x: 0,
            y: 1_000_000,
            cx: 2_000_000,
            cy: 741_680,
            columns,
            tblPr: '<a:tblPr/>',
            rows,
          }),
      },
      { shapes: shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100 }) },
    ],
  });
}

const PLAIN = row(cell('a'), cell('b')) + row(cell('c'), cell('d'));

function open(bytes: Uint8Array = tableDeck(PLAIN)): PptxDocumentStore {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('fixture');
  return new PptxDocumentStore(r.pkg);
}

const root = (s: PptxDocumentStore, part: string): Element => xmlPartOf(s.package, part)!.root;
const cellsOf = (s: PptxDocumentStore): Element[][] =>
  children(find(root(s, SLIDE), 'a:tbl'), 'a:tr').map((tr) => children(tr, 'a:tc'));
const set = (part: string, node: string, local: string, value: string | null): PptxOp => ({
  op: 'setAttribute',
  part,
  node,
  name: { ns: '', local },
  value,
});

/** Run a transaction that must be refused, and prove that nothing was committed. */
function refused(s: PptxDocumentStore, ops: PptxOp[], origin?: 'remote'): OpRejection {
  const revision = s.revision;
  const pkg = s.package;
  const canUndo = s.canUndo;
  const events: StoreChange[] = [];
  const off = s.subscribe((c) => events.push(c));
  const dry = s.dryRun(ops);
  const r = s.transact(ops, origin ? { origin } : {});
  off();
  expect(r.ok).toBe(false);
  expect(dry).toEqual(r);
  expect(s.revision).toBe(revision);
  expect(s.package).toBe(pkg);
  expect(s.canUndo).toBe(canUndo);
  expect(events).toEqual([]);
  if (r.ok) throw new Error('committed');
  return r.rejection;
}

function committed(s: PptxDocumentStore, ops: PptxOp[]): void {
  expect(s.dryRun(ops).ok).toBe(true);
  const r = s.transact(ops);
  if (!r.ok) throw new Error(JSON.stringify(r.rejection));
  // The committed package saves and reopens.
  expect(readPptx(writePptx(s.package)).ok).toBe(true);
}

describe('table-grid: one a:tc per a:gridCol in every a:tr', () => {
  test('removing one cell of a row is refused', () => {
    const s = open();
    const rejection = refused(s, [{ op: 'removeNode', part: SLIDE, node: cellsOf(s)[0][1].id }]);
    expect(rejection).toMatchObject({ code: 'tree-invariant', rule: 'table-grid', part: SLIDE });
  });

  test('removing a column (grid column and its cells) commits', () => {
    const s = open();
    const grid = elements(find(root(s, SLIDE), 'a:tblGrid'));
    const cells = cellsOf(s);
    committed(s, [
      { op: 'removeNode', part: SLIDE, node: grid[1].id },
      { op: 'removeNode', part: SLIDE, node: cells[0][1].id },
      { op: 'removeNode', part: SLIDE, node: cells[1][1].id },
    ]);
  });
});

describe('table-merge: spans and merge markers agree', () => {
  test('a gridSpan whose covered cell has no hMerge is refused', () => {
    const s = open();
    const rejection = refused(s, [set(SLIDE, cellsOf(s)[0][0].id, 'gridSpan', '2')]);
    expect(rejection).toMatchObject({ code: 'tree-invariant', rule: 'table-merge' });
  });

  test('a span past the grid is refused', () => {
    const s = open();
    const c = cellsOf(s);
    const rejection = refused(s, [
      set(SLIDE, c[1][0].id, 'rowSpan', '2'),
      set(SLIDE, c[1][1].id, 'rowSpan', '2'),
    ]);
    expect(rejection).toMatchObject({ code: 'tree-invariant', rule: 'table-merge' });
  });

  test('a merge marker outside a block is refused', () => {
    const s = open();
    const rejection = refused(s, [set(SLIDE, cellsOf(s)[1][1].id, 'vMerge', '1')]);
    expect(rejection).toMatchObject({ code: 'tree-invariant', rule: 'table-merge' });
  });

  test('a span into another merged block is refused', () => {
    const s = open(
      tableDeck(
        row(cell('a'), cell('b', { attrs: ' rowSpan="2"' })) +
          row(cell('c'), cell('', { attrs: ' vMerge="1"' }))
      )
    );
    const rejection = refused(s, [set(SLIDE, cellsOf(s)[1][0].id, 'gridSpan', '2')]);
    expect(rejection).toMatchObject({ code: 'tree-invariant', rule: 'table-merge' });
  });

  test('an orphan marker the deck had does not block an edit, but a new orphan is refused', () => {
    const s = open(tableDeck(row(cell('a'), cell('b', { attrs: ' vMerge="1"' })) + PLAIN));
    const c = cellsOf(s);
    committed(s, [set(SLIDE, c[0][0].id, 'marL', '0')]);
    const rejection = refused(s, [set(SLIDE, c[1][1].id, 'hMerge', '1')]);
    expect(rejection).toMatchObject({ code: 'tree-invariant', rule: 'table-merge' });
    committed(s, [set(SLIDE, c[0][1].id, 'vMerge', null)]);
  });

  test('a 2 x 2 merge with every marker commits, and so does its split', () => {
    const s = open();
    const c = cellsOf(s);
    const merge = [
      set(SLIDE, c[0][0].id, 'gridSpan', '2'),
      set(SLIDE, c[0][0].id, 'rowSpan', '2'),
      set(SLIDE, c[0][1].id, 'rowSpan', '2'),
      set(SLIDE, c[0][1].id, 'hMerge', '1'),
      set(SLIDE, c[1][0].id, 'gridSpan', '2'),
      set(SLIDE, c[1][0].id, 'vMerge', '1'),
      set(SLIDE, c[1][1].id, 'hMerge', '1'),
      set(SLIDE, c[1][1].id, 'vMerge', '1'),
    ];
    committed(s, merge);
    committed(
      s,
      merge.map((op) => ({ ...op, value: null }))
    );
  });
});

describe('shape-id: p:cNvPr/@id is unique in a slide, layout, or master', () => {
  const nvSpec = (id: string) => ({
    ns: NS.p,
    local: 'sp',
    children: [
      {
        ns: NS.p,
        local: 'nvSpPr',
        children: [
          {
            ns: NS.p,
            local: 'cNvPr',
            attrs: [
              { ns: '', local: 'id', value: id },
              { ns: '', local: 'name', value: `Shape ${id}` },
            ],
          },
          { ns: NS.p, local: 'cNvSpPr' },
          { ns: NS.p, local: 'nvPr' },
        ],
      },
      { ns: NS.p, local: 'spPr' },
    ],
  });
  const tree = (s: PptxDocumentStore, part = SLIDE): string => find(root(s, part), 'p:spTree')!.id;

  test('a shape with an id the slide already uses is refused', () => {
    const s = open();
    const insert: PptxOp = {
      op: 'insertNode',
      part: SLIDE,
      parent: tree(s),
      index: 2,
      spec: nvSpec('2'),
    };
    expect(refused(s, [insert])).toMatchObject({
      code: 'tree-invariant',
      rule: 'shape-id',
      part: SLIDE,
    });
  });

  test('a fresh id commits, and another slide may reuse the id', () => {
    const s = open();
    committed(s, [{ op: 'insertNode', part: SLIDE, parent: tree(s), index: 2, spec: nvSpec('9') }]);
    const other = '/ppt/slides/slide2.xml';
    committed(s, [
      { op: 'insertNode', part: other, parent: tree(s, other), index: 2, spec: nvSpec('9') },
    ]);
  });

  test('an mc:Fallback that repeats the shapes of its mc:Choice is not a duplicate', () => {
    const s = open();
    const alt = {
      ns: NS.mc,
      local: 'AlternateContent',
      nsDecls: [{ prefix: 'mc', uri: NS.mc }],
      prefix: 'mc',
      children: [
        {
          ns: NS.mc,
          local: 'Choice',
          prefix: 'mc',
          attrs: [{ ns: '', local: 'Requires', value: 'p14' }],
          children: [nvSpec('7')],
        },
        { ns: NS.mc, local: 'Fallback', prefix: 'mc', children: [nvSpec('7')] },
      ],
    };
    committed(s, [{ op: 'insertNode', part: SLIDE, parent: tree(s), index: 2, spec: alt }]);
  });

  test('a duplicate the deck already had does not block an edit', () => {
    const s = open(
      buildPptx({
        slides: [
          {
            shapes:
              shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100 }) +
              shape({ id: 2, x: 200, y: 0, cx: 100, cy: 100 }),
          },
        ],
      })
    );
    const off = find(root(s, SLIDE), 'a:off')!;
    committed(s, [set(SLIDE, off.id, 'x', '50')]);
    // A new duplicate is still refused: the shape tree itself has id 1.
    expect(
      refused(s, [{ op: 'insertNode', part: SLIDE, parent: tree(s), index: 2, spec: nvSpec('1') }])
    ).toMatchObject({ rule: 'shape-id' });
  });
});

describe('slide-list: every id list entry names a relationship to a part', () => {
  const relById = (s: PptxDocumentStore, part: string, id: string): Element =>
    elements(root(s, part)).find((r) => attr(r, 'Id') === id)!;

  test('removing the relationship of a listed slide is refused', () => {
    const s = open();
    const rejection = refused(s, [
      { op: 'removeNode', part: PRES_RELS, node: relById(s, PRES_RELS, 'rId101').id },
    ]);
    expect(rejection).toMatchObject({ code: 'package-invariant', rule: 'slide-list', part: PRES });
  });

  test('removing a listed slide part is refused', () => {
    const s = open();
    expect(refused(s, [{ op: 'removePart', part: SLIDE }])).toMatchObject({
      code: 'package-invariant',
      rule: 'slide-list',
    });
  });

  test('an entry that names a missing r:id is refused, for masters too', () => {
    const s = open();
    const master = find(root(s, PRES), 'p:sldMasterId')!;
    const rejection = refused(s, [
      {
        op: 'setAttribute',
        part: PRES,
        node: master.id,
        name: { ns: NS.r, local: 'id', prefix: 'r' },
        value: 'rId999',
      },
    ]);
    expect(rejection).toMatchObject({ code: 'package-invariant', rule: 'slide-list' });
  });

  test('deleting a slide whole (entry, relationship, parts, content type) commits', () => {
    const s = open();
    const entry = elements(find(root(s, PRES), 'p:sldIdLst')).find(
      (e) => attr(e, 'r:id') === 'rId102'
    )!;
    const override = elements(root(s, CONTENT_TYPES_PART)).find(
      (e) => attr(e, 'PartName') === '/ppt/slides/slide2.xml'
    )!;
    committed(s, [
      { op: 'removeNode', part: PRES, node: entry.id },
      { op: 'removeNode', part: PRES_RELS, node: relById(s, PRES_RELS, 'rId102').id },
      { op: 'removeNode', part: CONTENT_TYPES_PART, node: override.id },
      { op: 'removePart', part: '/ppt/slides/slide2.xml' },
      { op: 'removePart', part: '/ppt/slides/_rels/slide2.xml.rels' },
    ]);
  });
});

describe('relationship-target: a touched .rels part targets parts of the package', () => {
  const layoutRel = (s: PptxDocumentStore): Element => elements(root(s, SLIDE_RELS))[0];

  test('a relationship to a missing part is refused', () => {
    const s = open();
    const rejection = refused(s, [
      set(SLIDE_RELS, layoutRel(s).id, 'Target', '../slideLayouts/slideLayout99.xml'),
    ]);
    expect(rejection).toMatchObject({
      code: 'package-invariant',
      rule: 'relationship-target',
      part: SLIDE_RELS,
    });
  });

  test('a relationship to another existing part, or an external one, commits', () => {
    const s = open();
    committed(s, [set(SLIDE_RELS, layoutRel(s).id, 'Target', '../slideLayouts/slideLayout2.xml')]);
    committed(s, [
      {
        op: 'insertNode',
        part: SLIDE_RELS,
        parent: root(s, SLIDE_RELS).id,
        index: 1,
        spec: {
          ns: NS.rel,
          local: 'Relationship',
          attrs: [
            { ns: '', local: 'Id', value: 'rId50' },
            {
              ns: '',
              local: 'Type',
              value:
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
            },
            { ns: '', local: 'Target', value: 'https://example.com/missing' },
            { ns: '', local: 'TargetMode', value: 'External' },
          ],
        },
      },
    ]);
  });
});

describe('content-type: every part has a content type', () => {
  test('a part with no Override and no Default is refused', () => {
    const s = open();
    const rejection = refused(s, [{ op: 'addPart', part: '/ppt/media/a.bin', base64: 'AAEC' }]);
    expect(rejection).toEqual({
      code: 'package-invariant',
      rule: 'content-type',
      part: '/ppt/media/a.bin',
      detail:
        '/ppt/media/a.bin has no content type (no Override for it, no Default for its extension)',
    });
  });

  test('removing the Default that types the relationship parts is refused', () => {
    const s = open();
    const rels = elements(root(s, CONTENT_TYPES_PART)).find(
      (e) => attr(e, 'Extension') === 'rels'
    )!;
    const rejection = refused(s, [{ op: 'removeNode', part: CONTENT_TYPES_PART, node: rels.id }]);
    expect(rejection).toMatchObject({ code: 'package-invariant', rule: 'content-type' });
    expect(rejection.code === 'package-invariant' && rejection.part.endsWith('.rels')).toBe(true);
  });

  test('a part added with an Override commits', () => {
    const s = open();
    committed(s, [
      {
        op: 'insertNode',
        part: CONTENT_TYPES_PART,
        parent: root(s, CONTENT_TYPES_PART).id,
        index: 0,
        spec: {
          ns: NS.ct,
          local: 'Override',
          attrs: [
            { ns: '', local: 'PartName', value: '/ppt/media/a.bin' },
            { ns: '', local: 'ContentType', value: 'application/octet-stream' },
          ],
        },
      },
      { op: 'addPart', part: '/ppt/media/a.bin', base64: 'AAEC' },
    ]);
  });
});

describe('remote transactions', () => {
  test('a remote transaction is checked like a local one, and nothing commits', () => {
    const s = open();
    const node = cellsOf(s)[0][1].id;
    expect(refused(s, [{ op: 'removeNode', part: SLIDE, node }], 'remote')).toMatchObject({
      code: 'tree-invariant',
      rule: 'table-grid',
    });
    expect(refused(s, [{ op: 'removePart', part: SLIDE }], 'remote')).toMatchObject({
      rule: 'slide-list',
    });
  });
});

describe('cost follows the edit', () => {
  test('an edit of one slide does not read another slide', () => {
    // A broken table on slide 2 does not refuse an edit of slide 1.
    const s = open(
      buildPptx({
        slides: [
          { shapes: shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100 }) },
          {
            shapes: table({
              id: 2,
              x: 0,
              y: 0,
              cx: 2_000_000,
              cy: 370_840,
              columns: [1_000_000, 1_000_000],
              tblPr: '<a:tblPr/>',
              rows: row(cell('only one')),
            }),
          },
        ],
      })
    );
    committed(s, [set(SLIDE, find(root(s, SLIDE), 'a:off')!.id, 'x', '7')]);
  });
});

const binOverride = (s: PptxDocumentStore, name: string): PptxOp => ({
  op: 'insertNode',
  part: CONTENT_TYPES_PART,
  parent: root(s, CONTENT_TYPES_PART).id,
  index: 0,
  spec: {
    ns: NS.ct,
    local: 'Override',
    attrs: [
      { ns: '', local: 'PartName', value: name },
      { ns: '', local: 'ContentType', value: 'application/octet-stream' },
    ],
  },
});

const imageRel = (s: PptxDocumentStore, id: string, target: string): PptxOp => ({
  op: 'insertNode',
  part: SLIDE_RELS,
  parent: root(s, SLIDE_RELS).id,
  index: 0,
  spec: {
    ns: NS.rel,
    local: 'Relationship',
    attrs: [
      { ns: '', local: 'Id', value: id },
      {
        ns: '',
        local: 'Type',
        value: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      },
      { ns: '', local: 'Target', value: target },
    ],
  },
});

/** A store whose slide 1 has an image relationship `rId60` to the part `/ppt/media/a.bin`. */
function withMedia(): PptxDocumentStore {
  const s = open();
  committed(s, [
    binOverride(s, '/ppt/media/a.bin'),
    { op: 'addPart', part: '/ppt/media/a.bin', base64: 'AAEC' },
    imageRel(s, 'rId60', '../media/a.bin'),
  ]);
  return s;
}

describe('relationship-target: a removed part leaves no relationship behind', () => {
  test('removing a part that an untouched .rels part still targets is refused', () => {
    const s = withMedia();
    // The slide's relationships part is not touched: without the check this committed, and the
    // saved deck had a relationship to a missing part.
    expect(refused(s, [{ op: 'removePart', part: '/ppt/media/a.bin' }])).toEqual({
      code: 'package-invariant',
      rule: 'relationship-target',
      part: SLIDE_RELS,
      detail: 'relationship rId60 targets /ppt/media/a.bin, which the transaction removed',
    });
    // A slide layout that a slide and the master still use cannot go either.
    expect(
      refused(s, [{ op: 'removePart', part: '/ppt/slideLayouts/slideLayout1.xml' }])
    ).toMatchObject({ code: 'package-invariant', rule: 'relationship-target' });
  });

  test('removing the part with its relationship commits', () => {
    const s = withMedia();
    const rel = elements(root(s, SLIDE_RELS)).find((r) => attr(r, 'Id') === 'rId60')!;
    committed(s, [
      { op: 'removeNode', part: SLIDE_RELS, node: rel.id },
      { op: 'removePart', part: '/ppt/media/a.bin' },
    ]);
  });
});

describe('content-type: an edit of [Content_Types].xml reads only what it removed', () => {
  test('a part that arrived untyped does not block an edit of the content types', () => {
    const r = readPptx(tableDeck(PLAIN));
    if (!r.ok) throw new Error('fixture');
    const parts = new Map(r.pkg.parts);
    const old = '/ppt/media/old.bin';
    parts.set(old, { kind: 'binary', partName: old, bytes: new Uint8Array([1]) });
    const s = new PptxDocumentStore({ parts });
    const ops: PptxOp[] = [
      binOverride(s, '/ppt/media/new.bin'),
      { op: 'addPart', part: '/ppt/media/new.bin', base64: 'AAEC' },
    ];
    expect(s.transact(ops).ok).toBe(true);
  });

  test('removing an Override is refused for the part it typed', () => {
    const s = withMedia();
    const over = elements(root(s, CONTENT_TYPES_PART)).find(
      (e) => attr(e, 'PartName') === '/ppt/media/a.bin'
    )!;
    expect(
      refused(s, [{ op: 'removeNode', part: CONTENT_TYPES_PART, node: over.id }])
    ).toMatchObject({ code: 'package-invariant', rule: 'content-type', part: '/ppt/media/a.bin' });
  });

  test('removing a Default that no part uses commits', () => {
    const s = open();
    committed(s, [
      {
        op: 'insertNode',
        part: CONTENT_TYPES_PART,
        parent: root(s, CONTENT_TYPES_PART).id,
        index: 0,
        spec: {
          ns: NS.ct,
          local: 'Default',
          attrs: [
            { ns: '', local: 'Extension', value: 'gif' },
            { ns: '', local: 'ContentType', value: 'image/gif' },
          ],
        },
      },
    ]);
    const gif = elements(root(s, CONTENT_TYPES_PART)).find((e) => attr(e, 'Extension') === 'gif')!;
    committed(s, [{ op: 'removeNode', part: CONTENT_TYPES_PART, node: gif.id }]);
  });
});

describe('undo and redo are never refused', () => {
  test('undo restores a violation the deck had; redo and later edits still apply', () => {
    const s = open(
      buildPptx({
        slides: [
          {
            shapes:
              shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100 }) +
              shape({ id: 2, x: 200, y: 0, cx: 100, cy: 100 }),
          },
        ],
      })
    );
    const idOf = (): string | undefined => {
      const sps = elements(find(root(s, SLIDE), 'p:spTree')).filter((e) => e.local === 'sp');
      return attr(find(sps[1], 'p:cNvPr')!, 'id');
    };
    const sps = elements(find(root(s, SLIDE), 'p:spTree')).filter((e) => e.local === 'sp');
    // Fix the duplicate, then undo the fix: the restored revision has the duplicate again.
    committed(s, [set(SLIDE, find(sps[1], 'p:cNvPr')!.id, 'id', '3')]);
    expect(s.undo().ok).toBe(true);
    expect(idOf()).toBe('2');
    expect(s.redo().ok).toBe(true);
    expect(idOf()).toBe('3');
    expect(s.undo().ok).toBe(true);
    // An edit of the restored revision that keeps the old duplicate commits.
    committed(s, [set(SLIDE, find(root(s, SLIDE), 'a:off')!.id, 'x', '9')]);
  });
});
