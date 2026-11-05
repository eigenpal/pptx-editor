// Table cells through the automation host: cell text bodies, anchors, splits, resizes, clears.

import { describe, expect, test } from 'bun:test';
import { buildPptx, table } from '../../../../../scripts/fixtures/builder.ts';
import {
  elements,
  PptxDocumentStore,
  readPptx,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import { isStructureEdited } from '../commands/table-edit-marks.ts';
import {
  createServerAutomationHost,
  createStoreAutomationHost,
  type AutomationHandle,
  type AutomationHost,
  type AutomationOperation,
  type AutomationValue,
} from '../index.ts';

const REL_HYPERLINK =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

function open(bytes = buildPptx({ slides: [{ layout: 'blank', shapes: '' }] })): AutomationHost {
  const opened = createServerAutomationHost(bytes);
  if (!opened.ok) throw new Error(JSON.stringify(opened.error));
  return opened.host;
}

function ok(host: AutomationHost, operations: AutomationOperation[]): AutomationValue[] {
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results.map((r) => (r.status === 'ok' ? r.value : { kind: 'null' as const }));
}

function refused(host: AutomationHost, operations: AutomationOperation[]): string | undefined {
  const response = host.execute({ operations });
  expect(response.ok).toBe(false);
  const failed = response.results.find((r) => r.status === 'error');
  return failed?.status === 'error' ? failed.error.code : undefined;
}

function handle(value: AutomationValue | undefined): AutomationHandle {
  if (value?.kind !== 'handle') throw new Error(`not a handle: ${JSON.stringify(value)}`);
  return value.handle;
}

/** A host with one `rows` by `columns` table on its slide. */
function withTable(rows = 3, columns = 3): { host: AutomationHost; t: AutomationHandle } {
  const host = open();
  const [, table] = ok(host, [
    { op: 'getSlide', key: { index: 0 } },
    { op: 'addShape', container: { result: 0 }, shape: { kind: 'table', rows, columns } },
  ]);
  return { host, t: handle(table) };
}

function reopen(host: AutomationHost): AutomationHost {
  const saved = host.save();
  if (!saved.ok) throw new Error(saved.error.code);
  expect(readPptx(saved.bytes).ok).toBe(true);
  return open(saved.bytes);
}

/** The table of the first slide of a reopened host. */
function tableOf(host: AutomationHost): AutomationHandle {
  const [, shape] = ok(host, [
    { op: 'getSlide', key: { index: 0 } },
    { op: 'getShape', container: { result: 0 }, key: { index: 0 } },
  ]);
  return handle(shape);
}

describe('automation host: table cell text bodies', () => {
  test('text, font, and paragraph operations address a cell', () => {
    const { host, t } = withTable();
    const cell = { row: 1, column: 2 };
    ok(host, [
      { op: 'setText', shape: t, text: 'Total: 42', cell },
      { op: 'setFont', shape: t, span: { start: 7, length: 2 }, font: { bold: true }, cell },
      { op: 'setParagraph', shape: t, paragraph: { alignment: 'r', indentLevel: 2 }, cell },
    ]);
    const again = reopen(host);
    const u = tableOf(again);
    const [text, bold, plain, paragraph, value] = ok(again, [
      { op: 'readText', shape: u, cell },
      { op: 'readFont', shape: u, span: { start: 7, length: 2 }, cell },
      { op: 'readFont', shape: u, span: { start: 0, length: 5 }, cell },
      { op: 'readParagraph', shape: u, cell },
      { op: 'readTable', shape: u },
    ]);
    expect(text).toEqual({ kind: 'text', text: 'Total: 42' });
    expect(bold).toMatchObject({ font: { bold: true } });
    expect(plain).toMatchObject({ font: { bold: false } });
    expect(paragraph).toMatchObject({ paragraph: { alignment: 'r', indentLevel: 2 } });
    expect(value).toMatchObject({
      table: {
        values: [
          ['', '', ''],
          ['', '', 'Total: 42'],
          ['', '', ''],
        ],
      },
    });
  });

  test('a covered position addresses the anchor of its merged block', () => {
    const { host, t } = withTable();
    ok(host, [
      { op: 'mergeTableCells', shape: t, row: 0, column: 0, rowCount: 2, columnCount: 2 },
      { op: 'setText', shape: t, text: 'Merged', cell: { row: 1, column: 1 } },
    ]);
    expect(ok(host, [{ op: 'readText', shape: t, cell: { row: 0, column: 0 } }])[0]).toEqual({
      kind: 'text',
      text: 'Merged',
    });
  });

  test('a cell outside the grid, or a cell of a shape that is not a table, refuses', () => {
    const { host, t } = withTable(2, 2);
    expect(refused(host, [{ op: 'readText', shape: t, cell: { row: 2, column: 0 } }])).toBe(
      'invalid-argument'
    );
    const [, box] = ok(host, [
      { op: 'getSlide', key: { index: 0 } },
      {
        op: 'addShape',
        container: { result: 0 },
        shape: { kind: 'textBox', text: 'x' },
      },
    ]);
    expect(
      refused(host, [{ op: 'readText', shape: handle(box), cell: { row: 0, column: 0 } }])
    ).toBe('unsupported-content');
  });
});

describe('automation host: table cell structure and format', () => {
  test('the vertical anchor of a cell reads back, with its centering', () => {
    const { host, t } = withTable();
    expect(ok(host, [{ op: 'readTableCell', shape: t, row: 0, column: 0 }])[0]).toMatchObject({
      cell: { anchor: 't', anchorCenter: false },
    });
    ok(host, [
      {
        op: 'setTableCellAnchor',
        shape: t,
        row: 0,
        column: 0,
        anchor: 'ctr',
        anchorCenter: true,
      },
    ]);
    const again = reopen(host);
    expect(
      ok(again, [{ op: 'readTableCell', shape: tableOf(again), row: 0, column: 0 }])[0]
    ).toMatchObject({ cell: { anchor: 'ctr', anchorCenter: true } });
    expect(
      refused(host, [
        { op: 'setTableCellAnchor', shape: t, row: 0, column: 0, anchor: 'x', anchorCenter: false },
      ])
    ).toBe('invalid-argument');
  });

  test('a split adds grid lines for a plain cell and follows grid lines for a merged one', () => {
    const { host, t } = withTable(2, 2);
    ok(host, [{ op: 'splitTableCell', shape: t, row: 0, column: 0, rowCount: 2, columnCount: 2 }]);
    expect(ok(host, [{ op: 'readTable', shape: t }])[0]).toMatchObject({
      table: { rowCount: 3, columnCount: 3 },
    });
    ok(host, [
      { op: 'mergeTableCells', shape: t, row: 0, column: 0, rowCount: 2, columnCount: 2 },
      { op: 'splitTableCell', shape: t, row: 0, column: 0, rowCount: 1, columnCount: 2 },
    ]);
    reopen(host);
    const [table, left] = ok(host, [
      { op: 'readTable', shape: t },
      { op: 'readTableCell', shape: t, row: 1, column: 0 },
    ]);
    expect(table).toMatchObject({ table: { rowCount: 3, columnCount: 3 } });
    expect(left).toMatchObject({
      cell: { rowIndex: 0, columnIndex: 0, rowCount: 2, columnCount: 1 },
    });
  });

  test('a resize grows and shrinks the span of a cell', () => {
    const { host, t } = withTable();
    ok(host, [{ op: 'resizeTableCell', shape: t, row: 0, column: 0, rowCount: 2, columnCount: 3 }]);
    expect(ok(host, [{ op: 'readTableCell', shape: t, row: 1, column: 2 }])[0]).toMatchObject({
      cell: { rowIndex: 0, columnIndex: 0, rowCount: 2, columnCount: 3 },
    });
    ok(host, [{ op: 'resizeTableCell', shape: t, row: 0, column: 0, rowCount: 1, columnCount: 2 }]);
    const again = reopen(host);
    const u = tableOf(again);
    const [anchor, freed] = ok(again, [
      { op: 'readTableCell', shape: u, row: 0, column: 1 },
      { op: 'readTableCell', shape: u, row: 1, column: 0 },
    ]);
    expect(anchor).toMatchObject({ cell: { rowIndex: 0, columnIndex: 0, columnCount: 2 } });
    expect(freed).toMatchObject({ cell: { rowIndex: 1, columnIndex: 0, rowCount: 1 } });
  });

  test('a resize that cuts another merged block, or leaves the grid, refuses', () => {
    const { host, t } = withTable();
    ok(host, [{ op: 'mergeTableCells', shape: t, row: 1, column: 1, rowCount: 2, columnCount: 2 }]);
    expect(
      refused(host, [
        { op: 'resizeTableCell', shape: t, row: 0, column: 0, rowCount: 2, columnCount: 2 },
      ])
    ).toBe('unsupported-content');
    expect(
      refused(host, [
        { op: 'resizeTableCell', shape: t, row: 0, column: 0, rowCount: 4, columnCount: 1 },
      ])
    ).toBe('invalid-argument');
  });

  test('a cell reads its runs with their formats, joined like its text', () => {
    const { host, t } = withTable();
    // Row 0 is the header row, which the default table style makes bold.
    const cell = { row: 1, column: 1 };
    ok(host, [
      { op: 'setText', shape: t, text: 'Big bold\rplain', cell },
      { op: 'setFont', shape: t, span: { start: 4, length: 4 }, font: { bold: true }, cell },
      {
        op: 'setFont',
        shape: t,
        span: { start: 0, length: 3 },
        font: { size: 28, color: '#C00000', underline: 'sng' },
        cell,
      },
    ]);
    const again = reopen(host);
    const [runs, text, empty] = ok(again, [
      { op: 'readTextRuns', shape: tableOf(again), cell },
      { op: 'readText', shape: tableOf(again), cell },
      { op: 'readTextRuns', shape: tableOf(again), cell: { row: 2, column: 2 } },
    ]);
    if (runs?.kind !== 'textRuns') throw new Error('expected runs');
    expect(runs.runs.map((run) => run.text)).toEqual(['Big', ' ', 'bold\r', 'plain']);
    expect(runs.runs.map((run) => run.text).join('')).toBe((text as { text: string }).text);
    expect(runs.runs[0]!.font).toMatchObject({
      size: 28,
      color: '#C00000',
      underline: 'sng',
      bold: false,
    });
    expect(runs.runs[2]!.font).toMatchObject({ bold: true, underline: 'none' });
    expect(runs.runs[3]!.font).toMatchObject({ bold: false });
    expect(empty).toEqual({ kind: 'textRuns', runs: [] });
  });

  test('a clear removes text, direct format, or both', () => {
    const { host, t } = withTable(2, 2);
    const cell = { row: 0, column: 0 };
    const formatted: AutomationOperation[] = [
      { op: 'setText', shape: t, text: 'Keep', cell },
      { op: 'setFont', shape: t, font: { italic: true }, cell },
      { op: 'setTableCellAnchor', shape: t, ...cell, anchor: 'b', anchorCenter: false },
    ];
    ok(host, [...formatted, { op: 'clearTable', shape: t, text: false, format: true }]);
    const [text, font, anchor] = ok(host, [
      { op: 'readText', shape: t, cell },
      { op: 'readFont', shape: t, cell },
      { op: 'readTableCell', shape: t, ...cell },
    ]);
    expect(text).toEqual({ kind: 'text', text: 'Keep' });
    expect(font).toMatchObject({ font: { italic: false } });
    expect(anchor).toMatchObject({ cell: { anchor: 't' } });
    ok(host, [{ op: 'clearTable', shape: t, text: true, format: false }]);
    reopen(host);
    expect(ok(host, [{ op: 'readText', shape: t, cell }])[0]).toEqual({ kind: 'text', text: '' });
  });
});

describe('automation host: table format clear', () => {
  test('a format clear keeps the hyperlink of a run and removes its format', () => {
    const link =
      '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" b="1">' +
      '<a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:hlinkClick r:id="rId9"/></a:rPr>' +
      '<a:t>Site</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>';
    const host = open(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            rels: [['rId9', REL_HYPERLINK, 'https://example.com/', 'External']],
            shapes: table({
              id: 2,
              x: 0,
              y: 0,
              cx: 914_400,
              cy: 370_840,
              columns: [914_400],
              tblPr: '<a:tblPr/>',
              rows: `<a:tr h="370840">${link}</a:tr>`,
            }),
          },
        ],
      })
    );
    const t = tableOf(host);
    ok(host, [{ op: 'clearTable', shape: t, text: false, format: true }]);
    const [font] = ok(host, [{ op: 'readFont', shape: t, cell: { row: 0, column: 0 } }]);
    expect(font).toMatchObject({ font: { bold: false } });
    const saved = host.save();
    if (!saved.ok) throw new Error(saved.error.code);
    const read = readPptx(saved.bytes);
    if (!read.ok) throw new Error('the saved file does not read back');
    const slide = xmlPartOf(read.pkg, '/ppt/slides/slide1.xml')!.root;
    const named = (el: Element, local: string): Element[] =>
      elements(el).flatMap((c) => (c.local === local ? [c] : named(c, local)));
    const [rPr] = named(slide, 'rPr');
    expect(rPr?.attrs).toEqual([]);
    expect(elements(rPr).map((e) => e.local)).toEqual(['hlinkClick']);
  });
});

describe('automation host: table cell fill', () => {
  test('a cell fill writes the block, reads back, and clears', () => {
    const { host, t } = withTable();
    const cell = { row: 1, column: 1 };
    ok(host, [
      { op: 'mergeTableCells', shape: t, row: 1, column: 1, rowCount: 2, columnCount: 1 },
      { op: 'setFill', shape: t, fill: { kind: 'solid', color: '#C00000' }, cell },
      { op: 'setFill', shape: t, fill: { kind: 'transparency', transparency: 0.5 }, cell },
    ]);
    const again = reopen(host);
    const u = tableOf(again);
    expect(ok(again, [{ op: 'readFill', shape: u, cell: { row: 2, column: 1 } }])[0]).toEqual({
      kind: 'fill',
      fill: { type: 'solid', foregroundColor: '#C00000', transparency: 0.5 },
    });
    ok(again, [{ op: 'setFill', shape: u, fill: { kind: 'clear' }, cell }]);
    expect(ok(again, [{ op: 'readFill', shape: u, cell }])[0]).toMatchObject({
      fill: { type: 'noFill' },
    });
  });
});

describe('automation host: table structure edits in a collaboration room', () => {
  /** Merge two filled cells through a host over a store, and read the table frame. */
  function frameAfterMerge(collaborating: boolean): Element {
    const read = readPptx(buildPptx({ slides: [{ layout: 'blank', shapes: '' }] }));
    if (!read.ok) throw new Error('the deck does not read');
    const store = new PptxDocumentStore(read.pkg);
    const host = createStoreAutomationHost(store, { collaborating });
    const [, frame] = ok(host, [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'addShape', container: { result: 0 }, shape: { kind: 'table', rows: 1, columns: 2 } },
    ]);
    const t = handle(frame);
    ok(host, [
      { op: 'setTableCell', shape: t, row: 0, column: 0, text: 'A' },
      { op: 'setTableCell', shape: t, row: 0, column: 1, text: 'B' },
      { op: 'mergeTableCells', shape: t, row: 0, column: 0, rowCount: 1, columnCount: 2 },
    ]);
    let found: Element | undefined;
    const visit = (el: Element): void => {
      if (el.kind === 'graphicFrame') found = el;
      for (const c of elements(el)) visit(c);
    };
    visit(xmlPartOf(store.package, '/ppt/slides/slide1.xml')!.root);
    if (!found) throw new Error('no table frame');
    return found;
  }

  test('a merge marks the table structure only in a room', () => {
    expect(isStructureEdited(frameAfterMerge(true))).toBe(true);
    expect(isStructureEdited(frameAfterMerge(false))).toBe(false);
  });
});

// A cell may hold two text bodies: a file can have one, and so can a collaboration room where
// two peers typed at once into a cell with no text body. The save joins them (`cellAsSaved`).
describe('automation host: a cell with two text bodies', () => {
  const body = (text: string) =>
    `<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r></a:p></a:txBody>`;
  const twoBodies = () =>
    open(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: table({
              id: 2,
              x: 914_400,
              y: 914_400,
              cx: 1_828_800,
              cy: 370_840,
              columns: [914_400, 914_400],
              tblPr: '<a:tblPr/>',
              rows: `<a:tr h="370840"><a:tc>${body('A')}<a:tcPr/></a:tc><a:tc>${body('xx')}${body('yy')}<a:tcPr/></a:tc></a:tr>`,
            }),
          },
        ],
      })
    );
  const cell = { row: 0, column: 1 };

  test('a read gives the text of both bodies and writes nothing', () => {
    const host = twoBodies();
    const t = tableOf(host);
    const response = host.execute({
      operations: [
        { op: 'readTable', shape: t },
        { op: 'readTableCell', shape: t, ...cell },
        { op: 'readText', shape: t, cell },
      ],
    });
    expect(response.changed).toBe(false);
    const [read, one, text] = response.results.map((r) => (r.status === 'ok' ? r.value : null));
    expect(read).toMatchObject({ table: { values: [['A', 'xx\ryy']] } });
    expect(one).toMatchObject({ cell: { text: 'xx\ryy' } });
    expect(text).toMatchObject({ text: 'xx\ryy' });
  });

  test('a cell write replaces the text of both bodies', () => {
    const host = twoBodies();
    ok(host, [{ op: 'setTableCell', shape: tableOf(host), ...cell, text: 'z' }]);
    const again = reopen(host);
    expect(ok(again, [{ op: 'readTable', shape: tableOf(again) }])[0]).toMatchObject({
      table: { values: [['A', 'z']] },
    });
  });

  test('a text write in the second body edits the paragraph that a read gives', () => {
    const host = twoBodies();
    const t = tableOf(host);
    ok(host, [{ op: 'setText', shape: t, cell, span: { start: 3, length: 2 }, text: 'w' }]);
    expect(ok(host, [{ op: 'readText', shape: t, cell }])[0]).toMatchObject({ text: 'xx\rw' });
  });

  test('clearing the table text clears both bodies', () => {
    const host = twoBodies();
    const t = tableOf(host);
    ok(host, [{ op: 'clearTable', shape: t, text: true, format: false }]);
    expect(ok(host, [{ op: 'readTable', shape: t }])[0]).toMatchObject({
      table: { values: [['', '']] },
    });
  });

  test('a merge moves the text of both bodies into the merged cell', () => {
    const host = twoBodies();
    const t = tableOf(host);
    ok(host, [{ op: 'mergeTableCells', shape: t, row: 0, column: 0, rowCount: 1, columnCount: 2 }]);
    const again = reopen(host);
    expect(
      ok(again, [{ op: 'readTableCell', shape: tableOf(again), row: 0, column: 0 }])[0]
    ).toMatchObject({ cell: { text: 'A\rxx\ryy' } });
  });
});
