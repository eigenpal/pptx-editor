// Table format and preset adjust values through the automation host: cell borders and margins,
// table style options and styles, and `a:avLst` guides. Every write saves and opens again.

import { describe, expect, test } from 'bun:test';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { child, elements, readPptx, xmlPartOf, type Element } from '../../store/index.ts';
import {
  createServerAutomationHost,
  type AutomationHandle,
  type AutomationHost,
  type AutomationOperation,
  type AutomationValue,
} from '../index.ts';

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

function refused(
  host: AutomationHost,
  operations: AutomationOperation[]
): { code: string; detail?: string } {
  const before = host.revision();
  const response = host.execute({ operations });
  expect(response.ok).toBe(false);
  expect(host.revision()).toBe(before);
  const failed = response.results.find((r) => r.status === 'error');
  if (failed?.status !== 'error') throw new Error('no error result');
  return {
    code: failed.error.code,
    ...(failed.error.detail ? { detail: failed.error.detail } : {}),
  };
}

function handle(value: AutomationValue | undefined): AutomationHandle {
  if (value?.kind !== 'handle') throw new Error(`not a handle: ${JSON.stringify(value)}`);
  return value.handle;
}

function reopen(host: AutomationHost): AutomationHost {
  const saved = host.save();
  if (!saved.ok) throw new Error(saved.error.code);
  expect(readPptx(saved.bytes).ok).toBe(true);
  return open(saved.bytes);
}

/** The first shape of the first slide. */
function firstShape(host: AutomationHost): AutomationHandle {
  const [, s] = ok(host, [
    { op: 'getSlide', key: { index: 0 } },
    { op: 'getShape', container: { result: 0 }, key: { index: 0 } },
  ]);
  return handle(s);
}

function withTable(
  rows = 3,
  columns = 3,
  bytes?: Uint8Array
): { host: AutomationHost; t: AutomationHandle } {
  const host = open(bytes);
  const [, t] = ok(host, [
    { op: 'getSlide', key: { index: 0 } },
    { op: 'addShape', container: { result: 0 }, shape: { kind: 'table', rows, columns } },
  ]);
  return { host, t: handle(t) };
}

/** The `a:tc` elements of the first table of slide 1, row by row. */
function cellsOf(host: AutomationHost): Element[][] {
  const saved = host.save();
  if (!saved.ok) throw new Error(saved.error.code);
  const read = readPptx(saved.bytes);
  if (!read.ok) throw new Error('unreadable');
  const root = xmlPartOf(read.pkg, '/ppt/slides/slide1.xml')!.root;
  const found: Element[] = [];
  const visit = (el: Element): void => {
    if (el.local === 'tbl') found.push(el);
    for (const c of elements(el)) visit(c);
  };
  visit(root);
  return elements(found[0])
    .filter((e) => e.local === 'tr')
    .map((tr) => elements(tr).filter((e) => e.local === 'tc'));
}

function formatOf(
  host: AutomationHost,
  t: AutomationHandle,
  row: number,
  column: number
): Extract<AutomationValue, { kind: 'tableCellFormat' }>['format'] {
  const [value] = ok(host, [{ op: 'readTableCellFormat', shape: t, cell: { row, column } }]);
  if (value?.kind !== 'tableCellFormat') throw new Error('not a cell format');
  return value.format;
}

describe('automation host: table cell borders', () => {
  test('a border write keeps the fields it leaves out and agrees with the neighbor', () => {
    const { host, t } = withTable();
    const cell = { row: 1, column: 1 };
    ok(host, [
      { op: 'setTableCellBorder', shape: t, cell, side: 'left', line: { color: '#FF0000' } },
      { op: 'setTableCellBorder', shape: t, cell, side: 'left', line: { weight: 3 } },
      {
        op: 'setTableCellBorder',
        shape: t,
        cell,
        side: 'diagonalDown',
        line: { dashStyle: 'dash', transparency: 0.5, color: '#00FF00' },
      },
    ]);
    const again = reopen(host);
    const u = firstShape(again);
    const own = formatOf(again, u, 1, 1);
    expect(own.borders.left).toMatchObject({ visible: true, color: '#FF0000', weight: 3 });
    expect(own.borders.diagonalDown).toMatchObject({
      visible: true,
      color: '#00FF00',
      dashStyle: 'dash',
      transparency: 0.5,
    });
    // The cell across the left edge shows the same line on its right edge.
    expect(formatOf(again, u, 1, 0).borders.right).toMatchObject({ color: '#FF0000', weight: 3 });
    // Schema order inside a:tcPr: lnL before lnTlToBr.
    const tcPr = child(cellsOf(again)[1]![1], 'a:tcPr');
    const names = elements(tcPr).map((e) => e.local);
    expect(names.indexOf('lnL')).toBeLessThan(names.indexOf('lnTlToBr'));
  });

  test('a merged block writes its edge on every cell across it', () => {
    const { host, t } = withTable();
    ok(host, [
      { op: 'mergeTableCells', shape: t, row: 0, column: 0, rowCount: 2, columnCount: 1 },
      {
        op: 'setTableCellBorder',
        shape: t,
        cell: { row: 1, column: 0 },
        side: 'right',
        line: { color: '#0000FF', weight: 2 },
      },
    ]);
    const again = reopen(host);
    const u = firstShape(again);
    expect(formatOf(again, u, 0, 1).borders.left).toMatchObject({ color: '#0000FF', weight: 2 });
    expect(formatOf(again, u, 1, 1).borders.left).toMatchObject({ color: '#0000FF', weight: 2 });
  });

  test('bad sides, colors, dashes, and weights are refused, and nothing is written', () => {
    const { host, t } = withTable();
    const cell = { row: 0, column: 0 };
    const side = 'top' as const;
    expect(
      refused(host, [
        { op: 'setTableCellBorder', shape: t, cell, side: 'inside' as never, line: {} },
      ]).code
    ).toBe('invalid-argument');
    expect(
      refused(host, [
        { op: 'setTableCellBorder', shape: t, cell, side, line: { color: 'not a color' } },
      ]).code
    ).toBe('invalid-argument');
    expect(
      refused(host, [
        { op: 'setTableCellBorder', shape: t, cell, side, line: { dashStyle: 'wavy' } },
      ]).code
    ).toBe('invalid-argument');
    expect(
      refused(host, [{ op: 'setTableCellBorder', shape: t, cell, side, line: { weight: -1 } }]).code
    ).toBe('invalid-argument');
    expect(
      refused(host, [
        { op: 'setTableCellBorder', shape: t, cell: { row: 9, column: 0 }, side, line: {} },
      ]).code
    ).toBe('invalid-argument');
  });

  test('a shape that is not a table is refused', () => {
    const host = open(
      buildPptx({
        slides: [{ layout: 'blank', shapes: shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100 }) }],
      })
    );
    const s = firstShape(host);
    expect(
      refused(host, [{ op: 'readTableCellFormat', shape: s, cell: { row: 0, column: 0 } }])
    ).toEqual({ code: 'unsupported-content', detail: 'not-table' });
  });
});

describe('automation host: table cell margins', () => {
  test('margins read their defaults, write in points, and reset with null', () => {
    const { host, t } = withTable();
    expect(formatOf(host, t, 0, 0).margins).toEqual({
      left: 7.2,
      right: 7.2,
      top: 3.6,
      bottom: 3.6,
    });
    ok(host, [
      {
        op: 'setTableCellMargins',
        shape: t,
        cell: { row: 0, column: 0 },
        margins: { left: 10, top: 0, bottom: 12.5 },
      },
    ]);
    const again = reopen(host);
    const u = firstShape(again);
    expect(formatOf(again, u, 0, 0).margins).toEqual({
      left: 10,
      right: 7.2,
      top: 0,
      bottom: 12.5,
    });
    ok(again, [
      { op: 'setTableCellMargins', shape: u, cell: { row: 0, column: 0 }, margins: { left: null } },
    ]);
    expect(formatOf(again, u, 0, 0).margins.left).toBe(7.2);
    expect(
      refused(again, [
        {
          op: 'setTableCellMargins',
          shape: u,
          cell: { row: 0, column: 0 },
          margins: { right: -1 },
        },
      ]).code
    ).toBe('invalid-argument');
    expect(
      refused(again, [
        {
          op: 'setTableCellMargins',
          shape: u,
          cell: { row: 0, column: 0 },
          margins: { right: Number.NaN },
        },
      ]).code
    ).toBe('invalid-argument');
  });
});

const NAMED_STYLE =
  '<a:tblStyle styleId="{9D7B26C5-4107-4FEC-AEDC-1716B250EE53}" styleName="Light Style 1">' +
  '<a:wholeTbl><a:tcStyle><a:tcBdr/></a:tcStyle></a:wholeTbl></a:tblStyle>';

describe('automation host: table style settings', () => {
  test('style options and a named style write a:tblPr and read back', () => {
    const bytes = buildPptx({
      slides: [{ layout: 'blank', shapes: '' }],
      tableStyles: NAMED_STYLE,
    });
    const { host, t } = withTable(3, 3, bytes);
    const [before] = ok(host, [{ op: 'readTable', shape: t }]);
    expect(before).toMatchObject({
      table: {
        style: { styleName: 'Medium Style 2 - Accent 1', firstRow: true, bandedRows: true },
      },
    });
    ok(host, [
      {
        op: 'setTableStyle',
        shape: t,
        style: {
          firstRow: false,
          lastRow: true,
          firstColumn: true,
          bandedColumns: true,
          styleName: 'Light Style 1',
        },
      },
    ]);
    const again = reopen(host);
    const u = firstShape(again);
    const [after] = ok(again, [{ op: 'readTable', shape: u }]);
    expect(after).toMatchObject({
      table: {
        style: {
          styleId: '{9D7B26C5-4107-4FEC-AEDC-1716B250EE53}',
          styleName: 'Light Style 1',
          firstRow: false,
          lastRow: true,
          firstColumn: true,
          lastColumn: false,
          bandedRows: true,
          bandedColumns: true,
        },
      },
    });
    ok(again, [{ op: 'setTableStyle', shape: u, style: { styleName: null } }]);
    const [none] = ok(again, [{ op: 'readTable', shape: u }]);
    expect(none).toMatchObject({ table: { style: { styleId: null, styleName: null } } });
  });

  test('a style the presentation does not define is refused', () => {
    const { host, t } = withTable();
    expect(
      refused(host, [{ op: 'setTableStyle', shape: t, style: { styleName: 'Dark Style 2' } }])
    ).toEqual({ code: 'unsupported-content', detail: 'table-style-undefined' });
    expect(
      refused(host, [{ op: 'setTableStyle', shape: t, style: { firstRow: 'yes' as never } }]).code
    ).toBe('invalid-argument');
  });
});

describe('automation host: preset adjust values', () => {
  const rounded = (avLst = ''): Uint8Array =>
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: 0,
            y: 0,
            cx: 1_270_000,
            cy: 635_000,
            prst: 'roundRect',
            avLst,
          }),
        },
      ],
    });

  test('a preset reads its defaults, and a write bounds the value to the handle range', () => {
    const host = open(rounded());
    const s = firstShape(host);
    expect(ok(host, [{ op: 'readAdjustments', shape: s }])).toEqual([
      { kind: 'adjustments', adjustments: { preset: 'roundRect', values: [0.16667] } },
    ]);
    ok(host, [{ op: 'setAdjustment', shape: s, index: 0, value: 0.3 }]);
    const again = reopen(host);
    const u = firstShape(again);
    expect(ok(again, [{ op: 'readAdjustments', shape: u }])).toEqual([
      { kind: 'adjustments', adjustments: { preset: 'roundRect', values: [0.3] } },
    ]);
    // roundRect's handle runs from 0 to 50 000 (presetShapeDefinitions.xml).
    ok(again, [{ op: 'setAdjustment', shape: u, index: 0, value: 4 }]);
    expect(ok(again, [{ op: 'readAdjustments', shape: u }])).toEqual([
      { kind: 'adjustments', adjustments: { preset: 'roundRect', values: [0.5] } },
    ]);
  });

  test('an angle reads in degrees, and a range that names a guide is evaluated', () => {
    const host = open(
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes:
              shape({ id: 2, x: 0, y: 0, cx: 1_270_000, cy: 1_270_000, prst: 'arc' }) +
              shape({ id: 3, x: 0, y: 0, cx: 2_540_000, cy: 1_270_000, prst: 'rightArrow' }),
          },
        ],
      })
    );
    const [, arc, arrow] = ok(host, [
      { op: 'getSlide', key: { index: 0 } },
      { op: 'getShape', container: { result: 0 }, key: { index: 0 } },
      { op: 'getShape', container: { result: 0 }, key: { index: 1 } },
    ]);
    const a = handle(arc);
    const r = handle(arrow);
    expect(ok(host, [{ op: 'readAdjustments', shape: a }])).toEqual([
      { kind: 'adjustments', adjustments: { preset: 'arc', values: [270, 0] } },
    ]);
    ok(host, [{ op: 'setAdjustment', shape: a, index: 1, value: 90 }]);
    // rightArrow adj2 is bounded by maxAdj2 = 100 000 * w / ss = 200 000 at 2:1.
    ok(host, [{ op: 'setAdjustment', shape: r, index: 1, value: 9 }]);
    const [arcAfter, arrowAfter] = ok(host, [
      { op: 'readAdjustments', shape: a },
      { op: 'readAdjustments', shape: r },
    ]);
    expect(arcAfter).toMatchObject({ adjustments: { values: [270, 90] } });
    expect(arrowAfter).toMatchObject({ adjustments: { values: [0.5, 2] } });
  });

  test('a placeholder that takes its geometry from its layout gets its own a:prstGeom', () => {
    const bytes = buildPptx({
      layouts: [
        {
          name: 'Rounded',
          shapes: shape({
            id: 2,
            x: 0,
            y: 0,
            cx: 1_270_000,
            cy: 1_270_000,
            prst: 'roundRect',
            avLst: '<a:gd name="adj" fmla="val 25000"/>',
            ph: '<p:ph type="body" idx="13"/>',
          }),
        },
      ],
      slides: [
        {
          layoutName: 'Rounded',
          shapes:
            '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="13"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>',
        },
      ],
    });
    const host = open(bytes);
    const s = firstShape(host);
    expect(ok(host, [{ op: 'readAdjustments', shape: s }])).toEqual([
      { kind: 'adjustments', adjustments: { preset: 'roundRect', values: [0.25] } },
    ]);
    ok(host, [{ op: 'setAdjustment', shape: s, index: 0, value: 0.1 }]);
    const again = reopen(host);
    expect(ok(again, [{ op: 'readAdjustments', shape: firstShape(again) }])).toEqual([
      { kind: 'adjustments', adjustments: { preset: 'roundRect', values: [0.1] } },
    ]);
  });

  test('shapes with no preset geometry, bad indices, and bad values are refused', () => {
    const custom = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes:
            '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Free"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' +
            '<a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:custGeom><a:pathLst>' +
            '<a:path w="1" h="1"><a:moveTo><a:pt x="0" y="0"/></a:moveTo></a:path></a:pathLst>' +
            '</a:custGeom></p:spPr></p:sp>',
        },
      ],
    });
    const host = open(custom);
    expect(refused(host, [{ op: 'readAdjustments', shape: firstShape(host) }])).toEqual({
      code: 'unsupported-content',
      detail: 'no-preset-geometry',
    });
    const { host: tables, t } = withTable();
    expect(refused(tables, [{ op: 'readAdjustments', shape: t }]).detail).toBe(
      'no-preset-geometry'
    );
    const plain = open(rounded());
    const s = firstShape(plain);
    expect(refused(plain, [{ op: 'setAdjustment', shape: s, index: 1, value: 0 }]).code).toBe(
      'invalid-argument'
    );
    expect(
      refused(plain, [{ op: 'setAdjustment', shape: s, index: 0, value: Number.POSITIVE_INFINITY }])
        .code
    ).toBe('invalid-argument');
    expect(refused(plain, [{ op: 'setAdjustment', shape: s, index: 0, value: 1e9 }]).code).toBe(
      'invalid-argument'
    );
  });
});
