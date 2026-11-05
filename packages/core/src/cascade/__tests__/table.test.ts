import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import {
  buildPptx,
  cell,
  para,
  run,
  shape,
  table,
} from '../../../../../scripts/fixtures/builder.ts';
import { readPptx, type PptxPackage } from '../../store/index.ts';
import { resolveSlide, type ResolvedGraphicFrame, type ResolvedTable } from '../index.ts';
import { createFixedMeasurer, layoutTable } from '../../layout/index.ts';

const IN = 914_400;

function pkgOf(bytes: Uint8Array): PptxPackage {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('unreadable');
  return r.pkg;
}

function tableOn(pkg: PptxPackage, slide = 0): ResolvedTable {
  const frame = resolveSlide(pkg, slide)!.elements.find(
    (e): e is ResolvedGraphicFrame => e.kind === 'graphicFrame'
  )!;
  expect(frame.table).toBeDefined();
  return frame.table!;
}

const hex = (fill: { type: string; color?: { r: number; g: number; b: number } }): string =>
  fill.type === 'solid' && fill.color
    ? [fill.color.r, fill.color.g, fill.color.b]
        .map((v) => Math.round(v).toString(16).padStart(2, '0'))
        .join('')
    : fill.type;

const at = (t: ResolvedTable, r: number, c: number) => t.cells[r * t.columns.length + c];

describe('table styles', () => {
  const fixture = pkgOf(new Uint8Array(readFileSync('fixtures/tables.pptx')));

  test('the built-in default style: header, bands, and body', () => {
    const t = tableOn(fixture, 0);
    // The default theme accent1 is 4472C4.
    expect(hex(at(t, 0, 0).fill)).toBe('4472c4');
    expect(at(t, 0, 0).text.paragraphs[0].runs[0].props.bold).toBe(true);
    expect(hex(at(t, 0, 0).text.paragraphs[0].runs[0].props.fill)).toBe('ffffff');
    // band1H (tint 40 %) on the first body row, wholeTbl (tint 20 %) on the next.
    const band = at(t, 1, 0).fill;
    const whole = at(t, 2, 0).fill;
    expect(hex(band)).not.toBe(hex(whole));
    expect(at(t, 1, 0).text.paragraphs[0].runs[0].props.bold).toBe(false);
    // The header's bottom border is 3 pt; inner borders are 1 pt white.
    expect(at(t, 0, 1).borders.bottom?.width).toBe(38_100);
    expect(at(t, 2, 1).borders.left?.width).toBe(12_700);
  });

  test('a file style: first column, last row, banded columns, corner, and background', () => {
    const t = tableOn(fixture, 1);
    expect(t.styleId).toBe('{0C3A3B8E-5A47-4F0B-9C1D-6E2F7A8B9C01}');
    const first = at(t, 1, 0).text.paragraphs[0].runs[0].props;
    expect(first.bold && first.italic).toBe(true);
    // Banded columns count after the first column: column 1 is band1V, column 2 is not.
    expect(at(t, 1, 1).fill.type).toBe('solid');
    expect(at(t, 1, 2).fill.type).toBe('none');
    expect(t.background.type).toBe('solid');
    // The south-west corner part wins over firstCol and lastRow.
    expect(hex(at(t, 3, 0).fill)).toBe('70ad47');
    expect(at(t, 3, 2).borders.top?.width).toBe(38_100);
    // insideV is "no line"; the table's left edge comes from a theme line reference.
    expect(at(t, 1, 1).borders.left).toBeUndefined();
    expect(at(t, 1, 0).borders.left).toBeDefined();
    expect(at(t, 3, 2).text.paragraphs[0].runs[0].props.bold).toBe(true);
  });

  test('merges, direct borders, and diagonals', () => {
    const t = tableOn(fixture, 2);
    expect(at(t, 0, 0).gridSpan).toBe(2);
    expect(at(t, 0, 1).covered).toBe(true);
    expect(at(t, 0, 3).rowSpan).toBe(2);
    expect(at(t, 1, 3).covered).toBe(true);
    expect(at(t, 1, 0).borders.tlToBr).toBeDefined();
    expect(at(t, 1, 1).borders.top?.dash).toBe('dash');
    expect(at(t, 2, 1).text.body.lIns).toBe(0);
    expect(at(t, 1, 2).text.body.vert).toBe('vert270');
  });
});

describe('table layout', () => {
  test('rows grow to fit their text, spans cover their rows and columns', () => {
    const pkg = pkgOf(new Uint8Array(readFileSync('fixtures/tables.pptx')));
    const t = tableOn(pkg, 2);
    const laid = layoutTable(t, createFixedMeasurer());
    expect(laid.cells.length).toBe(t.cells.filter((c) => !c.covered).length);
    const rowH = laid.rowY.slice(1).map((y, i) => y - laid.rowY[i]);
    expect(rowH[1]).toBe(548_640);
    expect(rowH[2]).toBeGreaterThan(548_640);
    const merged = laid.cells.find((c) => c.cell.gridSpan === 2)!;
    expect(merged.width).toBe(t.columns[0] + t.columns[1]);
    const down = laid.cells.find((c) => c.cell.rowSpan === 2)!;
    expect(down.height).toBe(rowH[0] + rowH[1]);
    expect(laid.height).toBe(laid.rowY[3]);
  });

  test('an unknown style id renders without a style; a missing grid is not a table', () => {
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes:
            table({
              id: 2,
              x: 0,
              y: 0,
              cx: 2 * IN,
              cy: IN,
              columns: [IN, IN],
              tblPr:
                '<a:tblPr firstRow="1"><a:tableStyleId>{00000000-0000-0000-0000-000000000000}</a:tableStyleId></a:tblPr>',
              rows: `<a:tr h="${IN}">${cell('a')}${cell('b')}</a:tr>`,
            }) + shape({ id: 3, x: 0, y: 2 * IN, cx: IN, cy: IN, text: para(run('x')) }),
        },
      ],
    });
    const t = tableOn(pkgOf(bytes));
    expect(t.styleId).toBeUndefined();
    expect(at(t, 0, 0).fill.type).toBe('none');
    expect(Object.keys(at(t, 0, 0).borders)).toEqual([]);
  });
});

describe('shape style font reference', () => {
  test('a shape style color beats the default text style', () => {
    const style =
      '<p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
      '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="major"><a:schemeClr val="lt1"/></a:fontRef></p:style>';
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({ id: 2, x: 0, y: 0, cx: IN, cy: IN, text: para(run('x')) }).replace(
            '</p:spPr>',
            `</p:spPr>${style}`
          ),
        },
      ],
    });
    const s = resolveSlide(pkgOf(bytes), 0)!;
    const props = (
      s.elements[0] as {
        text?: { paragraphs: { runs: { props: { fill: never; latin: string } }[] }[] };
      }
    ).text!.paragraphs[0].runs[0].props;
    expect(hex(props.fill)).toBe('ffffff');
    expect(props.latin).toBe('Calibri Light');
  });
});
