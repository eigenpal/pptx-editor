import { describe, expect, test } from 'bun:test';
import { buildPptx, cell, table } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx } from '../../store/index.ts';
import {
  resolveSlide,
  type ResolvedGraphicFrame,
  type ResolvedTable,
} from '../../cascade/index.ts';
import { createFixedMeasurer, layoutTable, layoutText, type TextMeasurer } from '../index.ts';

const IN = 914_400;

function counting(): TextMeasurer & { calls: number } {
  const inner = createFixedMeasurer();
  const m = {
    calls: 0,
    resolve: inner.resolve,
    advances: (...args: Parameters<TextMeasurer['advances']>) => {
      m.calls += 1;
      return inner.advances(...args);
    },
  };
  return m;
}

function tableOf(rows: string, columns: number[]): ResolvedTable {
  const r = readPptx(
    buildPptx({
      slides: [
        {
          shapes: table({
            id: 2,
            x: 0,
            y: 0,
            cx: IN * 6,
            cy: IN,
            columns,
            tblPr: '<a:tblPr/>',
            rows,
          }),
          layout: 'blank',
        },
      ],
    })
  );
  if (!r.ok) throw new Error('open failed');
  const frame = resolveSlide(r.pkg, 0)!.elements.find(
    (e): e is ResolvedGraphicFrame => e.kind === 'graphicFrame'
  )!;
  return frame.table!;
}

describe('table cells are laid out once (P8)', () => {
  const long = 'A cell with enough text to wrap onto several lines and grow its row';
  const rows =
    '<a:tr h="300000">' +
    cell(long) +
    cell('Centered', { tcPrAttrs: ' anchor="ctr"' }) +
    cell('Bottom', { tcPrAttrs: ' anchor="b" marT="20000" marB="90000"' }) +
    '</a:tr>' +
    '<a:tr h="600000">' +
    cell('Spans two rows', { attrs: ' rowSpan="2"', tcPrAttrs: ' anchor="b"' }) +
    cell('Two columns', { attrs: ' gridSpan="2"', tcPrAttrs: ' anchor="ctr"' }) +
    cell('', { attrs: ' hMerge="1"' }) +
    '</a:tr>' +
    `<a:tr h="300000">${cell('', { attrs: ' vMerge="1"' })}${cell('y')}${cell('z')}</a:tr>`;
  const columns = [IN, IN * 2, IN * 3];

  test('each cell equals its text laid out in its final box', () => {
    const t = tableOf(rows, columns);
    const m = createFixedMeasurer();
    const laid = layoutTable(t, m);
    expect(laid.rowY[1]).toBeGreaterThan(300_000);
    for (const c of laid.cells)
      expect(c.text).toEqual(
        layoutText(c.cell.text, { x: c.x, y: c.y, width: c.width, height: c.height }, m)
      );
  });

  test('a table measures each cell once', () => {
    const t = tableOf(rows, columns);
    const once = counting();
    for (const c of layoutTable(t, createFixedMeasurer()).cells)
      layoutText(c.cell.text, { x: c.x, y: c.y, width: c.width, height: c.height }, once);
    const table = counting();
    layoutTable(t, table);
    expect(table.calls).toBe(once.calls);
  });
});

describe('bounded table border paint', () => {
  const border = (side: string) =>
    `<a:${side} w="1"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="sysDot"/></a:${side}>`;

  test('keeps small dashed borders and bounds long horizontal, vertical, and diagonal borders', () => {
    const borders = ['lnL', 'lnR', 'lnT', 'lnB', 'lnTlToBr', 'lnBlToTr'].map(border).join('');
    const small = tableOf('<a:tr h="100">' + cell('', { tcPr: borders }) + '</a:tr>', [100]);
    const smallCell = layoutTable(small, createFixedMeasurer()).cells[0];
    expect((smallCell.paintBorders ?? smallCell.cell.borders).top).toBe(smallCell.cell.borders.top);
    expect(smallCell.cell.borders.top?.dash).toBe('sysDot');
    const large = tableOf(
      '<a:tr h="2147483647">' + cell('', { tcPr: borders }) + '</a:tr>',
      [2147483647]
    );
    const painted = layoutTable(large, createFixedMeasurer()).cells[0];
    expect(painted.cell).toBe(large.cells[0]);
    for (const side of ['top', 'bottom', 'left', 'right', 'tlToBr', 'blToTr'] as const) {
      expect(painted.cell.borders[side]?.dash).toBe('sysDot');
      expect(painted.paintBorders?.[side]?.dash).toBe('solid');
    }
  });

  test('the horizontal budget includes the border start', () => {
    for (const width of [199_998, 200_000]) {
      const source = tableOf('<a:tr h="100">' + cell('', { tcPr: border('lnT') }) + '</a:tr>', [
        width,
      ]);
      const painted = layoutTable(source, createFixedMeasurer()).cells[0];
      expect(painted.paintBorders?.top?.dash).toBe(width === 199_998 ? undefined : 'solid');
      expect(source.cells[0].borders.top?.dash).toBe('sysDot');
    }
  });
});
