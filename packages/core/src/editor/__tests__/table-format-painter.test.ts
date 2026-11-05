// The format painter in a table cell (format-painter.ts): a copy from the text of one cell
// paints the text of another cell only, as one undo step, and the paint survives save and
// reopen.

import { describe, expect, test } from 'bun:test';
import { buildPptx, cell, table } from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  child,
  children,
  elements,
  find,
  readPptx,
  semanticDigest,
  serializeXmlPart,
  specOf,
  textOf,
  writePptx,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import { DEFAULT_TABLE_STYLE_ID } from '../../cascade/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const IN = 914_400;

/**
 * A 3 × 3 table: row 0 "A B C", row 1 "D E F", row 2 "G H I". It has no header row, so the
 * table style makes no cell bold.
 */
function deck(): Uint8Array {
  const letters = (r: number) =>
    ['ABC', 'DEF', 'GHI'][r]
      .split('')
      .map((t) => cell(t))
      .join('');
  return buildPptx({
    slides: [
      {
        shapes: table({
          id: 3,
          x: IN,
          y: IN,
          cx: 3 * IN,
          cy: 3 * 370_840,
          columns: [IN, IN, IN],
          tblPr: `<a:tblPr bandRow="1"><a:tableStyleId>${DEFAULT_TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>`,
          rows: [0, 1, 2].map((r) => `<a:tr h="370840">${letters(r)}</a:tr>`).join(''),
        }),
      },
    ],
  });
}

async function open(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(deck())).ok).toBe(true);
  return e;
}

const part = (e: PptxEditorInstance) => e.session!.info.slides[0].partName;
const root = (e: PptxEditorInstance): Element => xmlPartOf(e.session!.store.package, part(e))!.root;
const slideXml = (e: PptxEditorInstance): string =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, part(e))!);
const frameId = (e: PptxEditorInstance): string =>
  elements(find(root(e), 'p:spTree')!).find((c) => c.local === 'graphicFrame')!.id;
/** The `a:tc` elements of the first table under `top`, row by row. */
const tcs = (top: Element): Element[][] =>
  children(find(top, 'a:tbl'), 'a:tr').map((tr) => children(tr, 'a:tc'));
/** The `b` attribute of every run of a cell. */
const bold = (tc: Element): (string | undefined)[] =>
  children(child(tc, 'a:txBody'), 'a:p')
    .flatMap((p) => children(p, 'a:r'))
    .map((r) => attr(child(r, 'a:rPr'), 'b'));
/** A cell as JSON without node ids, to compare cells before and after an edit. */
const cellJson = (tc: Element): string => JSON.stringify(specOf(tc));

/** Select the text of a cell, from `from` to `to`. */
const textIn = (e: PptxEditorInstance, row: number, column: number, from: number, to: number) =>
  e.setSelection({
    kind: 'text',
    nodeId: frameId(e),
    cell: { row, column },
    anchor: { paragraph: 0, offset: from },
    focus: { paragraph: 0, offset: to },
  });

describe('format painter in a table cell', () => {
  test('a copy from one cell paints the text of another cell only, undoes, and round-trips', async () => {
    const e = await open();
    expect(bold(tcs(root(e))[0][0])).toEqual([undefined]);
    textIn(e, 0, 0, 0, 1);
    expect(e.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(bold(tcs(root(e))[0][0])).toEqual(['1']);

    // Copy at the bold text, then select the text of cell (1,1).
    textIn(e, 0, 0, 0, 1);
    expect(e.exec({ type: 'copyFormat' }).ok).toBe(true);
    textIn(e, 1, 1, 0, 1);
    expect(e.snapshot().format?.bold).toBe(false);
    const xmlBefore = slideXml(e);
    const before = tcs(root(e)).map((row) => row.map(cellJson));
    expect(bold(tcs(root(e))[1][1])).toEqual([undefined]);

    expect(e.exec({ type: 'pasteFormat' }).ok).toBe(true);
    const after = tcs(root(e));
    expect(bold(after[1][1])).toEqual(['1']);
    expect(textOf(child(after[1][1], 'a:txBody'))).toBe('E');
    // The target keeps its language.
    expect(attr(find(after[1][1], 'a:rPr'), 'lang')).toBe('en-US');
    expect(e.snapshot().format?.bold).toBe(true);
    // Every other cell is unchanged.
    for (let r = 0; r < 3; r += 1)
      for (let c = 0; c < 3; c += 1)
        if (r !== 1 || c !== 1)
          expect({ r, c, tc: cellJson(after[r][c]) }).toEqual({ r, c, tc: before[r][c] });

    // Save and reopen keep the paint.
    const bytes = e.save()!;
    const reopened = readPptx(bytes);
    expect(reopened.ok).toBe(true);
    if (!reopened.ok) return;
    expect(semanticDigest(reopened.pkg)).toBe(semanticDigest(e.session!.store.package));
    const again = readPptx(writePptx(reopened.pkg));
    expect(again.ok && semanticDigest(again.pkg)).toBe(semanticDigest(reopened.pkg));
    const saved = tcs(xmlPartOf(reopened.pkg, part(e))!.root);
    expect(bold(saved[1][1])).toEqual(['1']);
    expect(bold(saved[0][0])).toEqual(['1']);
    expect(bold(saved[1][0])).toEqual([undefined]);

    // The paste is one undo step.
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(slideXml(e)).toBe(xmlBefore);
    expect(bold(tcs(root(e))[1][1])).toEqual([undefined]);
    expect(bold(tcs(root(e))[0][0])).toEqual(['1']);
  });

  test('a caret in a cell paints the word under it', async () => {
    const e = await open();
    textIn(e, 0, 0, 0, 1);
    expect(e.exec({ type: 'toggleItalic' }).ok).toBe(true);
    textIn(e, 0, 0, 1, 1);
    expect(e.exec({ type: 'toggleFormatPainter' }).ok).toBe(true);
    expect(e.snapshot().formatPainter).toBe('once');
    textIn(e, 2, 2, 1, 1);
    expect(e.exec({ type: 'pasteFormat' }).ok).toBe(true);
    expect(attr(find(tcs(root(e))[2][2], 'a:rPr'), 'i')).toBe('1');
    expect(attr(find(tcs(root(e))[2][1], 'a:rPr'), 'i')).toBeUndefined();
    expect(e.snapshot().formatPainter).toBe('off');
  });

  test('a cell range and a selected table take the copied text format in every cell', async () => {
    const e = await open();
    textIn(e, 0, 0, 0, 1);
    expect(e.exec({ type: 'toggleItalic' }).ok).toBe(true);
    expect(e.exec({ type: 'copyFormat' }).ok).toBe(true);
    const nodeId = (e.snapshot().selection as { nodeId: string }).nodeId;
    e.setSelection({
      kind: 'cells',
      nodeId,
      anchor: { row: 1, column: 0 },
      focus: { row: 1, column: 1 },
    });
    expect(e.exec({ type: 'pasteFormat' }).ok).toBe(true);
    expect(attr(find(tcs(root(e))[1][0], 'a:rPr'), 'i')).toBe('1');
    expect(attr(find(tcs(root(e))[1][1], 'a:rPr'), 'i')).toBe('1');
    expect(attr(find(tcs(root(e))[2][2], 'a:rPr'), 'i')).toBeUndefined();
    e.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
    expect(e.exec({ type: 'pasteFormat' }).ok).toBe(true);
    expect(attr(find(tcs(root(e))[2][2], 'a:rPr'), 'i')).toBe('1');
  });
});
