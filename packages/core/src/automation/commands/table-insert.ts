// A new table frame (`p:graphicFrame` with `a:tbl`). DOM-free, so the editor's Insert > Table
// and the automation host's `addTable` write the same XML.

import { defaultTableStyleId } from '../../cascade/index.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { nextShapeId, spTreeOf } from './package-edits.ts';
import { specFromXml } from './shape-commands.ts';
import { MAX_TABLE_SIZE } from './table-model.ts';

/**
 * A new table frame centered on the slide, two thirds of its width, with 0.4" rows
 * (ECMA-376 Part 1, 21.1.3.13 tbl). Returns the frame's node id.
 */
export function insertTableFrame(
  batch: OpBatch,
  part: string,
  rows: number,
  cols: number,
  slideW: number,
  slideH: number
): string {
  if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows < 1 || cols < 1)
    throw new CommandRefusal('errors.unsupported', 'size');
  if (rows > MAX_TABLE_SIZE || cols > MAX_TABLE_SIZE)
    throw new CommandRefusal('errors.unsupported', 'size');
  // A new table is two thirds of the slide wide, with 0.4" (29.2 pt) rows.
  const colW = Math.floor((slideW * 2) / 3 / cols);
  const rowH = 370_840;
  const cx = colW * cols;
  const cy = rowH * rows;
  const tree = spTreeOf(batch, part);
  // In a collaboration session the id comes from the actor's stripe (`shapeIds`).
  const id = nextShapeId(batch, tree);
  const cell =
    '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p></a:txBody><a:tcPr/></a:tc>';
  const xml =
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id - 1}"/>` +
    '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>' +
    `<p:xfrm><a:off x="${Math.round((slideW - cx) / 2)}" y="${Math.round((slideH - cy) / 2)}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>' +
    `<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${defaultTableStyleId(batch.pkg)}</a:tableStyleId></a:tblPr>` +
    `<a:tblGrid>${`<a:gridCol w="${colW}"/>`.repeat(cols)}</a:tblGrid>` +
    `<a:tr h="${rowH}">${cell.repeat(cols)}</a:tr>`.repeat(rows) +
    '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
  const ext = tree.children.findIndex((c) => c.type === 'element' && c.local === 'extLst');
  return batch.insert({
    op: 'insertNode',
    part,
    parent: tree.id,
    index: ext < 0 ? tree.children.length : ext,
    spec: specFromXml(xml),
  });
}
