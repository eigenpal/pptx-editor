// A new table frame (`p:graphicFrame` with `a:tbl`). DOM-free, so the editor's Insert > Table
// and the automation host's `addTable` write the same XML. Insert > Table can fill an empty
// placeholder (`into`); `addTable` places the table where the caller asks.

import { defaultTableStyleId } from '../../cascade/index.ts';
import { escapeXmlAttribute } from '../../store/index.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { nextShapeId, spTreeOf } from './package-edits.ts';
import {
  placeholderIndex,
  removeFilledPlaceholder,
  type EmptyPlaceholder,
} from './placeholder-fill.ts';
import { specFromXml } from './shape-commands.ts';
import { MAX_TABLE_SIZE } from './table-model.ts';

/**
 * A new table frame with 0.4" rows (ECMA-376 Part 1, 21.1.3.13 tbl). It is centered on the slide
 * at two thirds of its width, or it takes the place of the empty placeholder `into`: at its top
 * left, as wide as it, and as a placeholder with the same index (§19.3.1.36). Returns the
 * frame's node id.
 */
export function insertTableFrame(
  batch: OpBatch,
  part: string,
  rows: number,
  cols: number,
  slideW: number,
  slideH: number,
  into?: EmptyPlaceholder
): string {
  if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows < 1 || cols < 1)
    throw new CommandRefusal('errors.unsupported', 'size');
  if (rows > MAX_TABLE_SIZE || cols > MAX_TABLE_SIZE)
    throw new CommandRefusal('errors.unsupported', 'size');
  // A new table is two thirds of the slide wide, or as wide as its placeholder, with 0.4"
  // (29.2 pt) rows.
  const colW = Math.max(1, Math.floor((into ? into.cx : (slideW * 2) / 3) / cols));
  const rowH = 370_840;
  const cx = colW * cols;
  const cy = rowH * rows;
  const x = into ? into.x : Math.round((slideW - cx) / 2);
  const y = into ? into.y : Math.round((slideH - cy) / 2);
  const tree = spTreeOf(batch, part);
  // In a collaboration session the id comes from the actor's stripe (`shapeIds`).
  const id = nextShapeId(batch, tree);
  const cell =
    '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p></a:txBody><a:tcPr/></a:tc>';
  const xml =
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${into ? escapeXmlAttribute(into.name) : `Table ${id - 1}`}"/>` +
    '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>' +
    `<p:nvPr>${into ? into.ph : ''}</p:nvPr></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>' +
    `<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${defaultTableStyleId(batch.pkg)}</a:tableStyleId></a:tblPr>` +
    `<a:tblGrid>${`<a:gridCol w="${colW}"/>`.repeat(cols)}</a:tblGrid>` +
    `<a:tr h="${rowH}">${cell.repeat(cols)}</a:tr>`.repeat(rows) +
    '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
  // The table takes the placeholder's place in the z-order; a new one goes on top.
  const at = into ? placeholderIndex(batch, part, into) : -1;
  const ext = tree.children.findIndex((c) => c.type === 'element' && c.local === 'extLst');
  const frame = batch.insert({
    op: 'insertNode',
    part,
    parent: tree.id,
    index: at >= 0 ? at : ext < 0 ? tree.children.length : ext,
    spec: specFromXml(xml),
  });
  if (into && at >= 0) removeFilledPlaceholder(batch, part, into, String(id));
  return frame;
}
