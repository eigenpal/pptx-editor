// A new table frame (`p:graphicFrame` with `a:tbl`). DOM-free, so the editor's Insert > Table
// and the automation host's `addTable` write the same XML.

import { defaultTableStyleId, resolveSlide } from '../../cascade/index.ts';
import {
  attr,
  boolAttr,
  child,
  elements,
  escapeXmlAttribute,
  path,
  textOf,
  type Element,
} from '../../store/index.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { nextShapeId, spTreeOf } from './package-edits.ts';
import { deleteShapes, specFromXml } from './shape-commands.ts';
import { MAX_TABLE_SIZE } from './table-model.ts';

/** The placeholder types a table can fill (§19.7.10 `ST_PlaceholderType`): content and table. */
const TABLE_PLACEHOLDERS = new Set(['obj', 'tbl']);

/** The `p:ph` attributes a table frame keeps from the placeholder it fills. */
const KEPT_PH_ATTRS = ['type', 'orient', 'sz', 'idx'] as const;

/** Whether a text body holds no character (a field counts as one). */
function emptyText(txBody: Element | undefined): boolean {
  if (!txBody) return true;
  const walk = (el: Element, depth: number): boolean =>
    depth < 64 &&
    elements(el).every(
      (c) => !((c.local === 't' && textOf(c).length > 0) || c.local === 'fld') && walk(c, depth + 1)
    );
  return walk(txBody, 0);
}

/** An empty content or table placeholder on the slide that a new table can take the place of. */
export interface TablePlaceholder {
  /** The `p:sp` node id. */
  readonly nodeId: string;
  /** Its box on the slide, EMU, from the slide, layout, or master. */
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  /** The `p:ph` element of the new frame, with the placeholder's own type, size, and index. */
  readonly ph: string;
}

/**
 * The empty content or table placeholder that a new table fills: `preferred` when it is one,
 * otherwise the first on the slide. A placeholder with text, a hidden one, and a rotated one are
 * left alone. Undefined when the slide has none.
 */
export function tablePlaceholderOf(
  batch: OpBatch,
  part: string,
  slideIndex: number,
  preferred?: string
): TablePlaceholder | undefined {
  const candidates = elements(spTreeOf(batch, part)).filter((el) => {
    if (el.local !== 'sp') return false;
    const ph = path(el, 'p:nvSpPr', 'p:nvPr', 'p:ph');
    if (!ph || !TABLE_PLACEHOLDERS.has(attr(ph, 'type') ?? 'obj')) return false;
    if (boolAttr(path(el, 'p:nvSpPr', 'p:cNvPr'), 'hidden')) return false;
    return emptyText(child(el, 'p:txBody'));
  });
  if (!candidates.length) return undefined;
  const el = candidates.find((c) => c.id === preferred) ?? candidates[0]!;
  const resolved = resolveSlide(batch.pkg, slideIndex)?.elements.find((r) => r.nodeId === el.id);
  if (!resolved || resolved.xfrm.rot !== 0 || resolved.xfrm.cx <= 0) return undefined;
  const ph = path(el, 'p:nvSpPr', 'p:nvPr', 'p:ph')!;
  const kept = KEPT_PH_ATTRS.map((name) => [name, attr(ph, name)] as const)
    .filter(([, v]) => v !== undefined)
    .map(([name, v]) => ` ${name}="${escapeXmlAttribute(v!)}"`)
    .join('');
  const { x, y, cx } = resolved.xfrm;
  return { nodeId: el.id, x, y, cx, ph: `<p:ph${kept}/>` };
}

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
  into?: TablePlaceholder
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
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id - 1}"/>` +
    '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>' +
    `<p:nvPr>${into ? into.ph : ''}</p:nvPr></p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>' +
    `<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${defaultTableStyleId(batch.pkg)}</a:tableStyleId></a:tblPr>` +
    `<a:tblGrid>${`<a:gridCol w="${colW}"/>`.repeat(cols)}</a:tblGrid>` +
    `<a:tr h="${rowH}">${cell.repeat(cols)}</a:tr>`.repeat(rows) +
    '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
  // The table takes the placeholder's place in the z-order; a new one goes on top.
  const at = into
    ? tree.children.findIndex((c) => c.type === 'element' && c.id === into.nodeId)
    : -1;
  const ext = tree.children.findIndex((c) => c.type === 'element' && c.local === 'extLst');
  const frame = batch.insert({
    op: 'insertNode',
    part,
    parent: tree.id,
    index: at >= 0 ? at : ext < 0 ? tree.children.length : ext,
    spec: specFromXml(xml),
  });
  if (into && at >= 0) deleteShapes(batch, part, [into.nodeId]);
  return frame;
}
