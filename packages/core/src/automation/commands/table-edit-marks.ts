// Structure marks, the write side: which tables of a collaboration room a structure command
// (insert or delete rows and columns, merge, split) changed. The read side, and why a marked
// table draws and saves by `editedTableGrid`, is `store/package/table-marks.ts`.
//
// A mark is an `a:ext` element that the room keeps as a direct child of the table frame's
// `p:cNvPr`, so it travels with the room's state: every replica reads the same marks, whatever
// deck it opened, and saves the same bytes. Each structure command adds its own mark, and no
// command writes into another's: an undo in a room takes away only the writes of the step it
// undoes, so a peer's later structure edit keeps its mark. That is also why the mark is not in
// an `a:extLst`: a list that one step created would go with that step's undo, and the marks of
// later steps with it. The room is not a file, so the position does not need to follow
// `CT_NonVisualDrawingProps`; a mark never reaches a file, as a save, a copy, and an open remove
// every mark. A peer on a build without marks would save them, so the room format names them
// (`SCHEMA_VERSION` 3 and later in `@pptx-editor.dev/pro`), and such a peer refuses the room.
// Outside a room every table command repairs the table it edits, so no structure command leaves
// faults and nothing is marked.

import {
  child,
  elements,
  isStructureEdited,
  NS,
  structureMarksOf as marksOf,
  TABLE_EDIT_URI,
  type Element,
  type NodeSpec,
  type PptxPackage,
} from '../../store/index.ts';
import { OpBatch } from './op-batch.ts';

export { isStructureEdited, TABLE_EDIT_URI } from '../../store/index.ts';

/**
 * The most marks one frame keeps. Past it, a step adds its mark and removes the oldest one, so
 * the newest steps keep theirs: an undo of every older step still leaves the table marked.
 */
const MAX_MARKS = 64;

const cNvPrOf = (frame: Element | undefined): Element | undefined =>
  child(child(frame, 'p:nvGraphicFramePr'), 'p:cNvPr');

/** Add a structure mark of this step to the table of frame `frameId`. */
export function markStructureEdit(batch: OpBatch, part: string, frameId: string): void {
  const frame = batch.element(part, frameId);
  const cNvPr = cNvPrOf(frame);
  if (!cNvPr) return;
  const marks = marksOf(frame);
  if (marks.length >= MAX_MARKS) batch.push({ op: 'removeNode', part, node: marks[0].id });
  batch.insert({
    op: 'insertNode',
    part,
    parent: cNvPr.id,
    index: cNvPrOf(batch.element(part, frameId))!.children.length,
    spec: {
      ns: NS.a,
      local: 'ext',
      prefix: 'a',
      attrs: [{ ns: '', local: 'uri', value: TABLE_EDIT_URI }],
      children: [],
    },
  });
}

/** Remove every structure mark of frame `frameId`. */
export function unmarkStructureEdit(batch: OpBatch, part: string, frameId: string): void {
  for (const mark of marksOf(batch.element(part, frameId)))
    batch.push({ op: 'removeNode', part, node: mark.id });
}

/**
 * The ids of the marked frames under `root`, in groups and in each branch of
 * `mc:AlternateContent` too (the cascade draws the frame of one branch, so a command can mark
 * it). The walk stops at the depth the reader caps.
 */
export function markedFrames(root: Element): string[] {
  const out: string[] = [];
  const walk = (el: Element, depth: number): void => {
    if (el.kind === 'graphicFrame') {
      if (isStructureEdited(el)) out.push(el.id);
      return;
    }
    if (depth < 256) for (const c of elements(el)) walk(c, depth + 1);
  };
  walk(root, 0);
  return out;
}

/** Whether a spec (a clipboard payload) is a structure mark. */
export const isStructureMarkSpec = (spec: NodeSpec): boolean =>
  spec.local === 'ext' &&
  (spec.attrs ?? []).some((a) => a.ns === '' && a.local === 'uri' && a.value === TABLE_EDIT_URI);

/** Whether an `a:extLst` spec holds structure marks and nothing else. */
export function isOnlyStructureMarks(spec: NodeSpec): boolean {
  const kids = (spec.children ?? []).filter((c) => !('text' in c));
  return kids.length > 0 && kids.every((x) => !('text' in x) && isStructureMarkSpec(x));
}

/**
 * The package without structure marks. A mark never reaches a file, so one in a file that is
 * opened is forged: a save would otherwise rewrite the merge markup of the table it is on.
 */
export function withoutEditorMarks(pkg: PptxPackage): PptxPackage {
  const batch = new OpBatch(pkg);
  const walk = (part: string, el: Element, depth: number): void => {
    if (el.kind === 'graphicFrame' && isStructureEdited(el))
      unmarkStructureEdit(batch, part, el.id);
    if (depth < 256) for (const c of elements(el)) walk(part, c, depth + 1);
  };
  for (const part of pkg.parts.values())
    if (part.kind === 'xml' && part.xml.root.ns !== '' && /\/slide/i.test(part.partName))
      walk(part.partName, part.xml.root, 0);
  return batch.ops.length ? batch.pkg : pkg;
}

/**
 * A copied shape without structure marks, so a copy never carries one out of the editor. The
 * walk reads the whole spec, `mc:AlternateContent` branches included; a spec is as deep as the
 * tree it came from, which the reader caps.
 */
export function withoutStructureMarks(spec: NodeSpec): NodeSpec {
  if (!spec.children?.length) return spec;
  return {
    ...spec,
    children: spec.children
      .filter((c) => 'text' in c || !isStructureMarkSpec(c))
      .map((c) => ('text' in c ? c : withoutStructureMarks(c))),
  };
}
