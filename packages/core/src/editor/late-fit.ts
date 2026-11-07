// Text fit once shaped measurement arrives late. While the deck measures with the fixed
// measurer, an edit fits its text (`normAutofit` scales, `spAutoFit` sizes, table rows) with
// estimated widths and writes the result. When the text engine and its faces arrive, the shapes
// that those edits fit fit again with real widths, so the saved values match the text that
// shows. A shape whose text is the opened file's again (an undone edit, a refused commit) and
// every shape that no edit fit keep the values the file has, also after a move. The fit is one
// undo step of its own.

import { refitAll } from '../automation/commands/autofit.ts';
import { OpBatch } from '../automation/commands/op-batch.ts';
import { fittedWith } from '../automation/commands/table-fit.ts';
import type { TextMeasurer } from '../layout/index.ts';
import {
  child,
  nodeById,
  presentationInfo,
  xmlPartOf,
  type Element,
  type Node,
  type PptxPackage,
} from '../store/index.ts';
import type { DocumentSession } from './document-session.ts';

/** The content of a node as text: names, attributes, and text, without node ids. */
function content(node: Node | undefined, out: string[] = []): string[] {
  if (!node) return out;
  if (node.type === 'text') {
    out.push(JSON.stringify(node.value));
    return out;
  }
  out.push(`<${node.ns} ${node.local}`);
  for (const a of node.attrs) out.push(` ${a.ns} ${a.local}=${JSON.stringify(a.value)}`);
  for (const c of node.children) content(c, out);
  out.push('>');
  return out;
}

/** The text a fit reads: a shape's text body, or a table frame's table. */
function fitText(el: Element): Node | undefined {
  return (
    child(el, 'p:txBody') ??
    child(child(child(el, 'a:graphic'), 'a:graphicData'), 'a:tbl') ??
    undefined
  );
}

/** Whether the fitted text of `now` differs from that of `was`, the opened file's node. */
function textChanged(now: Node, was: Node | undefined): boolean {
  if (now === was) return false;
  if (!was || now.type !== 'element' || was.type !== 'element') return true;
  return content(fitText(now)).join('') !== content(fitText(was)).join('');
}

/**
 * Fit again, with the session's measurer, the shapes and table frames that `fixed` fit whose
 * text differs from `opened`, as one undoable change. Does nothing when no fit changes.
 */
export function refitEditedText(
  session: DocumentSession,
  fixed: TextMeasurer,
  opened: PptxPackage
): void {
  if (session.disposed || session.fixedMeasurer) return;
  const fitted = fittedWith(fixed);
  if (fitted.size === 0) return;
  const pkg = session.store.package;
  const batch = new OpBatch(pkg);
  presentationInfo(pkg).slides.forEach((s, index) => {
    const ids = fitted.get(s.partName);
    const part = ids && xmlPartOf(pkg, s.partName);
    if (!ids || !part) return;
    const was = xmlPartOf(opened, s.partName);
    const changed = [...ids].filter((id) => {
      const node = nodeById(part, id);
      return node !== undefined && textChanged(node, was && nodeById(was, id));
    });
    refitAll(batch, index, s.partName, changed, session.measurer);
  });
  if (batch.ops.length > 0) session.store.transact(batch.ops, { label: 'fit text' });
}
