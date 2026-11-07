// Text fit once shaped measurement arrives late. While the deck measures with the fixed
// measurer, an edit fits its text (`normAutofit` scales, `spAutoFit` sizes, table rows) with
// estimated widths and writes the result. When the text engine and its faces arrive, the shapes
// that those edits fit fit again with real widths, so the saved values match the text that
// shows. A shape whose XML is back to the opened file's (an undone edit) and every shape that
// no edit fit keep the values the file has. The fit is one undo step of its own, "fit text".

import { refitAll } from '../automation/commands/autofit.ts';
import { OpBatch } from '../automation/commands/op-batch.ts';
import { fittedWith } from '../automation/commands/table-fit.ts';
import type { TextMeasurer } from '../layout/index.ts';
import { nodeById, presentationInfo, xmlPartOf, type PptxPackage } from '../store/index.ts';
import type { DocumentSession } from './document-session.ts';

/**
 * Fit again, with the session's measurer, the shapes and table frames that `fixed` fit whose
 * XML differs from `opened`, as one undoable change. Does nothing when no fit changes.
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
    // The store is persistent: a shape that no landed edit changed keeps its object.
    const changed = [...ids].filter((id) => {
      const node = nodeById(part, id);
      return node !== undefined && (!was || nodeById(was, id) !== node);
    });
    refitAll(batch, index, s.partName, changed, session.measurer);
  });
  if (batch.ops.length > 0) session.store.transact(batch.ops, { label: 'fit text' });
}
