// The selected equation, as the snapshot reports it (`EditorSnapshot.equation`) for Edit
// Equation: the one equation (`a14:m`) of the one selected text box, or of the text box that
// holds the caret, with its linear form.

import type { EditorSelection, SelectedEquationState } from '../contracts/index.ts';
import { elements, nodeIndex, xmlPartOf, type Element, type PptxPackage } from '../store/index.ts';
import { isMathZone, isSupportedMath } from '../cascade/math.ts';
import { linearOfMath } from '../automation/commands/equation-text.ts';

/** The shape the selection is about: one selected element, or the text box of a caret. */
function selectedShape(selection: EditorSelection): string | undefined {
  if (selection.kind === 'shapes')
    return selection.nodeIds.length === 1 ? selection.nodeIds[0] : undefined;
  if (selection.kind === 'text') return selection.nodeId;
  return undefined;
}

/** The equations of a shape's text body that the layout draws. */
function equationsOf(shape: Element): Element[] {
  const body = elements(shape).find((c) => c.local === 'txBody');
  const out: Element[] = [];
  for (const p of elements(body))
    for (const c of elements(p)) if (isMathZone(c) && isSupportedMath(c)) out.push(c);
  return out;
}

/** The selected equation, or null. */
export function equationStateAt(
  pkg: PptxPackage | undefined,
  part: string | undefined,
  selection: EditorSelection
): SelectedEquationState | null {
  const id = selectedShape(selection);
  const xml = pkg && part && id ? xmlPartOf(pkg, part) : undefined;
  const node = xml && id ? nodeIndex(xml).get(id)?.node : undefined;
  if (node?.type !== 'element' || node.kind !== 'sp') return null;
  const zones = equationsOf(node);
  if (zones.length !== 1) return null;
  return { nodeId: zones[0]!.id, linear: linearOfMath(zones[0]!) ?? null };
}
