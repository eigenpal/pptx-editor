// Field rules of the Equation dialog (Insert > Equation): the gallery of common equations, the
// draft the dialog opens with, what blocks OK, and the command OK applies. Framework-free, so
// the React and Vue dialogs agree. Every export is adapter glue, not a stable API: `@internal`.

import type { EditorCommand, SelectedEquationState } from '../contracts/index.ts';
import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { MAX_LINEAR_LENGTH, parseLinear } from '../automation/commands/equation-linear.ts';
import type { DialogErrors } from './dialog-draft.ts';

export { MAX_LINEAR_LENGTH };

/** One equation of the gallery: its name and its linear form. @internal */
export interface EquationGalleryItem {
  readonly id: string;
  readonly labelKey: TranslationKey;
  readonly linear: string;
}

/** The gallery of the Equation dialog. A pick fills the linear form. @internal */
export const EQUATION_GALLERY: readonly EquationGalleryItem[] = [
  { id: 'circleArea', labelKey: 'equation.galleryItems.circleArea', linear: 'A=\\pi r^2' },
  {
    id: 'quadratic',
    labelKey: 'equation.galleryItems.quadratic',
    linear: 'x=(-b\\pm\\sqrt(b^2-4ac))/(2a)',
  },
  { id: 'pythagoras', labelKey: 'equation.galleryItems.pythagoras', linear: 'a^2+b^2=c^2' },
  { id: 'sum', labelKey: 'equation.galleryItems.sum', linear: '\\sum_(k=1)^n k=(n(n+1))/2' },
  { id: 'integral', labelKey: 'equation.galleryItems.integral', linear: '\\int_0^1 x^2 dx=1/3' },
  {
    id: 'limit',
    labelKey: 'equation.galleryItems.limit',
    linear: 'lim_(n\\to\\infty) (1+1/n)^n=e',
  },
  { id: 'matrix', labelKey: 'equation.galleryItems.matrix', linear: '(\\matrix(a&b@c&d))' },
  { id: 'euler', labelKey: 'equation.galleryItems.euler', linear: 'e^(i\\pi)+1=0' },
];

/** The draft of the Equation dialog. @internal */
export interface EquationDraftFields {
  /** The linear form of the equation. */
  readonly linear: string;
}

/**
 * The draft the dialog opens with: the linear form of the selected equation, or an empty
 * field to insert a new one. @internal
 */
export function equationDraftFrom(
  selected: SelectedEquationState | null | undefined
): EquationDraftFields {
  return { linear: selected?.linear ?? '' };
}

/** Errors that block OK: an empty field, or text that the linear form cannot read. @internal */
export function validateEquationDraft(
  fields: EquationDraftFields
): DialogErrors<EquationDraftFields> {
  if (!fields.linear.trim()) return { linear: 'errors.equationEmpty' };
  try {
    parseLinear(fields.linear);
    return {};
  } catch {
    return { linear: 'errors.equationSyntax' };
  }
}

/**
 * The command OK applies: Edit Equation on the selected equation when the linear form can
 * write it, else Insert Equation. Null when the selected equation did not change. @internal
 */
export function equationDraftCommand(
  fields: EquationDraftFields,
  selected: SelectedEquationState | null | undefined
): EditorCommand | null {
  if (selected && selected.linear !== null) {
    if (fields.linear === selected.linear) return null;
    return { type: 'setEquation', nodeId: selected.nodeId, linear: fields.linear };
  }
  return { type: 'insertEquation', linear: fields.linear };
}

/** Whether the dialog edits the selected equation (else it inserts a new one). @internal */
export function equationDialogEdits(selected: SelectedEquationState | null | undefined): boolean {
  return selected != null && selected.linear !== null;
}
