// The equation contract: the commands that insert and edit Office Math (ECMA-376 Part 1,
// 22.1), and the selected equation that the snapshot reports. DOM-free.

/**
 * The equation commands. An equation's math is written from a linear form, a short text
 * syntax: `a/b` is a fraction, `x^2` and `x_i` are scripts, `\sqrt(x)` is a square root,
 * `\sum_(i=1)^n` is a sum, and `\alpha` is a Greek letter. The equations guide lists the whole
 * syntax. @public
 */
export type EquationCommand =
  /**
   * Insert > Equation: a new text box centered on the slide, with the equation of `linear`.
   * `font` names the font of its math runs; default `Cambria Math`. The text box is selected.
   */
  | { readonly type: 'insertEquation'; readonly linear: string; readonly font?: string }
  /**
   * Replace the math of an equation (`SelectedEquationState.nodeId`) with the equation of
   * `linear`. The equation keeps its alignment and the font, size, and color of its first run.
   */
  | { readonly type: 'setEquation'; readonly nodeId: string; readonly linear: string };

/** The selected equation: one text box with one equation is selected. @public */
export interface SelectedEquationState {
  /** The node id of the equation (`a14:m`), for `setEquation`. */
  readonly nodeId: string;
  /**
   * The linear form of the equation, for Edit Equation. Null when the equation has a construct,
   * run property, or character that the linear form does not cover: then only its box, its
   * place, and its paragraph can change.
   */
  readonly linear: string | null;
}
