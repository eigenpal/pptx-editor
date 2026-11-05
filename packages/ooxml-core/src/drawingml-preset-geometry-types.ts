// The shape of one geometry definition, as generated from ECMA-376 presetShapeDefinitions.xml
// (`CT_PresetGeometry2D` data) and as read from `a:custGeom` (`CT_CustomGeometry2D`).

/** One path command with its operands as guide names or literals. */
export type PresetCommand =
  | readonly ['M' | 'L', string, string]
  | readonly ['A', string, string, string, string]
  | readonly ['Q', string, string, string, string]
  | readonly ['C', string, string, string, string, string, string]
  | readonly ['Z'];

/** One path of a preset. `w`/`h` set the path's own coordinate space. */
export interface PresetPath {
  readonly w?: number;
  readonly h?: number;
  readonly fill?: 'none' | 'norm' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';
  readonly stroke?: false;
  readonly cmds: readonly PresetCommand[];
}

/** One preset shape: adjust defaults, guides, text rectangle, and paths. */
export interface PresetShapeDefinition {
  readonly av: readonly (readonly [string, string])[];
  readonly gd: readonly (readonly [string, string])[];
  /** Text rectangle as guide names: l, t, r, b. */
  readonly rect?: readonly [string, string, string, string];
  readonly paths: readonly PresetPath[];
  /**
   * Adjust handles (`a:ahLst`): the guide each handle moves, its axis (`x`, `y`, `r`, or `ang`
   * for an angle), and its minimum and maximum as a guide name or a literal, or null.
   */
  readonly ah?: readonly (readonly [
    string,
    'x' | 'y' | 'r' | 'ang',
    string | null,
    string | null,
  ])[];
}
