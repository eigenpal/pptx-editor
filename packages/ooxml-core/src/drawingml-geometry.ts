// DrawingML geometry evaluation, format-neutral: guide formulas (ECMA-376 Part 1, 20.1.9.11),
// path commands (20.1.9.15 to 20.1.9.20), the text rectangle, the preset table, and the
// `a:custGeom` reader. Presets and custom geometry share one evaluator, so a preset is drawn
// from the spec's own formulas and never from a hand-drawn approximation. It reads element
// names and attribute strings only, through `DrawingGeometryElement`, so it works over any XML
// tree.
//
// Output coordinates are shape-local EMU: (0,0) is the shape's top-left, (w,h) its extent.
// Arcs become cubic Beziers, which are exact for any scale, so painters draw curves as curves.
//
// This is PPTX Editor's module (from `packages/core/src/layout/geometry/geometry.ts`), not a
// copy of a DOCX Editor file. DOCX Editor evaluates the same definitions in
// docx-editor/packages/core/src/store/package/drawing-geometry-engine.ts and
// drawing-geometry-formula.ts at 5993bd223, and flattens every curve to a polyline. Fixes for
// the DOCX Editor team (this module):
// - Only `a:moveTo` sets the current point. A path that starts with `a:lnTo`, `a:arcTo`,
//   `a:quadBezTo`, or `a:cubicBezTo` has no current point, so that command draws nothing, and
//   neither the fill nor the outline shows it (ECMA-376 Part 1, 20.1.9, `CT_Path2D`). The
//   reference PDF of the `custom-geometry` fixture, slide 2, shows no part of such a path.
//   drawing-geometry-engine.ts starts such a path at (0,0) and draws it.
// - An `a:arcTo` sweep beyond one turn redraws the same ellipse. Here it is reduced to one turn
//   plus the remainder, so the drawing and the end point stay the same with at most 9 Beziers.
//   drawing-geometry-engine.ts divides the whole sweep into at most 64 chords, so a sweep of
//   many turns draws a coarse polygon instead of an ellipse.
// drawing-geometry-engine.ts has these rules that this module does not have. Only the count
// caps can change a valid file; the schema sets no count limit:
// - a guide with an unknown operator, a path verb outside the schema, or more than 256 guides,
//   64 paths, or 2048 commands refuses the whole custom geometry, where this module reads an
//   unknown operator as 0, skips the verb, and caps the counts as below;
// - an operand literal must be an integer of at most 15 digits, where this module reads any
//   finite JavaScript number;
// - a text rectangle reference that is not a guide reads as 0, where this module reads the
//   whole shape's edge for it.
//
// Names: `customGeometryDefinition` and `evaluateGeometry` (a definition at a size, with adjust
// values) are the DOCX Editor names. `PresetShapeDefinition` is DOCX Editor's
// `GeometryDefinition` with other field names, and `evaluatePresetGeometry` adds the mirrored
// presets and the rectangle for an unknown name.
//
// Resource limits (ECMA-376 leaves them to the consumer): at most 4096 guides for each
// definition, 100 000 path commands read for each path, and 100 000 output commands for each
// evaluated geometry. Every guide value is finite.

import { PRESET_SHAPES } from './drawingml-preset-geometry.generated.ts';
import type {
  PresetCommand,
  PresetPath,
  PresetShapeDefinition,
} from './drawingml-preset-geometry-types.ts';

export type {
  PresetCommand,
  PresetPath,
  PresetShapeDefinition,
} from './drawingml-preset-geometry-types.ts';

/** One evaluated path command, absolute shape-local EMU. Arcs are already cubic Béziers. */
export type PathCommand =
  | { readonly op: 'M' | 'L'; readonly x: number; readonly y: number }
  | {
      readonly op: 'C';
      readonly x1: number;
      readonly y1: number;
      readonly x2: number;
      readonly y2: number;
      readonly x: number;
      readonly y: number;
    }
  | {
      readonly op: 'Q';
      readonly x1: number;
      readonly y1: number;
      readonly x: number;
      readonly y: number;
    }
  | { readonly op: 'Z' };

/** One evaluated path with its fill mode and stroke flag. */
export interface GeometryPath {
  readonly commands: readonly PathCommand[];
  readonly fill: 'none' | 'norm' | 'lighten' | 'lightenLess' | 'darken' | 'darkenLess';
  readonly stroke: boolean;
}

/** Evaluated geometry: paths and the text rectangle (shape-local EMU). */
export interface EvaluatedGeometry {
  readonly paths: readonly GeometryPath[];
  readonly textRect: {
    readonly l: number;
    readonly t: number;
    readonly r: number;
    readonly b: number;
  };
  /** True when the preset name is unknown and a rectangle was substituted. */
  readonly unknownPreset?: string;
}

/**
 * One element of an `a:custGeom` subtree, as the caller's XML tree holds it. The reader asks
 * only for the namespace, the local name, unprefixed attributes, and the child elements.
 */
export interface DrawingGeometryElement {
  readonly namespaceUri: string;
  readonly localName: string;
  /** The value of an unprefixed attribute (`name`, `fmla`, `w`, `x`), or `undefined`. */
  attribute(name: string): string | undefined;
  /** The child elements in document order, in every namespace. */
  childElements(): readonly DrawingGeometryElement[];
}

/** The `ST_PathFillMode` values (ECMA-376 Part 1, `CT_Path2D/@fill`); `norm` is the default. */
const PATH_FILL_MODES: ReadonlySet<string> = new Set([
  'none',
  'norm',
  'lighten',
  'lightenLess',
  'darken',
  'darkenLess',
]);

const DRAWINGML_MAIN_NAMESPACE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const ANGLE = 60_000; // units per degree
const DEG = Math.PI / 180;

function builtins(w: number, h: number): Map<string, number> {
  const ss = Math.min(w, h);
  const ls = Math.max(w, h);
  const m = new Map<string, number>([
    ['w', w],
    ['h', h],
    ['l', 0],
    ['t', 0],
    ['r', w],
    ['b', h],
    ['hc', w / 2],
    ['vc', h / 2],
    ['ss', ss],
    ['ls', ls],
    ['cd2', 10_800_000],
    ['cd4', 5_400_000],
    ['cd8', 2_700_000],
    ['3cd4', 16_200_000],
    ['3cd8', 8_100_000],
    ['5cd8', 13_500_000],
    ['7cd8', 18_900_000],
  ]);
  for (const n of [2, 3, 4, 5, 6, 8, 10, 12, 16, 32]) {
    m.set(`wd${n}`, w / n);
    m.set(`hd${n}`, h / n);
    m.set(`ssd${n}`, ss / n);
  }
  return m;
}

/** A guide formula split once into its operator and operand tokens. */
interface ParsedFormula {
  readonly op: string;
  readonly args: readonly string[];
}

/** A guide: its name and parsed formula. */
interface ParsedGuide extends ParsedFormula {
  readonly name: string;
}

function parseFormula(fmla: string): ParsedFormula {
  const parts = fmla.trim().split(/\s+/);
  return { op: parts[0] ?? '', args: parts.slice(1, 4) };
}

/** An operand: a guide defined so far, or a literal. Unknown or non-finite values are 0. */
function operand(token: string | undefined, vars: ReadonlyMap<string, number>): number {
  if (token === undefined) return 0;
  const v = vars.get(token);
  if (v !== undefined) return v;
  const n = Number(token);
  return Number.isFinite(n) ? n : 0;
}

/** Evaluate one guide formula against the variables defined so far. */
export function evaluateFormula(fmla: string, vars: ReadonlyMap<string, number>): number {
  return applyFormula(parseFormula(fmla), vars);
}

/**
 * Apply a parsed formula. The result is always finite: a formula that overflows (a product
 * of two 1e300 literals, for example) gives 0, so no later step sees an infinite value.
 */
function applyFormula(f: ParsedFormula, vars: ReadonlyMap<string, number>): number {
  const r = formulaValue(f.op, operand(f.args[0], vars), f.args, vars);
  return Number.isFinite(r) ? r : 0;
}

function formulaValue(
  op: string,
  x: number,
  args: readonly string[],
  vars: ReadonlyMap<string, number>
): number {
  const y = operand(args[1], vars);
  const z = operand(args[2], vars);
  switch (op) {
    case 'val':
      return x;
    case '*/':
      return z === 0 ? 0 : (x * y) / z;
    case '+-':
      return x + y - z;
    case '+/':
      return z === 0 ? 0 : (x + y) / z;
    case '?:':
      return x > 0 ? y : z;
    case 'abs':
      return Math.abs(x);
    case 'at2':
      return (Math.atan2(y, x) / DEG) * ANGLE;
    case 'cat2':
      return x * Math.cos(Math.atan2(z, y));
    case 'sat2':
      return x * Math.sin(Math.atan2(z, y));
    case 'cos':
      return x * Math.cos((y / ANGLE) * DEG);
    case 'sin':
      return x * Math.sin((y / ANGLE) * DEG);
    case 'tan':
      return x * Math.tan((y / ANGLE) * DEG);
    case 'max':
      return Math.max(x, y);
    case 'min':
      return Math.min(x, y);
    case 'mod':
      return Math.sqrt(x * x + y * y + z * z);
    case 'pin':
      return y < x ? x : y > z ? z : y;
    case 'sqrt':
      return Math.sqrt(Math.max(0, x));
  }
  return 0;
}

const MAX_GUIDES = 4096;
/** Input path commands read per path. */
const MAX_COMMANDS = 100_000;
/** Output commands per evaluated geometry, all paths together. */
const MAX_GEOMETRY_COMMANDS = 100_000;
/** One full turn in `ST_Angle` units (60 000ths of a degree). */
const FULL_TURN = 21_600_000;

/** Guide lists parsed once per definition (presets are static, custom ones are cached). */
const parsedGuides = new WeakMap<readonly (readonly [string, string])[], ParsedGuide[]>();

function parseGuides(guides: readonly (readonly [string, string])[]): ParsedGuide[] {
  let parsed = parsedGuides.get(guides);
  if (!parsed) {
    parsed = [];
    for (const [name, fmla] of guides) {
      if (parsed.length >= MAX_GUIDES) break;
      parsed.push({ name, ...parseFormula(fmla) });
    }
    parsedGuides.set(guides, parsed);
  }
  return parsed;
}

function defineGuides(
  vars: Map<string, number>,
  guides: readonly (readonly [string, string])[],
  overrides?: ReadonlyMap<string, number>
): void {
  for (const g of parseGuides(guides)) {
    const over = overrides?.get(g.name);
    vars.set(
      g.name,
      over !== undefined ? (Number.isFinite(over) ? over : 0) : applyFormula(g, vars)
    );
  }
}

/**
 * The sweep an arc draws, in degrees. A sweep beyond one turn redraws the same ellipse, so it
 * is reduced to one full turn plus the remainder: the drawing and the end point stay the same,
 * and the Bézier count stays bounded (at most 9) whatever a guide formula returns.
 */
function boundedSweep(swAng: number): number {
  if (!Number.isFinite(swAng)) return 0;
  const turn = FULL_TURN / ANGLE;
  const size = Math.abs(swAng);
  return size <= turn ? swAng : Math.sign(swAng) * (turn + (size % turn));
}

/** Where path output goes: the shape-local scale and the shared command budget. */
interface PathSink {
  readonly out: PathCommand[];
  readonly sx: number;
  readonly sy: number;
  /** Commands still allowed for the whole geometry. */
  budget: { left: number };
}

function emit(sink: PathSink, cmd: PathCommand): boolean {
  if (sink.budget.left <= 0) return false;
  sink.budget.left -= 1;
  sink.out.push(cmd);
  return true;
}

/**
 * Convert an `arcTo` from the current point into cubic Bézier segments. Everything is in the
 * path's own coordinate space (ECMA-376 Part 1 §20.1.9.3), and the output points are scaled
 * to the shape afterwards, so a non-uniform scale never changes the arc's angles. `stAng` and
 * `swAng` are visual angles (a ray from the ellipse center), converted to the parametric angle.
 */
function arcToBeziers(
  x0: number,
  y0: number,
  wR: number,
  hR: number,
  stAng: number,
  swAng: number,
  sink: PathSink
): { x: number; y: number } {
  if (wR === 0 || hR === 0) return { x: x0, y: y0 };
  const sw = boundedSweep(swAng);
  const param = (a: number): number => Math.atan2(wR * Math.sin(a), hR * Math.cos(a));
  const st = stAng * DEG;
  const t0 = param(st);
  // Sweep in parametric space follows the visual sweep's direction and turn count.
  const endVisual = st + sw * DEG;
  let t1 = param(endVisual);
  const turns = Math.trunc(sw / 360);
  t1 += turns * 2 * Math.PI;
  let sweep = t1 - t0;
  if (sw > 0 && sweep <= 0) sweep += 2 * Math.PI;
  if (sw < 0 && sweep >= 0) sweep -= 2 * Math.PI;
  if (sw === 0) sweep = 0;
  const cx = x0 - wR * Math.cos(t0);
  const cy = y0 - hR * Math.sin(t0);
  const segments = Math.min(16, Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2))));
  const step = sweep / segments;
  const k = (4 / 3) * Math.tan(step / 4);
  const { sx, sy } = sink;
  let a = t0;
  let px = x0;
  let py = y0;
  for (let i = 0; i < segments; i += 1) {
    const b = a + step;
    const x = cx + wR * Math.cos(b);
    const y = cy + hR * Math.sin(b);
    const pushed = emit(sink, {
      op: 'C',
      x1: (px - k * wR * Math.sin(a)) * sx,
      y1: (py + k * hR * Math.cos(a)) * sy,
      x2: (x + k * wR * Math.sin(b)) * sx,
      y2: (y - k * hR * Math.cos(b)) * sy,
      x: x * sx,
      y: y * sy,
    });
    if (!pushed) break;
    px = x;
    py = y;
    a = b;
  }
  return { x: px, y: py };
}

function evaluatePath(
  path: PresetPath,
  vars: ReadonlyMap<string, number>,
  w: number,
  h: number,
  budget: { left: number }
): GeometryPath {
  const sink: PathSink = {
    out: [],
    sx: path.w ? w / path.w : 1,
    sy: path.h ? h / path.h : 1,
    budget,
  };
  const { sx, sy } = sink;
  const v = (token: string): number => operand(token, vars);
  // The current point and the subpath start, in path units.
  let cur = { x: 0, y: 0 };
  let start = cur;
  /**
   * True once a `moveTo` has set the current point. `lnTo`, `arcTo`, `quadBezTo`, and
   * `cubicBezTo` draw from the current point (ECMA-376 Part 1, 20.1.9, `CT_Path2D`). The schema
   * makes the commands an unordered choice, so a path can start with one of them, but then it
   * has no current point. Such a command draws nothing, and neither the fill nor the
   * outline shows it: the reference PDF of fixture `custom-geometry` slide 2 shows no part of a
   * path that starts with `lnTo`.
   */
  let open = false;
  const cmds = path.cmds;
  const n = Math.min(cmds.length, MAX_COMMANDS);
  for (let i = 0; i < n && budget.left > 0; i += 1) {
    const cmd = cmds[i] as PresetCommand;
    switch (cmd[0]) {
      case 'M':
        cur = { x: v(cmd[1]), y: v(cmd[2]) };
        start = cur;
        open = true;
        emit(sink, { op: 'M', x: cur.x * sx, y: cur.y * sy });
        break;
      case 'L':
        if (!open) break;
        cur = { x: v(cmd[1]), y: v(cmd[2]) };
        emit(sink, { op: 'L', x: cur.x * sx, y: cur.y * sy });
        break;
      case 'A':
        if (!open) break;
        cur = arcToBeziers(
          cur.x,
          cur.y,
          v(cmd[1]),
          v(cmd[2]),
          v(cmd[3]) / ANGLE,
          v(cmd[4]) / ANGLE,
          sink
        );
        break;
      case 'Q': {
        if (!open) break;
        const x = v(cmd[3]);
        const y = v(cmd[4]);
        emit(sink, { op: 'Q', x1: v(cmd[1]) * sx, y1: v(cmd[2]) * sy, x: x * sx, y: y * sy });
        cur = { x, y };
        break;
      }
      case 'C': {
        if (!open) break;
        const x = v(cmd[5]);
        const y = v(cmd[6]);
        emit(sink, {
          op: 'C',
          x1: v(cmd[1]) * sx,
          y1: v(cmd[2]) * sy,
          x2: v(cmd[3]) * sx,
          y2: v(cmd[4]) * sy,
          x: x * sx,
          y: y * sy,
        });
        cur = { x, y };
        break;
      }
      case 'Z':
        // A close with nothing drawn has no subpath to close.
        if (open) emit(sink, { op: 'Z' });
        cur = start;
        break;
    }
  }
  return { commands: sink.out, fill: path.fill ?? 'norm', stroke: path.stroke !== false };
}

/**
 * Evaluate a geometry definition at a size (EMU). `adjust` replaces adjust values (`avLst`) by
 * name, as `a:prstGeom/a:avLst` does; a value that is not finite reads as 0.
 */
export function evaluateGeometry(
  def: PresetShapeDefinition,
  w: number,
  h: number,
  adjust?: ReadonlyMap<string, number>
): EvaluatedGeometry {
  const vars = builtins(w, h);
  defineGuides(vars, def.av, adjust);
  defineGuides(vars, def.gd);
  const budget = { left: MAX_GEOMETRY_COMMANDS };
  const paths = def.paths.map((p) => evaluatePath(p, vars, w, h, budget));
  const rv = (t: string, d: number): number =>
    vars.get(t) ?? (Number.isFinite(Number(t)) ? Number(t) : d);
  const textRect = def.rect
    ? { l: rv(def.rect[0], 0), t: rv(def.rect[1], 0), r: rv(def.rect[2], w), b: rv(def.rect[3], h) }
    : { l: 0, t: 0, r: w, b: h };
  return { paths, textRect };
}

/**
 * Presets missing from the reference `presetShapeDefinitions.xml`, drawn as the vertical mirror
 * of their counterpart. The file lists `upDownArrow` twice and has no `upArrow`; `upArrow` is
 * `downArrow` mirrored, with the same adjust values (ECMA-376 Part 1 §20.1.10.56).
 */
const MIRRORED_PRESETS: Readonly<Record<string, string>> = { upArrow: 'downArrow' };

function mirrorVertically(g: EvaluatedGeometry, h: number): EvaluatedGeometry {
  const y = (v: number): number => h - v;
  return {
    ...g,
    paths: g.paths.map((p) => ({
      ...p,
      commands: p.commands.map((c) =>
        c.op === 'Z'
          ? c
          : c.op === 'C'
            ? { ...c, y1: y(c.y1), y2: y(c.y2), y: y(c.y) }
            : c.op === 'Q'
              ? { ...c, y1: y(c.y1), y: y(c.y) }
              : { ...c, y: y(c.y) }
      ),
    })),
    textRect: { ...g.textRect, t: y(g.textRect.b), b: y(g.textRect.t) },
  };
}

const own = (record: object, key: string): boolean =>
  Object.prototype.hasOwnProperty.call(record, key);

/**
 * Evaluate the preset named by `a:prstGeom/@prst` at a size (EMU), with its adjust values
 * (ECMA-376 Part 1, 20.1.9.18 `prstGeom`). An unknown name draws a rectangle and reports the
 * name in `unknownPreset`.
 */
export function evaluatePresetGeometry(
  name: string,
  w: number,
  h: number,
  adjust: ReadonlyMap<string, number>
): EvaluatedGeometry {
  const mirror = own(MIRRORED_PRESETS, name) ? MIRRORED_PRESETS[name] : undefined;
  if (mirror) return mirrorVertically(evaluateGeometry(PRESET_SHAPES[mirror], w, h, adjust), h);
  const def = own(PRESET_SHAPES, name) ? PRESET_SHAPES[name] : undefined;
  if (!def)
    return {
      ...evaluateGeometry(PRESET_SHAPES.rect, w, h, new Map()),
      unknownPreset: name,
    };
  return evaluateGeometry(def, w, h, adjust);
}

/** Every preset name the evaluator can draw. */
export function presetNames(): string[] {
  return [...Object.keys(PRESET_SHAPES), ...Object.keys(MIRRORED_PRESETS)];
}

/** The first child element in the DrawingML main namespace with this local name. */
function drawingChild(
  el: DrawingGeometryElement | undefined,
  localName: string
): DrawingGeometryElement | undefined {
  if (!el) return undefined;
  for (const c of el.childElements())
    if (c.namespaceUri === DRAWINGML_MAIN_NAMESPACE_URI && c.localName === localName) return c;
  return undefined;
}

/** Every child element in the DrawingML main namespace with this local name. */
function drawingChildren(
  el: DrawingGeometryElement | undefined,
  localName: string
): DrawingGeometryElement[] {
  if (!el) return [];
  const out: DrawingGeometryElement[] = [];
  for (const c of el.childElements())
    if (c.namespaceUri === DRAWINGML_MAIN_NAMESPACE_URI && c.localName === localName) out.push(c);
  return out;
}

/** A `ST_PositiveCoordinate` path size; anything that is not an integer reads as absent. */
function pathSize(raw: string | undefined): number | undefined {
  if (raw === undefined || !/^-?\d{1,15}$/.test(raw.trim())) return undefined;
  const v = Number(raw);
  return Number.isSafeInteger(v) ? v : undefined;
}

/**
 * Read an `a:custGeom` element (ECMA-376 Part 1, 20.1.9.8 `CT_CustomGeometry2D`) into the
 * definition shape that the presets use. A missing guide formula reads as `val 0`, and a
 * missing point coordinate as 0.
 */
export function customGeometryDefinition(cust: DrawingGeometryElement): PresetShapeDefinition {
  const guides = (list: DrawingGeometryElement | undefined): [string, string][] =>
    drawingChildren(list, 'gd').map((g) => [
      g.attribute('name') ?? '',
      g.attribute('fmla') ?? 'val 0',
    ]);
  const rect = drawingChild(cust, 'rect');
  const paths: PresetPath[] = drawingChildren(drawingChild(cust, 'pathLst'), 'path').map((p) => {
    const cmds: PresetCommand[] = [];
    for (const c of p.childElements()) {
      const pts = drawingChildren(c, 'pt').map(
        (pt) => [pt.attribute('x') ?? '0', pt.attribute('y') ?? '0'] as const
      );
      switch (c.localName) {
        case 'moveTo':
          if (pts[0]) cmds.push(['M', pts[0][0], pts[0][1]]);
          break;
        case 'lnTo':
          if (pts[0]) cmds.push(['L', pts[0][0], pts[0][1]]);
          break;
        case 'arcTo':
          cmds.push([
            'A',
            c.attribute('wR') ?? '0',
            c.attribute('hR') ?? '0',
            c.attribute('stAng') ?? '0',
            c.attribute('swAng') ?? '0',
          ]);
          break;
        case 'quadBezTo':
          if (pts[1]) cmds.push(['Q', pts[0][0], pts[0][1], pts[1][0], pts[1][1]]);
          break;
        case 'cubicBezTo':
          if (pts[2])
            cmds.push(['C', pts[0][0], pts[0][1], pts[1][0], pts[1][1], pts[2][0], pts[2][1]]);
          break;
        case 'close':
          cmds.push(['Z']);
          break;
      }
    }
    const fill = p.attribute('fill');
    const stroke = p.attribute('stroke');
    const w = pathSize(p.attribute('w'));
    const h = pathSize(p.attribute('h'));
    return {
      ...(w ? { w } : {}),
      ...(h ? { h } : {}),
      // A fill mode outside `ST_PathFillMode` reads as the default, `norm`.
      ...(fill && PATH_FILL_MODES.has(fill) ? { fill: fill as PresetPath['fill'] } : {}),
      ...(stroke === '0' || stroke === 'false' ? { stroke: false as const } : {}),
      cmds,
    };
  });
  return {
    av: guides(drawingChild(cust, 'avLst')),
    gd: guides(drawingChild(cust, 'gdLst')),
    ...(rect
      ? {
          rect: [
            rect.attribute('l') ?? 'l',
            rect.attribute('t') ?? 't',
            rect.attribute('r') ?? 'r',
            rect.attribute('b') ?? 'b',
          ] as const,
        }
      : {}),
    paths,
  };
}

/** One adjust value of a preset (`a:avLst/a:gd`), with the range of the handle that moves it. */
export interface PresetAdjustment {
  readonly name: string;
  /** The value in guide units: 1/100 000 for a length, 1/60 000 of a degree for an angle. */
  readonly value: number;
  /** An angle (`a:ahPolar/@gdRefAng`). */
  readonly angle: boolean;
  /** The handle range at this size, or null when the preset gives none. */
  readonly min: number | null;
  readonly max: number | null;
}

/**
 * The adjust values of a preset in `a:avLst` order, each with the range of its adjust handle
 * (ECMA-376 Part 1, 20.1.9.1 ahLst), evaluated at the size `w` by `h` (EMU) with the shape's
 * own values. A range that names a guide needs a size: at zero size only literal ranges apply.
 * Undefined for a preset that the engine does not know.
 */
export function presetAdjustments(
  preset: string,
  w: number,
  h: number,
  adjust: ReadonlyMap<string, number>
): PresetAdjustment[] | undefined {
  const name = Object.prototype.hasOwnProperty.call(MIRRORED_PRESETS, preset)
    ? MIRRORED_PRESETS[preset]!
    : preset;
  if (!Object.prototype.hasOwnProperty.call(PRESET_SHAPES, name)) return undefined;
  const def = PRESET_SHAPES[name]!;
  const sized = w > 0 && h > 0;
  const vars = builtins(sized ? w : 0, sized ? h : 0);
  defineGuides(vars, def.av, adjust);
  defineGuides(vars, def.gd);
  const literal = (token: string | null): number | null => {
    if (token === null) return null;
    const n = Number(token);
    if (Number.isFinite(n)) return n;
    return sized && vars.has(token) ? vars.get(token)! : null;
  };
  return def.av.map(([guide]) => {
    const handle = def.ah?.find((a) => a[0] === guide);
    return {
      name: guide,
      value: vars.get(guide) ?? 0,
      angle: handle?.[1] === 'ang',
      min: handle ? literal(handle[2]) : null,
      max: handle ? literal(handle[3]) : null,
    };
  });
}

/**
 * The guide values of a preset at a size, with its definition, for the adjust handles
 * (`drawingml-handles.ts`). `mirrored`: the preset draws its base preset upside down.
 * Undefined for a preset that the engine does not know. @internal
 */
export function presetGuides(
  preset: string,
  w: number,
  h: number,
  adjust: ReadonlyMap<string, number>
):
  | {
      readonly def: PresetShapeDefinition;
      readonly vars: Map<string, number>;
      readonly mirrored: boolean;
    }
  | undefined {
  const mirrored = Object.prototype.hasOwnProperty.call(MIRRORED_PRESETS, preset);
  const name = mirrored ? MIRRORED_PRESETS[preset]! : preset;
  if (!Object.prototype.hasOwnProperty.call(PRESET_SHAPES, name)) return undefined;
  const def = PRESET_SHAPES[name]!;
  const vars = builtins(w, h);
  defineGuides(vars, def.av, adjust);
  defineGuides(vars, def.gd);
  return { def, vars, mirrored };
}
