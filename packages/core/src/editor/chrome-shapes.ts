// The Shape Format chrome: Quick Styles (the theme styles gallery), Shape Effects (one menu with
// Shadow, Reflection, Glow, and Soft Edges, on Shape Format), and
// Format Shape in the `shape` group, and the Arrange rows that rotate, flip, distribute, and open
// the selection pane. Slot names follow the Names rule of `editor-ui-parity`: an existing PPTX
// group for the idea (`shape.*`, `arrange.*`), never a ribbon tab or group name.

import type { TranslationKey } from '@pptx-editor.dev/i18n';
import type {
  EditorCommand,
  EditorSnapshot,
  ReflectionPreset,
  RotateChoice,
  ShadowPreset,
  ShapeEffectChoice,
  ShapeStyleChoice,
  ShapeStyleColor,
  ShapeStyleRow,
} from '../contracts/index.ts';
import type { ChromeControl, ChromeSlotId } from './chrome-controls.ts';
import {
  GLOW_SIZES,
  REFLECTION_PRESETS,
  SHADOW_PRESETS,
  SHAPE_STYLE_COLORS,
  SHAPE_STYLE_ROWS,
  SOFT_EDGE_SIZES,
} from '../automation/commands/effect-commands.ts';
import { ICON_PATHS } from './icon-paths.ts';
import { applyTransforms, parseHex, toHex } from '../cascade/index.ts';
import { NS, type Element } from '../store/index.ts';

const icon = (name: string): readonly string[] => ICON_PATHS[name] ?? [];

const SHADOWS = Object.keys(SHADOW_PRESETS) as ShadowPreset[];
const REFLECTIONS = Object.keys(REFLECTION_PRESETS) as ReflectionPreset[];

/** The Align menu rows that distribute (after the six alignments). */
export const DISTRIBUTE_OPTIONS = [
  { value: 'distributeHorizontal', labelKey: 'toolbar.distributeHorizontal' },
  { value: 'distributeVertical', labelKey: 'toolbar.distributeVertical' },
] as const satisfies readonly { value: string; labelKey: TranslationKey }[];

/** The Rotate dropdown of the Arrange group. */
export const ROTATE_CONTROL = {
  id: 'rotate',
  shape: 'dropdown',
  labelKey: 'toolbar.rotate',
  paths: icon('rotate_right'),
  defaultToolbar: false,
  options: [
    { value: 'right90', labelKey: 'toolbar.rotations.right90', paths: icon('rotate_right') },
    { value: 'left90', labelKey: 'toolbar.rotations.left90', paths: icon('rotate_left') },
    {
      value: 'flipVertical',
      labelKey: 'toolbar.rotations.flipVertical',
      paths: icon('swap_vert'),
    },
    {
      value: 'flipHorizontal',
      labelKey: 'toolbar.rotations.flipHorizontal',
      paths: icon('swap_horiz'),
    },
  ],
  state: { kind: 'value' },
} as const satisfies ChromeControl;

/** The Selection Pane button of the Arrange group. */
export const SELECTION_PANE_CONTROL = {
  id: 'selectionPane',
  labelKey: 'toolbar.selectionPane',
  paths: icon('view_sidebar'),
  defaultToolbar: false,
  state: { kind: 'pane', pane: 'selection' },
} as const satisfies ChromeControl;

/** The effect kinds of the Shape Effects menu, in menu order, with their labels. */
const EFFECT_GROUPS = [
  { id: 'shadow', labelKey: 'toolbar.shadow' },
  { id: 'reflection', labelKey: 'toolbar.reflection' },
  { id: 'glow', labelKey: 'toolbar.glow' },
  { id: 'softEdge', labelKey: 'toolbar.softEdges' },
] as const satisfies readonly { id: ShapeEffectChoice['kind']; labelKey: TranslationKey }[];

const none = (group: ShapeEffectChoice['kind']) => ({
  value: `${group}.none`,
  labelKey: `toolbar.noEffect.${group}` as const,
  group,
});
const sizes = (group: 'glow' | 'softEdge', list: readonly number[], labelKey: TranslationKey) =>
  list.map((s) => ({ value: `${group}.${s}`, labelKey, labelParams: { value: s }, group }));

/** The colors of the Glow Variations: Accent 1 to 6, the columns of each size row. @internal */
export const GLOW_ACCENTS = [
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
] as const satisfies readonly ShapeStyleColor[];

/** Glow Variations: a row of the six accents for each size. */
const glowVariations = () =>
  GLOW_SIZES.flatMap((s) =>
    GLOW_ACCENTS.map((accent) => ({
      value: `glow.${s}.${accent}`,
      labelKey: `toolbar.glowVariations.${accent}` as const,
      labelParams: { value: s },
      group: 'glow',
    }))
  );

/**
 * Quick Styles: the theme styles gallery, first in the `shape` group, before Shape Fill and
 * Shape Outline. Not in the default formatting bar: the bar is full at 1440 px, and one more
 * control moves the outline controls into the More menu. The
 * ribbon and a host place it.
 */
export const QUICK_STYLES_CONTROL: ChromeControl = {
  // Six rows of seven colors. The chrome draws each cell.
  id: 'styles',
  shape: 'gallery',
  labelKey: 'toolbar.shapeStyles',
  paths: null,
  defaultToolbar: false,
  state: { kind: 'value' },
};

/**
 * Shape Effects and Format Shape: the controls the `shape` group adds after its outline
 * controls. They are not in the default formatting bar, as the Home tab does not show
 * them. A host composes them; the ribbon's Shape Format tab and the
 * menus place them in a later change.
 */
export const SHAPE_FORMAT_CONTROLS: readonly ChromeControl[] = [
  {
    // One option per gallery choice; `group` puts each in its effect's submenu. The value is
    // the chosen option of every group (`chromeOptionChecked`).
    id: 'effects',
    shape: 'dropdown',
    labelKey: 'toolbar.shapeEffects',
    paths: null,
    defaultToolbar: false,
    optionGroups: EFFECT_GROUPS,
    options: [
      none('shadow'),
      ...SHADOWS.map((s) => ({
        value: `shadow.${s}`,
        labelKey: `toolbar.shadows.${s}` as const,
        group: 'shadow',
      })),
      none('reflection'),
      ...REFLECTIONS.map((r) => ({
        value: `reflection.${r}`,
        labelKey: `toolbar.reflections.${r}` as const,
        group: 'reflection',
      })),
      none('glow'),
      ...glowVariations(),
      none('softEdge'),
      ...sizes('softEdge', SOFT_EDGE_SIZES, 'toolbar.softEdgeValue'),
    ],
    state: { kind: 'value' },
  },
  {
    // Format Shape: the Format pane, after DOCX Editor's `image.properties`.
    id: 'properties',
    labelKey: 'toolbar.formatShape',
    paths: icon('format_shapes'),
    defaultToolbar: false,
    state: { kind: 'pane', pane: 'format' },
  },
];

/** A gallery cell from a value: `{ row, color }`, or the string `row.color`. @internal */
export function shapeStyleChoice(value: unknown): ShapeStyleChoice | null {
  const pair =
    typeof value === 'string'
      ? value.split('.')
      : typeof value === 'object' && value !== null
        ? [(value as { row?: unknown }).row, (value as { color?: unknown }).color]
        : [];
  const [row, color] = pair;
  return pair.length === 2 &&
    typeof row === 'string' &&
    Object.prototype.hasOwnProperty.call(SHAPE_STYLE_ROWS, row) &&
    SHAPE_STYLE_COLORS.includes(color as ShapeStyleColor)
    ? { row: row as ShapeStyleRow, color: color as ShapeStyleColor }
    : null;
}

/** The effect command of a Shape Effects option (`shadow.outerTop`, `glow.8`, `softEdge.none`). @internal */
export function shapeEffectCommand(value: unknown): EditorCommand | null {
  if (typeof value !== 'string') return null;
  const dot = value.indexOf('.');
  const kind = value.slice(0, dot);
  const pick = value.slice(dot + 1);
  const off = pick === 'none';
  switch (kind) {
    case 'shadow':
      return off || SHADOWS.includes(pick as ShadowPreset)
        ? {
            type: 'setShapeEffect',
            effect: { kind: 'shadow', preset: off ? null : (pick as ShadowPreset) },
          }
        : null;
    case 'reflection':
      return off || REFLECTIONS.includes(pick as ReflectionPreset)
        ? {
            type: 'setShapeEffect',
            effect: { kind: 'reflection', preset: off ? null : (pick as ReflectionPreset) },
          }
        : null;
    case 'glow':
    case 'softEdge': {
      // `glow.8.accent2`: a size and, for a glow, an optional Glow Variations color.
      const [amount = '', accent, ...rest] = pick.split(/\.(?=accent)/);
      const size = off ? null : Number(amount);
      if (size !== null && !(Number.isFinite(size) && size > 0)) return null;
      if (rest.length > 0 || (accent !== undefined && (kind !== 'glow' || off))) return null;
      if (accent !== undefined && !(GLOW_ACCENTS as readonly string[]).includes(accent))
        return null;
      if (kind === 'softEdge') return { type: 'setShapeEffect', effect: { kind, size } };
      return {
        type: 'setShapeEffect',
        effect: {
          kind,
          size,
          ...(accent ? { color: { theme: accent as (typeof GLOW_ACCENTS)[number] } } : {}),
        },
      };
    }
    default:
      return null;
  }
}

/**
 * The command a Shape Format slot dispatches for a picked value, `null` for an invalid value,
 * or `undefined` for a slot this module does not own. @internal
 */
export function shapeFormatSlotCommand(
  slot: ChromeSlotId,
  value: unknown
): EditorCommand | null | undefined {
  switch (slot) {
    case 'shape.styles': {
      const style = shapeStyleChoice(value);
      return style ? { type: 'setShapeStyle', style } : null;
    }
    case 'shape.effects':
      return shapeEffectCommand(value);
    case 'arrange.rotate':
      switch (value as RotateChoice) {
        case 'right90':
          return { type: 'rotateObjects', degrees: 90 };
        case 'left90':
          return { type: 'rotateObjects', degrees: -90 };
        case 'flipVertical':
          return { type: 'flipObjects', axis: 'vertical' };
        case 'flipHorizontal':
          return { type: 'flipObjects', axis: 'horizontal' };
        default:
          return null;
      }
    default:
      return undefined;
  }
}

/** Probe values for `Editor.can`: well-formed picks, never executed. @internal */
export const SHAPE_FORMAT_PROBES: Partial<Record<ChromeSlotId, unknown>> = {
  'shape.styles': 'coloredFill.accent1',
  // A preset, not `none`: removing an effect a shape lacks is a no-op refusal.
  'shape.effects': 'shadow.outerBottomRight',
  'arrange.rotate': 'right90',
};

/**
 * The value a Shape Format slot shows: the selection's style (`row.color`), or for Shape Effects
 * the chosen option of each effect, separated by spaces (`shadow.outerTop reflection.none …`).
 * @internal
 */
export function shapeFormatSlotValue(slot: ChromeSlotId, snap: EditorSnapshot): string | undefined {
  const g = snap.geometry;
  if (slot === 'shape.styles')
    return g?.shapeStyle ? `${g.shapeStyle.row}.${g.shapeStyle.color}` : undefined;
  if (slot !== 'shape.effects' || !g?.effects) return undefined;
  const fx = g.effects;
  const size = (n: number): string => (n ? String(n) : 'none');
  return [
    `shadow.${fx.shadow}`,
    `reflection.${fx.reflection}`,
    fx.glow && fx.glowAccent ? `glow.${fx.glow}.${fx.glowAccent}` : `glow.${size(fx.glow)}`,
    `softEdge.${size(fx.softEdge)}`,
  ].join(' ');
}

/**
 * Whether a dropdown option is checked: the slot's value is the option, or, for a slot whose
 * options have groups (Shape Effects), one of its space-separated choices. Use it in a host's
 * own menu: `chromeOptionChecked(cmd.value, 'shadow.outerTop')`. @public
 */
export function chromeOptionChecked(value: unknown, option: string): boolean {
  return typeof value === 'string' && (value === option || value.split(' ').includes(option));
}

/** The Shape Styles gallery: its rows, top to bottom, each with its seven cells. @internal */
export const SHAPE_STYLE_GALLERY: readonly {
  readonly row: ShapeStyleRow;
  readonly cells: readonly ShapeStyleChoice[];
}[] = (Object.keys(SHAPE_STYLE_ROWS) as ShapeStyleRow[]).map((row) => ({
  row,
  cells: SHAPE_STYLE_COLORS.map((color) => ({ row, color })),
}));

/** How a gallery cell draws: fill, outline, and text colors as `#RRGGBB`, or null for none. @internal */
export interface ShapeStylePreview {
  readonly fill: string | null;
  readonly line: string | null;
  readonly text: string;
  /** The theme effect style index: 0 none, 1 subtle, 2 moderate, 3 intense. */
  readonly effect: number;
}

/**
 * The colors a gallery cell draws with, from the slide's theme colors (`snapshot.themeColors`,
 * first row: Background 1, Text 1, Background 2, Text 2, Accent 1–6). A preview for the
 * chrome, not the slide: theme fills draw as their base color. @internal
 */
export function shapeStylePreview(
  cell: ShapeStyleChoice,
  themeColors: EditorSnapshot['themeColors']
): ShapeStylePreview {
  const row = SHAPE_STYLE_ROWS[cell.row];
  const swatch = (slot: string): string =>
    themeColors[0]?.find((s) => s.slot === slot)?.color ?? '#808080';
  // `a:schemeClr val="dk1"` in `p:style` names the theme color itself, not a mapped slot: on a
  // dark color map (`bg1="dk1"`), Dark 1 is the Background 1 swatch.
  const direct = (scheme: 'dk1' | 'lt1', fallback: string): string =>
    themeColors[0]?.find((s) => s.scheme === scheme)?.color ?? swatch(fallback);
  const named = (name: string): string =>
    name === 'dk1' ? direct('dk1', 'tx1') : name === 'lt1' ? direct('lt1', 'bg1') : swatch(name);
  const refColor = (ref: readonly [number, string]): string | null => {
    if (ref[0] === 0) return null;
    // `a:shade val="15000"`: 15% of the color.
    if (ref[1] === 'Cshade') return shade(named(cell.color), 0.15);
    return named(ref[1] === 'C' ? cell.color : ref[1]);
  };
  const fill = refColor(row.fill);
  return {
    fill,
    line: refColor(row.ln),
    text: readableText(named(row.font), fill),
    effect: row.effect,
  };
}

/** The relative luminance of `#RRGGBB` (WCAG 2.2), or null for a color it cannot read. */
function luminance(hex: string): number | null {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex);
  if (!m) return null;
  const n = parseInt(m[1]!, 16);
  const lin = (c: number): number => {
    const v = c / 255;
    return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
  };
  return 0.2126 * lin((n >> 16) & 255) + 0.7152 * lin((n >> 8) & 255) + 0.0722 * lin(n & 255);
}

/**
 * The preview's text color: the style's own font color, unless it is unreadable on the cell's
 * fill (contrast under 3:1, as Dark 1 text on the Dark 1 base color of a Subtle Effect cell,
 * whose real fill is a light tint the preview does not draw). Then black or white, the one
 * that reads on the fill. @internal
 */
export function readableText(text: string, fill: string | null): string {
  const t = luminance(text);
  const f = fill === null ? null : luminance(fill);
  if (t === null || f === null) return text;
  const contrast = (Math.max(t, f) + 0.05) / (Math.min(t, f) + 0.05);
  if (contrast >= 3) return text;
  return f > 0.18 ? '#000000' : '#FFFFFF';
}

/** `#RRGGBB` with `a:shade` applied as the slide draws it: the cascade's transform math. */
function shade(hex: string, factor: number): string {
  const base = parseHex(hex.replace('#', ''));
  if (!base) return hex;
  const el: Element = {
    type: 'element',
    id: '',
    kind: 'generic',
    ns: NS.a,
    local: 'shade',
    nsDecls: [],
    attrs: [{ ns: '', local: 'val', value: String(Math.round(factor * 100_000)) }],
    children: [],
  };
  return toHex(applyTransforms(base, [el])).toUpperCase();
}
