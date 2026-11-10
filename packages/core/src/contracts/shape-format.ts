// The Shape Format types of the public contract: effect presets, Quick Styles cells, and the
// selection pane's list. Split from `index.ts`, which re-exports them, for the line cap.

import type { ColorChoice } from './neutral.ts';

/** The axis of a flip or a distribute: `horizontal` acts along x. @public */
export type ObjectAxis = 'horizontal' | 'vertical';

/**
 * The presets of the Shadow gallery: the nine Outer shadows ("Offset: Bottom Right")
 * and the nine Inner shadows ("Inside: Top Left"). @public
 */
export type ShadowPreset =
  | 'outerBottomRight'
  | 'outerBottom'
  | 'outerBottomLeft'
  | 'outerRight'
  | 'outerCenter'
  | 'outerLeft'
  | 'outerTopRight'
  | 'outerTop'
  | 'outerTopLeft'
  | 'innerTopLeft'
  | 'innerTop'
  | 'innerTopRight'
  | 'innerLeft'
  | 'innerCenter'
  | 'innerRight'
  | 'innerBottomLeft'
  | 'innerBottom'
  | 'innerBottomRight';

/** The presets of the Reflection gallery: size, and the gap to the shape. @public */
export type ReflectionPreset =
  | 'tightTouching'
  | 'halfTouching'
  | 'fullTouching'
  | 'tight4pt'
  | 'half4pt'
  | 'full4pt'
  | 'tight8pt'
  | 'half8pt'
  | 'full8pt';

/**
 * The presets of the Bevel gallery: the top bevel presets of `ST_BevelPresetType`
 * (ECMA-376 Part 1, 20.1.10.9), each 6 pt wide and 6 pt high. @public
 */
export type BevelPreset =
  | 'circle'
  | 'relaxedInset'
  | 'cross'
  | 'coolSlant'
  | 'angle'
  | 'softRound'
  | 'convex'
  | 'slope'
  | 'divot'
  | 'riblet'
  | 'hardEdge'
  | 'artDeco';

/**
 * The presets of the 3-D Rotation gallery: camera presets of `ST_PresetCameraType`
 * (ECMA-376 Part 1, 20.1.10.47) with parallel projection. @public
 */
export type Rotation3DPreset =
  | 'isometricLeftDown'
  | 'isometricRightUp'
  | 'isometricTopUp'
  | 'isometricOffAxis1Left'
  | 'isometricOffAxis1Right'
  | 'obliqueTopLeft'
  | 'obliqueTopRight'
  | 'obliqueBottomLeft'
  | 'obliqueBottomRight';

/** One effect to set; a null preset or size removes the effect. @public */
export type ShapeEffectChoice =
  | { readonly kind: 'shadow'; readonly preset: ShadowPreset | null }
  | { readonly kind: 'reflection'; readonly preset: ReflectionPreset | null }
  | {
      readonly kind: 'glow';
      /** Glow size in points (the gallery offers 5, 8, 11, and 18). */
      readonly size: number | null;
      /** The glow color; absent keeps the current one, else Accent 1. */
      readonly color?: ColorChoice;
    }
  /** Soft edge size in points (the gallery offers 1, 2.5, 5, 10, 25, and 50). */
  | { readonly kind: 'softEdge'; readonly size: number | null }
  /** Blur radius in points (`a:blur`). No gallery offers it. */
  | { readonly kind: 'blur'; readonly size: number | null }
  /** A top bevel (`a:sp3d/a:bevelT`), with a front camera when the shape has no scene. */
  | { readonly kind: 'bevel'; readonly preset: BevelPreset | null }
  /** A camera preset (`a:scene3d/a:camera`); null turns the camera back to the front. */
  | { readonly kind: 'rotation3d'; readonly preset: Rotation3DPreset | null };

/**
 * The rows of the Shape Styles gallery (theme styles), top to bottom: Colored Outline,
 * Colored Fill, Light 1 Outline Colored Fill, Subtle Effect, Moderate Effect, Intense Effect. @public
 */
export type ShapeStyleRow =
  | 'coloredOutline'
  | 'coloredFill'
  | 'lightOutline'
  | 'subtleEffect'
  | 'moderateEffect'
  | 'intenseEffect';

/** The columns of the Shape Styles gallery: Dark 1, then Accent 1 to 6 (theme color names). @public */
export type ShapeStyleColor =
  'dk1' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6';

/** A `shape.styles` value: a gallery cell as `row.color` (`intenseEffect.accent2`). @public */
export type ShapeStyleValue = `${ShapeStyleRow}.${ShapeStyleColor}`;

/**
 * A `shape.effects` option value: an effect and its choice (`shadow.outerTop`, `glow.8.accent2`,
 * `softEdge.none`). A glow without a color keeps the shape's glow color. Sizes are points
 * above 0; the command refuses others. The slot's state value lists the current choice of each
 * effect, separated by spaces: test an option with `chromeOptionChecked`. @public
 */
export type ShapeEffectsValue =
  | `shadow.${ShadowPreset | 'none'}`
  | `reflection.${ReflectionPreset | 'none'}`
  | `glow.${number | 'none'}`
  | `glow.${number}.${Exclude<ShapeStyleColor, 'dk1'>}`
  | `softEdge.${number | 'none'}`
  | `bevel.${BevelPreset | 'none'}`
  | `rotation3d.${Rotation3DPreset | 'none'}`;

/**
 * A preset text warp (`a:prstTxWarp/@prst`, `ST_TextShapeType`, ECMA-376 Part 1, 20.1.10.76).
 * `textNoShape` draws the text flat. @public
 */
export type TextWarpPreset =
  | 'textNoShape'
  | 'textPlain'
  | 'textStop'
  | 'textTriangle'
  | 'textTriangleInverted'
  | 'textChevron'
  | 'textChevronInverted'
  | 'textRingInside'
  | 'textRingOutside'
  | 'textArchUp'
  | 'textArchDown'
  | 'textCircle'
  | 'textButton'
  | 'textArchUpPour'
  | 'textArchDownPour'
  | 'textCirclePour'
  | 'textButtonPour'
  | 'textCurveUp'
  | 'textCurveDown'
  | 'textCanUp'
  | 'textCanDown'
  | 'textWave1'
  | 'textWave2'
  | 'textDoubleWave1'
  | 'textWave4'
  | 'textInflate'
  | 'textDeflate'
  | 'textInflateBottom'
  | 'textDeflateBottom'
  | 'textInflateTop'
  | 'textDeflateTop'
  | 'textDeflateInflate'
  | 'textDeflateInflateDeflate'
  | 'textFadeRight'
  | 'textFadeLeft'
  | 'textFadeUp'
  | 'textFadeDown'
  | 'textSlantUp'
  | 'textSlantDown'
  | 'textCascadeUp'
  | 'textCascadeDown';

/**
 * A `shape.textEffects` option value: `transform.` and a text warp, or `transform.none` to
 * remove the warp. @public
 */
export type TextEffectsValue = `transform.${Exclude<TextWarpPreset, 'textNoShape'> | 'none'}`;

/** An `arrange.rotate` value: the 90° turns, then the flips. @public */
export type RotateChoice = 'right90' | 'left90' | 'flipVertical' | 'flipHorizontal';

/** One cell of the Shape Styles gallery. @public */
export interface ShapeStyleChoice {
  readonly row: ShapeStyleRow;
  readonly color: ShapeStyleColor;
}

/**
 * The effects of the selected element as the chrome shows them. A shadow or reflection that
 * is not a gallery preset reads `other`. Sizes are in points; 0 is no effect. @public
 */
export interface ShapeEffectsState {
  readonly shadow: ShadowPreset | 'none' | 'other';
  readonly reflection: ReflectionPreset | 'none' | 'other';
  readonly glow: number;
  /** The glow color as `#RRGGBB`, when there is a glow. */
  readonly glowColor?: string;
  /** The theme accent of a Glow Variations color (40% opacity), when the glow has one. */
  readonly glowAccent?: Exclude<ShapeStyleColor, 'dk1'>;
  readonly softEdge: number;
  /** The top bevel preset, or `none`. */
  readonly bevel: BevelPreset | 'none';
  /** The camera preset when it is a 3-D Rotation gallery choice, `none` for a front camera. */
  readonly rotation3d: Rotation3DPreset | 'none' | 'other';
  /** Blur radius in points; 0 is no blur. */
  readonly blur: number;
}

/** One element of the current slide, as the selection pane lists it. @public */
export interface ShapeSummary {
  readonly nodeId: string;
  /** `p:cNvPr/@name`. */
  readonly name: string;
  readonly kind: 'shape' | 'picture' | 'group' | 'connector' | 'graphicFrame';
  /** Whether the element supports start and end arrowhead types. */
  readonly lineEnds?: boolean;
  /** `p:cNvPr/@hidden`: the element does not draw. */
  readonly hidden: boolean;
  /** A group's children, top first. */
  readonly children?: readonly ShapeSummary[];
}
