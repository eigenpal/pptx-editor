// The 3-D scene and shape properties of a shape (`a:scene3d`, ECMA-376 Part 1, 20.1.4.1.26, and
// `a:sp3d`, 20.1.5.12), and the preset shadows (`a:prstShdw`, 20.1.8.49). Values are clamped
// and checked here, at the parse boundary: the paint reads only the resolved records.

import { attr, child, intAttr, type Element } from '../store/index.ts';
import { colorChild, resolveColor, type ColorContext, type Rgba } from './color.ts';
import type { EffectPlacement } from './paint-props.ts';

/** A bevel (`a:bevelT`, `a:bevelB`): width and height in EMU, and its preset. */
export interface ResolvedBevel {
  /** Width, inward from the outline. */
  readonly w: number;
  /** Height, toward the viewer. */
  readonly h: number;
  /** `ST_BevelPresetType`; an unknown value reads `circle`, the schema default. */
  readonly preset: BevelPresetType;
}

/** The bevel presets of `ST_BevelPresetType` (ECMA-376 Part 1, 20.1.10.9). */
export const BEVEL_PRESET_TYPES = [
  'relaxedInset',
  'circle',
  'slope',
  'cross',
  'angle',
  'softRound',
  'convex',
  'coolSlant',
  'divot',
  'riblet',
  'hardEdge',
  'artDeco',
] as const;

/** One `ST_BevelPresetType` value. */
export type BevelPresetType = (typeof BEVEL_PRESET_TYPES)[number];

/** A rotation in degrees (`a:rot`: latitude, longitude, revolution). */
export interface Rotation3D {
  /** Latitude: the turn about the x axis. */
  readonly lat: number;
  /** Longitude: the turn about the y axis. */
  readonly lon: number;
  /** Revolution: the turn about the z axis. */
  readonly rev: number;
}

/** The resolved 3-D properties of a shape. Lengths in EMU. */
export interface ResolvedThreeD {
  /** `a:camera/@prst` (`ST_PresetCameraType`), as written. */
  readonly camera: string;
  /** `a:camera/a:rot`, when present: it replaces the preset's rotation. */
  readonly cameraRot?: Rotation3D;
  /** `a:lightRig/@rig` (`ST_LightRigType`) and `@dir` (`ST_LightRigDirection`). */
  readonly lightRig: string;
  readonly lightDir: string;
  /** The top bevel (`a:bevelT`), on the face. */
  readonly bevelTop?: ResolvedBevel;
  /** The bottom bevel (`a:bevelB`), on the back. */
  readonly bevelBottom?: ResolvedBevel;
  /** The extrusion depth (`@extrusionH`) and its color (`a:extrusionClr`). */
  readonly extrusionH: number;
  readonly extrusionColor?: Rgba;
  /** The contour width (`@contourW`) and its color (`a:contourClr`). */
  readonly contourW: number;
  readonly contourColor?: Rgba;
  /** `ST_PresetMaterialType`; `warmMatte` is the schema default. */
  readonly material: string;
}

/** The largest 3-D length kept (`ST_PositiveCoordinate` is larger): 4000 pt. */
const MAX_LENGTH = 50_800_000;

const length = (e: Element | undefined, name: string, d: number): number =>
  Math.max(0, Math.min(intAttr(e, name) ?? d, MAX_LENGTH));

/** An `ST_PositiveFixedAngle` in degrees, 0 to 360. */
function angle(e: Element | undefined, name: string): number {
  const v = ((intAttr(e, name) ?? 0) / 60_000) % 360;
  return v < 0 ? v + 360 : v;
}

const NAME = /^[A-Za-z0-9]{1,40}$/;

/** A preset name as written, when it is a plain token; otherwise `d`. */
const token = (e: Element | undefined, name: string, d: string): string => {
  const v = attr(e, name);
  return v && NAME.test(v) ? v : d;
};

function bevelOf(e: Element | undefined): ResolvedBevel | undefined {
  if (!e) return undefined;
  const preset = attr(e, 'prst');
  const w = length(e, 'w', 76_200);
  const h = length(e, 'h', 76_200);
  if (w === 0 || h === 0) return undefined;
  return {
    w,
    h,
    preset: (BEVEL_PRESET_TYPES as readonly string[]).includes(preset ?? '')
      ? (preset as BevelPresetType)
      : 'circle',
  };
}

/**
 * The 3-D properties from a scene (`a:scene3d`) and a shape (`a:sp3d`) element, each from the
 * nearest level that has one. Undefined when neither adds anything to draw: a front camera with
 * no bevel, extrusion, or contour is a flat shape.
 */
export function resolveThreeD(
  scene: Element | undefined,
  sp3d: Element | undefined,
  ctx: ColorContext
): ResolvedThreeD | undefined {
  if (!scene && !sp3d) return undefined;
  const camera = child(scene, 'a:camera');
  const rot = child(camera, 'a:rot');
  const rig = child(scene, 'a:lightRig');
  const color = (name: string): Rgba | undefined => {
    const el = child(sp3d, name);
    return el ? resolveColor(colorChild(el), ctx) : undefined;
  };
  const extrusionColor = color('a:extrusionClr');
  const contourColor = color('a:contourClr');
  const bevelTop = bevelOf(child(sp3d, 'a:bevelT'));
  const bevelBottom = bevelOf(child(sp3d, 'a:bevelB'));
  const out: ResolvedThreeD = {
    camera: token(camera, 'prst', 'orthographicFront'),
    ...(rot
      ? { cameraRot: { lat: angle(rot, 'lat'), lon: angle(rot, 'lon'), rev: angle(rot, 'rev') } }
      : {}),
    lightRig: token(rig, 'rig', 'threePt'),
    lightDir: token(rig, 'dir', 't'),
    ...(bevelTop ? { bevelTop } : {}),
    ...(bevelBottom ? { bevelBottom } : {}),
    extrusionH: length(sp3d, 'extrusionH', 0),
    ...(extrusionColor ? { extrusionColor } : {}),
    contourW: length(sp3d, 'contourW', 0),
    ...(contourColor ? { contourColor } : {}),
    material: token(sp3d, 'prstMaterial', 'warmMatte'),
  };
  const rotated = out.cameraRot
    ? out.cameraRot.lat !== 0 || out.cameraRot.lon !== 0 || out.cameraRot.rev !== 0
    : !/^(orthographicFront|perspectiveFront|legacyObliqueFront|legacyPerspectiveFront)$/.test(
        out.camera
      );
  const drawn = rotated || out.bevelTop || out.extrusionH > 0 || out.contourW > 0;
  return drawn ? out : undefined;
}

/**
 * The 3-D properties of a shape: the nearest `a:scene3d` and `a:sp3d` of the property elements
 * in `chain` (the shape's own last), each falling back to the theme effect style that the
 * shape's `a:effectRef` names (CT_EffectStyleItem).
 */
export function threeDOf(
  chain: readonly (Element | undefined)[],
  effectStyle: Element | undefined,
  ctx: ColorContext,
  styleCtx: ColorContext
): ResolvedThreeD | undefined {
  let scene: Element | undefined;
  let sp3d: Element | undefined;
  let sp3dCtx = ctx;
  for (let i = chain.length - 1; i >= 0; i -= 1) {
    scene ??= child(chain[i], 'a:scene3d');
    sp3d ??= child(chain[i], 'a:sp3d');
  }
  scene ??= child(effectStyle, 'a:scene3d');
  if (!sp3d) {
    sp3d = child(effectStyle, 'a:sp3d');
    sp3dCtx = styleCtx;
  }
  return resolveThreeD(scene, sp3d, sp3dCtx);
}

type Shadow = EffectPlacement & { readonly color: Rgba };

/** How a preset shadow places its copies, beyond `dist` and `dir`. */
interface PresetShadowShape {
  readonly sx?: number;
  readonly sy?: number;
  /** Horizontal skew, degrees. */
  readonly kx?: number;
  readonly align?: EffectPlacement['align'];
  /** A second, lighter copy: `far` at twice the distance, `back` in the opposite direction. */
  readonly second?: 'far' | 'back';
}

const BACK_LEFT: PresetShadowShape = { sy: 0.5, kx: 60, align: 'b' };
const BACK_RIGHT: PresetShadowShape = { sy: 0.5, kx: -60, align: 'b' };
const FRONT_LEFT: PresetShadowShape = { sy: -0.5, kx: -60, align: 'b' };
const FRONT_RIGHT: PresetShadowShape = { sy: -0.5, kx: 60, align: 'b' };

/**
 * The twenty preset shadows (`ST_PresetShadowVal`), as measured on the `three-d-effects`
 * fixture (slide 5): each is a hard-edged copy at `dist` and `dir`; the perspective presets
 * squash it to half the height on the bottom edge and lean it by 60 degrees.
 */
const PRESET_SHADOWS: Readonly<Record<string, PresetShadowShape>> = {
  shdw1: {},
  shdw2: {},
  shdw3: BACK_LEFT,
  shdw4: BACK_RIGHT,
  shdw5: {},
  shdw6: {},
  shdw7: FRONT_LEFT,
  shdw8: FRONT_RIGHT,
  shdw9: { sx: 0.9, sy: 0.9, align: 'ctr' },
  shdw10: { sx: 1.25, sy: 1.25, align: 'br' },
  shdw11: BACK_LEFT,
  shdw12: BACK_RIGHT,
  shdw13: { second: 'far' },
  shdw14: {},
  shdw15: FRONT_LEFT,
  shdw16: FRONT_RIGHT,
  shdw17: { second: 'back' },
  shdw18: { second: 'back' },
  shdw19: { sy: 0.5, align: 'b' },
  shdw20: { sy: -0.5, align: 'b' },
};

/** The share of the color's alpha that the second copy of a preset shadow has. */
const SECOND_ALPHA = 0.6;

/** The copies of a preset shadow (`a:prstShdw`): one, or two for the double and box presets. */
export function presetShadows(e: Element, ctx: ColorContext): Shadow[] {
  const shape = PRESET_SHADOWS[attr(e, 'prst') ?? ''] ?? {};
  const color = resolveColor(colorChild(e), ctx) ?? { r: 0, g: 0, b: 0, a: 1 };
  const dist = length(e, 'dist', 0);
  const dir = angle(e, 'dir');
  const first: Shadow = {
    blur: 0,
    dist,
    dir,
    sx: shape.sx ?? 1,
    sy: shape.sy ?? 1,
    kx: shape.kx ?? 0,
    ky: 0,
    align: shape.align ?? 'b',
    rotWithShape: true,
    color,
  };
  if (!shape.second) return [first];
  const faint = { ...color, a: color.a * SECOND_ALPHA };
  return [
    first,
    shape.second === 'far'
      ? { ...first, dist: Math.min(dist * 2, MAX_LENGTH), color: faint }
      : { ...first, dir: (dir + 180) % 360, color: faint },
  ];
}
