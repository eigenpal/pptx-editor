/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The camera and the light of a 3-D shape (ECMA-376 Part 1, 20.1.5): how the camera turns the
// shape, and how bright a surface is for its normal. Pure math: no DOM.
//
// Vectors use slide axes: x to the right, y down, and z toward the viewer for normals and light
// directions. A turned point keeps z as depth away from the viewer.
//
// The camera rotations are measured on the `three-d-effects` reference PDF (slide 3): the
// projected outlines of 12 isometric presets fall within 0.03 in. The perspective presets draw
// as parallel projections.
//
// The light is fitted to the bevel cross-sections of that fixture (slide 1): the three-point rig
// from the top, on the warm matte and plastic materials, within 7 levels of 255 in the green
// channel. Other rigs, directions, and materials are approximate.

import type { ResolvedThreeD } from '@pptx-editor.dev/core/cascade';

export type Vec3 = readonly [number, number, number];

const DEG = Math.PI / 180;

/** Measured rotations of the isometric and perspective presets: lat, lon, rev in 60000ths. */
const CAMERA_ROTATIONS: Readonly<Record<string, readonly [number, number, number]>> = {
  isometricBottomDown: [2_124_000, 18_883_200, 17_990_400],
  isometricBottomUp: [2_124_000, 2_716_800, 3_609_600],
  isometricLeftDown: [2_100_000, 2_700_000, 0],
  isometricLeftUp: [19_500_000, 2_700_000, 0],
  isometricOffAxis1Left: [1_080_000, 3_840_000, 0],
  isometricOffAxis1Right: [1_080_000, 20_040_000, 0],
  isometricOffAxis1Top: [18_075_600, 18_393_000, 3_458_400],
  isometricOffAxis2Left: [1_080_000, 1_560_000, 0],
  isometricOffAxis2Right: [1_080_000, 17_760_000, 0],
  isometricOffAxis2Top: [18_075_600, 3_207_000, 18_141_600],
  isometricOffAxis3Bottom: [3_524_400, 18_393_000, 18_141_600],
  isometricOffAxis3Left: [20_520_000, 3_840_000, 0],
  isometricOffAxis3Right: [20_520_000, 20_040_000, 0],
  isometricOffAxis4Bottom: [3_524_400, 3_207_000, 3_458_400],
  isometricOffAxis4Left: [20_520_000, 1_560_000, 0],
  isometricOffAxis4Right: [20_520_000, 17_760_000, 0],
  isometricRightDown: [19_500_000, 18_900_000, 0],
  isometricRightUp: [2_100_000, 18_900_000, 0],
  isometricTopDown: [19_476_000, 2_716_800, 17_990_400],
  isometricTopUp: [19_476_000, 18_883_200, 3_609_600],
  perspectiveAbove: [20_400_000, 0, 0],
  perspectiveAboveLeftFacing: [2_358_600, 860_400, 20_467_200],
  perspectiveAboveRightFacing: [2_358_600, 20_739_600, 1_132_800],
  perspectiveBelow: [1_200_000, 0, 0],
  perspectiveContrastingLeftFacing: [624_000, 2_636_400, 21_387_000],
  perspectiveContrastingRightFacing: [624_000, 18_963_600, 213_000],
  perspectiveHeroicExtremeLeftFacing: [487_200, 2_067_600, 21_425_400],
  perspectiveHeroicExtremeRightFacing: [487_200, 19_532_400, 174_600],
  perspectiveHeroicLeftFacing: [20_938_800, 857_400, 157_800],
  perspectiveHeroicRightFacing: [20_938_800, 20_742_600, 21_442_200],
  perspectiveLeft: [0, 1_200_000, 0],
  perspectiveRelaxed: [18_573_600, 0, 0],
  perspectiveRelaxedModerately: [19_490_400, 0, 0],
  perspectiveRight: [0, 20_400_000, 0],
};

/**
 * The oblique presets: the direction of the depth, in degrees, and its length as a share of the
 * extrusion. Measured on the fixture (slide 3): the depth shows at 30 % of its length.
 */
const OBLIQUE: Readonly<Record<string, readonly [number, number]>> = {
  obliqueTopLeft: [-45, 0.3],
  obliqueTop: [-90, 0.3],
  obliqueTopRight: [-135, 0.3],
  obliqueLeft: [0, 0.3],
  obliqueRight: [180, 0.3],
  obliqueBottomLeft: [45, 0.3],
  obliqueBottom: [90, 0.3],
  obliqueBottomRight: [135, 0.3],
  legacyObliqueTopLeft: [-45, 0.5],
  legacyObliqueTop: [-90, 0.5],
  legacyObliqueTopRight: [-135, 0.5],
  legacyObliqueLeft: [0, 0.5],
  legacyObliqueRight: [180, 0.5],
  legacyObliqueBottomLeft: [45, 0.5],
  legacyObliqueBottom: [90, 0.5],
  legacyObliqueBottomRight: [135, 0.5],
};

/** A 3 by 3 rotation, row by row. */
export type Rotation = readonly [Vec3, Vec3, Vec3];

const IDENTITY: Rotation = [
  [1, 0, 0],
  [0, 1, 0],
  [0, 0, 1],
];

function product(a: Rotation, b: Rotation): Rotation {
  const row = (r: Vec3): Vec3 => [
    r[0] * b[0][0] + r[1] * b[1][0] + r[2] * b[2][0],
    r[0] * b[0][1] + r[1] * b[1][1] + r[2] * b[2][1],
    r[0] * b[0][2] + r[1] * b[1][2] + r[2] * b[2][2],
  ];
  return [row(a[0]), row(a[1]), row(a[2])];
}

export function turn(r: Rotation, v: Vec3): Vec3 {
  return [
    r[0][0] * v[0] + r[0][1] * v[1] + r[0][2] * v[2],
    r[1][0] * v[0] + r[1][1] * v[1] + r[1][2] * v[2],
    r[2][0] * v[0] + r[2][1] * v[1] + r[2][2] * v[2],
  ];
}

/** How the camera shows the shape. */
export type CameraView =
  | { readonly kind: 'front' }
  | { readonly kind: 'turn'; readonly rotation: Rotation }
  | { readonly kind: 'oblique'; readonly dx: number; readonly dy: number };

/**
 * The view of a camera: the rotation `rev` about z, after `lat` about x, after `lon` about y.
 * The perspective presets draw as parallel projections.
 */
export function cameraView(t: ResolvedThreeD): CameraView {
  const preset = CAMERA_ROTATIONS[t.camera];
  const rot = t.cameraRot
    ? ([t.cameraRot.lat, t.cameraRot.lon, t.cameraRot.rev] as const)
    : preset
      ? ([preset[0] / 60_000, preset[1] / 60_000, preset[2] / 60_000] as const)
      : undefined;
  if (rot && (rot[0] !== 0 || rot[1] !== 0 || rot[2] !== 0)) {
    const [lat, lon, rev] = rot.map((d) => d * DEG) as [number, number, number];
    const ry: Rotation = [
      [Math.cos(lon), 0, Math.sin(lon)],
      [0, 1, 0],
      [-Math.sin(lon), 0, Math.cos(lon)],
    ];
    const rx: Rotation = [
      [1, 0, 0],
      [0, Math.cos(lat), -Math.sin(lat)],
      [0, Math.sin(lat), Math.cos(lat)],
    ];
    const rz: Rotation = [
      [Math.cos(-rev), -Math.sin(-rev), 0],
      [Math.sin(-rev), Math.cos(-rev), 0],
      [0, 0, 1],
    ];
    return { kind: 'turn', rotation: product(rz, product(rx, ry)) };
  }
  const oblique = t.cameraRot ? undefined : OBLIQUE[t.camera];
  if (oblique) {
    const [deg, share] = oblique;
    return { kind: 'oblique', dx: -Math.cos(deg * DEG) * share, dy: Math.sin(deg * DEG) * share };
  }
  return { kind: 'front' };
}

export { IDENTITY };

/** A light: its direction (toward the light) and its strength. */
interface Light {
  readonly dir: Vec3;
  readonly k: number;
}

/** A highlight: the direction toward its light, its strength, and its power. */
interface Lobe extends Light {
  readonly power: number;
}

/**
 * The three-point rig from the top, in the shape's own axes: a camera turn moves the view, not
 * the light. An ambient term, three lights (from above, from below, and from behind), and two
 * highlights at the half vector of the light and the view (Blinn). Fitted together to the
 * fixture: the angle bevels (slides 1 and 2), the faces under 18 cameras (slides 3 and 4), the
 * extrusion sides (slides 3 and 4), and the turned faces (`three-d-turned-effects`, slide 1),
 * within 3.3 levels of 255 on average.
 */
const AMBIENT = 0.1253;
const LIGHTS: readonly Light[] = [
  { dir: [-0.194, -0.722, 0.664], k: 0.857 },
  { dir: [-0.027, 0.669, 0.743], k: 0.486 },
  { dir: [-0.286, 0.571, -0.769], k: 0.929 },
];
const HIGHLIGHTS: readonly Lobe[] = [
  { dir: [-0.355, -0.749, 0.559], k: 0.484, power: 28.5 },
  { dir: [-0.457, 0.459, -0.762], k: 1.125, power: 82.4 },
];

/**
 * How far the rig turns the light about the view axis, by `dir`, in degrees. The top-left
 * direction is measured (the bevel of slide 2 lit from the left); the others follow the compass
 * and are approximate.
 */
const DIR_TURN: Readonly<Record<string, number>> = {
  t: 0,
  tl: -75,
  l: -90,
  bl: -135,
  b: 180,
  br: 135,
  r: 90,
  tr: 45,
};

/**
 * The strength of a rig against the three-point rig, measured on the faces of slide 2 of the
 * fixture. The rigs that the fixture does not show keep the strength of the three-point rig.
 */
const RIG_STRENGTH: Readonly<Record<string, number>> = {
  balanced: 0.93,
  harsh: 0.77,
};

/** The highlight strength of each material: none for matte ones, more for metal. */
function highlightOf(material: string): number {
  if (/^(matte|powder|translucentPowder|dkEdge|softEdge|clear)$/.test(material)) return 0;
  if (/^(metal|softmetal)$/.test(material)) return 2;
  return 1;
}

/** A brightness factor and a white highlight, from 0 to 1, for one surface. */
export interface Shade {
  readonly level: number;
  readonly white: number;
}

/** The light of a scene: the brightness and the white highlight for a surface normal. */
export interface Lighting {
  /** The shade of a normal in the shape's axes. A flat material or rig gives level 1. */
  readonly shade: (n: Vec3) => Shade;
}

/**
 * The direction toward the viewer in the shape's axes (x right, y down, z out of the face). A
 * turn gives the camera axis turned back; an oblique view looks along its depth direction.
 */
export function viewDirection(view: CameraView): Vec3 {
  if (view.kind === 'turn') {
    const R = view.rotation;
    return normalize([-R[2][0], -R[2][1], R[2][2]]);
  }
  if (view.kind === 'oblique') return normalize([view.dx, view.dy, 1]);
  return [0, 0, 1];
}

const dot = (a: Vec3, b: Vec3): number => a[0] * b[0] + a[1] * b[1] + a[2] * b[2];

function spin(v: Vec3, deg: number): Vec3 {
  const c = Math.cos(deg * DEG);
  const s = Math.sin(deg * DEG);
  return [v[0] * c - v[1] * s, v[0] * s + v[1] * c, v[2]];
}

export function normalize(v: Vec3): Vec3 {
  const l = Math.hypot(v[0], v[1], v[2]) || 1;
  return [v[0] / l, v[1] / l, v[2] / l];
}

/**
 * The light of a scene seen from `view` (`viewDirection`). The shade is absolute: the face of a
 * shape (normal `[0, 0, 1]`) is brighter than its fill, as in the fixture. A flat material or
 * rig gives no shading.
 */
export function lightingOf(t: ResolvedThreeD, view: Vec3 = [0, 0, 1]): Lighting {
  if (
    t.material === 'flat' ||
    t.material === 'legacyWireframe' ||
    /^(flat|legacyFlat\d)$/.test(t.lightRig)
  )
    return { shade: () => ({ level: 1, white: 0 }) };
  const turnBy = DIR_TURN[t.lightDir] ?? 0;
  const strength = RIG_STRENGTH[t.lightRig] ?? 1;
  const lights = LIGHTS.map((l) => ({ dir: spin(normalize(l.dir), turnBy), k: strength * l.k }));
  const ambient = strength * AMBIENT;
  const gloss = strength * highlightOf(t.material);
  // The half vector of each highlight is fixed for one view.
  const halves = HIGHLIGHTS.map((h) => {
    const l = spin(normalize(h.dir), turnBy);
    return { half: normalize([l[0] + view[0], l[1] + view[1], l[2] + view[2]]), lobe: h };
  });
  return {
    shade: (n) => ({
      level: lights.reduce((sum, l) => sum + l.k * Math.max(0, dot(n, l.dir)), ambient),
      white: halves.reduce(
        (sum, { half, lobe }) => sum + gloss * lobe.k * Math.max(0, dot(n, half)) ** lobe.power,
        0
      ),
    }),
  };
}
