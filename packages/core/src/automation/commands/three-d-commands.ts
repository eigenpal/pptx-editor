// The Bevel and 3-D Rotation galleries: a top bevel in `a:sp3d` and a camera preset in
// `a:scene3d` (ECMA-376 Part 1, 20.1.5.12 and 20.1.4.1.26). Each edit changes the elements in
// place: attributes and children it does not own stay, and new children go in schema order
// (CT_ShapeProperties, CT_Scene3D, CT_Camera, CT_Shape3D), so saved files stay valid.

import { attr, child, NS, type Element, type NodeSpec } from '../../store/index.ts';
import type { ResolvedEffects } from '../../cascade/index.ts';
import type { BevelPreset, Rotation3DPreset, ShapeEffectChoice } from '../../contracts/neutral.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { specFromXml } from './shape-commands.ts';

/** The Bevel gallery, in gallery order: every top bevel preset at 6 pt by 6 pt. */
export const BEVEL_PRESETS: readonly BevelPreset[] = [
  'circle',
  'relaxedInset',
  'cross',
  'coolSlant',
  'angle',
  'softRound',
  'convex',
  'slope',
  'divot',
  'riblet',
  'hardEdge',
  'artDeco',
];

/** The 3-D Rotation gallery, in gallery order: parallel camera presets. */
export const ROTATION_3D_PRESETS: readonly Rotation3DPreset[] = [
  'isometricLeftDown',
  'isometricRightUp',
  'isometricTopUp',
  'isometricOffAxis1Left',
  'isometricOffAxis1Right',
  'obliqueTopLeft',
  'obliqueTopRight',
  'obliqueBottomLeft',
  'obliqueBottomRight',
];

/** The bevel size the gallery writes: 6 pt (`ST_PositiveCoordinate` default of `CT_Bevel`). */
const BEVEL_SIZE = 76_200;

/** The `p:spPr` children that come before `a:scene3d` (CT_ShapeProperties). */
const BEFORE_SCENE = new Set([
  'xfrm',
  'custGeom',
  'prstGeom',
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'ln',
  'effectLst',
  'effectDag',
]);
const BEFORE_SP3D = new Set([...BEFORE_SCENE, 'scene3d']);
const FRONT = 'orthographicFront';
const DEFAULT_SCENE = `<a:scene3d><a:camera prst="${FRONT}"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>`;

/** Insert `spec` into `parent` after its last `a:` child named in `before`. */
function insertAfter(
  batch: OpBatch,
  part: string,
  parent: Element,
  before: ReadonlySet<string>,
  spec: NodeSpec
): string {
  let index = 0;
  parent.children.forEach((c, i) => {
    if (c.type === 'element' && c.ns === NS.a && before.has(c.local)) index = i + 1;
  });
  return batch.insert({ op: 'insertNode', part, parent: parent.id, index, spec });
}

function setAttr(
  batch: OpBatch,
  part: string,
  el: Element,
  name: string,
  value: string | null
): void {
  if ((attr(el, name) ?? null) === value) return;
  batch.push({ op: 'setAttribute', part, node: el.id, name: { ns: '', local: name }, value });
}

/** The shape's own `a:scene3d`, written with a front camera when it has none. */
function sceneOf(batch: OpBatch, part: string, spPr: Element): Element {
  const own = child(spPr, 'a:scene3d');
  if (own) return own;
  return batch.element(
    part,
    insertAfter(batch, part, spPr, BEFORE_SCENE, specFromXml(DEFAULT_SCENE))
  );
}

/** Whether a choice is a 3-D choice that this module writes. */
export function isThreeDChoice(
  choice: ShapeEffectChoice
): choice is Extract<ShapeEffectChoice, { kind: 'bevel' | 'rotation3d' }> {
  return choice?.kind === 'bevel' || choice?.kind === 'rotation3d';
}

/**
 * Set or remove the top bevel of one shape. Returns false when removing changes nothing. An
 * inherited bevel (from a theme effect style) is removed with an own `a:sp3d` that has none. A
 * shape that inherits a scene keeps it: the bevel writes no scene of its own.
 */
function setBevel(
  batch: OpBatch,
  part: string,
  spPr: Element,
  preset: BevelPreset | null,
  inherited: InheritedThreeD
): boolean {
  let sp3d = child(spPr, 'a:sp3d');
  const bevel = child(sp3d, 'a:bevelT');
  if (preset === null) {
    if (bevel) {
      batch.push({ op: 'removeNode', part, node: bevel.id });
      return true;
    }
    if (sp3d || !inherited.bevel) return false;
    insertAfter(batch, part, spPr, BEFORE_SP3D, specFromXml('<a:sp3d/>'));
    return true;
  }
  if (!inherited.anyScene) sceneOf(batch, part, spPr);
  if (!sp3d) {
    const current = batch.element(part, spPr.id);
    sp3d = batch.element(
      part,
      insertAfter(batch, part, current, BEFORE_SP3D, specFromXml('<a:sp3d/>'))
    );
  }
  if (bevel) {
    setAttr(batch, part, bevel, 'w', String(BEVEL_SIZE));
    setAttr(batch, part, bevel, 'h', String(BEVEL_SIZE));
    setAttr(batch, part, bevel, 'prst', preset === 'circle' ? null : preset);
    return true;
  }
  const xml = `<a:bevelT w="${BEVEL_SIZE}" h="${BEVEL_SIZE}"${preset === 'circle' ? '' : ` prst="${preset}"`}/>`;
  // CT_Shape3D: bevelT is the first child.
  batch.insert({ op: 'insertNode', part, parent: sp3d.id, index: 0, spec: specFromXml(xml) });
  return true;
}

/** Set the camera preset of one shape, or turn it back to the front. */
function setRotation(
  batch: OpBatch,
  part: string,
  spPr: Element,
  preset: Rotation3DPreset | null,
  inherited: boolean
): boolean {
  const own = child(spPr, 'a:scene3d');
  if (preset === null && !own && !inherited) return false;
  const camera = child(own, 'a:camera');
  const rot = child(camera, 'a:rot');
  if (preset === null && own && camera && attr(camera, 'prst') === FRONT && !rot) return false;
  const scene = sceneOf(batch, part, spPr);
  const cam = child(batch.element(part, scene.id), 'a:camera');
  if (!cam) {
    // CT_Scene3D: the camera is the first child.
    batch.insert({
      op: 'insertNode',
      part,
      parent: scene.id,
      index: 0,
      spec: specFromXml(`<a:camera prst="${preset ?? FRONT}"/>`),
    });
    return true;
  }
  setAttr(batch, part, cam, 'prst', preset ?? FRONT);
  // A rotation replaces the preset's: the gallery choice sets the preset alone.
  const turn = child(cam, 'a:rot');
  if (turn) batch.push({ op: 'removeNode', part, node: turn.id });
  return true;
}

/** What an element inherits from its placeholders and its theme effect style. */
export interface InheritedThreeD {
  /** An `a:scene3d` with a turned camera that is not the shape's own. */
  readonly scene: boolean;
  /** An `a:scene3d` of any camera that is not the shape's own. */
  readonly anyScene: boolean;
  /** An `a:sp3d` with a top bevel that is not the shape's own. */
  readonly bevel: boolean;
}

/**
 * Apply a Bevel or 3-D Rotation choice to the selected shapes, connectors, and pictures.
 * `inherited` says what 3-D properties an element inherits. Throws `disabled.noEffectChange`
 * when no element changes.
 */
export function setThreeD(
  batch: OpBatch,
  part: string,
  targets: readonly Element[],
  choice: Extract<ShapeEffectChoice, { kind: 'bevel' | 'rotation3d' }>,
  inherited: (el: Element) => InheritedThreeD
): void {
  const preset = choice.preset;
  const allowed: readonly string[] = choice.kind === 'bevel' ? BEVEL_PRESETS : ROTATION_3D_PRESETS;
  if (preset !== null && !allowed.includes(preset))
    throw new CommandRefusal('errors.unsupported', choice.kind);
  let changed = 0;
  for (const el of targets) {
    const spPr = child(el, 'p:spPr');
    if (!spPr) throw new CommandRefusal('errors.malformed');
    const from = inherited(el);
    const done =
      choice.kind === 'bevel'
        ? setBevel(batch, part, spPr, choice.preset, from)
        : setRotation(batch, part, spPr, choice.preset, from.scene);
    if (done) changed += 1;
  }
  if (changed === 0) throw new CommandRefusal('disabled.noEffectChange');
}

/** The bevel preset of resolved effects, for the chrome. */
export function bevelStateOf(fx: ResolvedEffects | undefined): BevelPreset | 'none' {
  const bevel = fx?.threeD?.bevelTop;
  return bevel && (BEVEL_PRESETS as readonly string[]).includes(bevel.preset)
    ? (bevel.preset as BevelPreset)
    : 'none';
}

/** The 3-D Rotation choice of resolved effects, for the chrome. */
export function rotationStateOf(
  fx: ResolvedEffects | undefined
): Rotation3DPreset | 'none' | 'other' {
  const t = fx?.threeD;
  if (!t) return 'none';
  const front = /^(orthographicFront|perspectiveFront)$/.test(t.camera);
  const rot = t.cameraRot;
  const turned = rot !== undefined && (rot.lat !== 0 || rot.lon !== 0 || rot.rev !== 0);
  if (front && !turned) return 'none';
  if (rot === undefined && (ROTATION_3D_PRESETS as readonly string[]).includes(t.camera))
    return t.camera as Rotation3DPreset;
  return 'other';
}
