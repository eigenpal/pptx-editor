// The chrome choices this build offers. Glow, inner shadow, reflection, and soft edge paint only
// through the `effects` drawing renderer (`@pptx-editor.dev/pro/drawing`). Without it, the chrome
// does not offer the Shape Effects and Picture Styles options that write them, as DOCX Editor
// leaves out the controls of a module that is not installed. Removing an effect needs no
// renderer: No Reflection, No Glow, and No Soft Edges stay while the selection has that effect,
// so an effect from the file can still be removed. No Shadow always stays, and it also removes
// an inner shadow. The registry (`CHROME_GROUPS`) keeps every option, and the commands keep
// every effect: a host or an automation can still write them, and save keeps them.

import type { Editor, ImageStyleId, ShadowPreset } from '../contracts/index.ts';
import { drawingCapabilities } from '../layout/index.ts';
import { SHADOW_PRESETS } from '../automation/commands/effect-commands.ts';
import type { ChromeControl, ChromeSlotId } from './chrome-controls.ts';
import { PICTURE_STYLES } from './picture-presets.ts';

/** The effect list elements that only the `effects` renderer paints (ECMA-376 Part 1, 20.1.8). */
const RENDERER_EFFECTS = /<a:(?:glow|innerShdw|reflection|softEdge)\b/;

/** The effects whose `none` row a build without the renderer offers when the selection has one. */
const REMOVABLE = ['reflection', 'glow', 'softEdge'] as const;
type Removable = (typeof REMOVABLE)[number];

const NOTHING_TO_REMOVE: ReadonlySet<Removable> = new Set();

const own = (record: object, key: string): boolean =>
  Object.prototype.hasOwnProperty.call(record, key);

/**
 * Whether a slot option writes an effect that only the `effects` drawing renderer paints: an
 * inner shadow, reflection, glow, or soft edge option of `shape.effects` (their `none` rows too),
 * or a picture style of `image.styles` whose effect list has one. @internal
 */
export function chromeOptionNeedsEffects(slot: ChromeSlotId, value: string): boolean {
  if (slot === 'shape.effects') {
    const [kind = '', preset = ''] = value.split('.');
    if (kind === 'shadow')
      return own(SHADOW_PRESETS, preset) && SHADOW_PRESETS[preset as ShadowPreset].inner;
    return kind === 'reflection' || kind === 'glow' || kind === 'softEdge';
  }
  if (slot === 'image.styles') {
    const effects = own(PICTURE_STYLES, value)
      ? PICTURE_STYLES[value as ImageStyleId].effectLst
      : null;
    return effects !== null && RENDERER_EFFECTS.test(effects);
  }
  return false;
}

/**
 * The effects whose removal would change the selection: a selected shape, connector, or picture
 * has the effect, its own or inherited. `editor.can` runs the removal command dry, so a
 * multi-selection, a group (which carries no effect), and a read-only editor answer as the
 * command does. It answers from its cache, once per editor state.
 */
function removableEffects(editor: Pick<Editor, 'can'>): ReadonlySet<Removable> {
  const out = new Set<Removable>();
  for (const kind of REMOVABLE) {
    const ok = editor.can({
      type: 'setShapeEffect',
      effect: kind === 'reflection' ? { kind, preset: null } : { kind, size: null },
    }).ok;
    if (ok) out.add(kind);
  }
  return out;
}

/** Whether `value` is the `none` row of an effect in `removable`. */
function removalRow(value: string, removable: ReadonlySet<Removable>): boolean {
  const [kind = '', choice] = value.split('.');
  return choice === 'none' && (removable as ReadonlySet<string>).has(kind);
}

/** The controls already given, by registry control and then by the removable effects. */
const offered = new WeakMap<ChromeControl, Map<string, ChromeControl>>();

/**
 * The control a slot shows in this build: its registry entry without the options whose result
 * does not paint, and without an option group that has no option left. Pass `editor` to keep
 * the rows that remove an effect the selection has (No Reflection, No Glow, No Soft Edges):
 * removing an effect needs no renderer, so a community build can remove an effect that is in
 * the file. No Shadow always stays, and it also removes an inner shadow. With the
 * `@pptx-editor.dev/pro/drawing` module installed, or for a slot with no such option, it is the
 * control itself. The same choices give the same object. The adapters read it again on each
 * editor change and after each install (`subscribeDrawingRenderers`), so their menus follow the
 * selection. A custom control does the same. @public
 */
export function chromeOfferedControl(
  slot: ChromeSlotId,
  control: ChromeControl,
  editor?: Pick<Editor, 'can'> | null
): ChromeControl {
  if (!control.options || drawingCapabilities().effects) return control;
  if (slot !== 'shape.effects' && slot !== 'image.styles') return control;
  const removable =
    slot === 'shape.effects' && editor ? removableEffects(editor) : NOTHING_TO_REMOVE;
  const key = REMOVABLE.filter((k) => removable.has(k)).join(' ');
  let known = offered.get(control);
  if (!known) offered.set(control, (known = new Map()));
  const hit = known.get(key);
  if (hit) return hit;
  const options = control.options.filter(
    (o) => !chromeOptionNeedsEffects(slot, o.value) || removalRow(o.value, removable)
  );
  const optionGroups = control.optionGroups?.filter((g) => options.some((o) => o.group === g.id));
  const shown: ChromeControl = Object.freeze({
    ...control,
    options,
    ...(optionGroups ? { optionGroups } : {}),
  });
  known.set(key, shown);
  return shown;
}
