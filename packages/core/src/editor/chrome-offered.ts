// The chrome choices this build offers. Glow, inner shadow, reflection, and soft edge paint only
// through the `effects` drawing renderer (`@pptx-editor.dev/pro/drawing`). Without it, the chrome
// does not offer the Shape Effects and Picture Styles options that write them, as DOCX Editor
// leaves out the controls of a module that is not installed. The registry (`CHROME_GROUPS`) keeps
// every option, and the commands keep every effect: a host or an automation can still write
// them, and save keeps them.

import type { ImageStyleId, ShadowPreset } from '../contracts/index.ts';
import { drawingCapabilities } from '../layout/index.ts';
import { SHADOW_PRESETS } from '../automation/commands/effect-commands.ts';
import type { ChromeControl, ChromeSlotId } from './chrome-controls.ts';
import { PICTURE_STYLES } from './picture-presets.ts';

/** The effect list elements that only the `effects` renderer paints (ECMA-376 Part 1, 20.1.8). */
const RENDERER_EFFECTS = /<a:(?:glow|innerShdw|reflection|softEdge)\b/;

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

const offered = new WeakMap<ChromeControl, ChromeControl>();

/**
 * The control a slot shows in this build: its registry entry without the options whose result
 * does not paint, and without an option group that has no option left. With the
 * `@pptx-editor.dev/pro/drawing` module installed, or for a slot with no such option, it is the
 * control itself. The adapters read it each time they draw a slot's choices, and draw again after
 * `installDrawingRenderers` (`subscribeDrawingRenderers`), so the controls refresh after an
 * install. @public
 */
export function chromeOfferedControl(slot: ChromeSlotId, control: ChromeControl): ChromeControl {
  if (!control.options || drawingCapabilities().effects) return control;
  if (slot !== 'shape.effects' && slot !== 'image.styles') return control;
  const known = offered.get(control);
  if (known) return known;
  const options = control.options.filter((o) => !chromeOptionNeedsEffects(slot, o.value));
  const optionGroups = control.optionGroups?.filter((g) => options.some((o) => o.group === g.id));
  const shown: ChromeControl = Object.freeze({
    ...control,
    options,
    ...(optionGroups ? { optionGroups } : {}),
  });
  offered.set(control, shown);
  return shown;
}
