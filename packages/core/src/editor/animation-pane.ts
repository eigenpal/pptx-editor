// The Animations tab of the side panel, headless: the menus it offers, the labels of an effect,
// and the seconds fields. The React and Vue panes (`AnimationsPane.tsx`) render from these, so
// both show the same choices with the same keys.

import type {
  AnimationEffect,
  AnimationPreset,
  AnimationTrigger,
  SlideTransitionState,
  TransitionKind,
} from '../contracts/index.ts';

/** The presets of the Add effect menu, by class. @public */
export const ANIMATION_PRESET_GROUPS: readonly {
  readonly effectClass: 'entrance' | 'emphasis' | 'exit';
  readonly presets: readonly AnimationPreset[];
}[] = [
  { effectClass: 'entrance', presets: ['appear', 'fade', 'flyIn', 'wipe', 'zoom'] },
  { effectClass: 'emphasis', presets: ['pulse', 'spin', 'growShrink'] },
  { effectClass: 'exit', presets: ['disappear', 'fadeOut', 'flyOut', 'wipeOut', 'zoomOut'] },
];

/** The transition kinds of the Transition menu, `none` first. @public */
export const TRANSITION_KINDS: readonly (TransitionKind | 'none')[] = [
  'none',
  'fade',
  'push',
  'wipe',
  'split',
  'cut',
];

/** The start options of an effect. @public */
export const ANIMATION_TRIGGERS: readonly AnimationTrigger[] = [
  'onClick',
  'withPrevious',
  'afterPrevious',
];

/** The i18n key of an effect's name: its preset, else its class. @public */
export function animationEffectLabelKey(
  effect: Pick<AnimationEffect, 'preset' | 'effectClass'>
):
  | `panel.animations.presets.${AnimationPreset}`
  | `panel.animations.classes.${AnimationEffect['effectClass']}` {
  return effect.preset
    ? `panel.animations.presets.${effect.preset}`
    : `panel.animations.classes.${effect.effectClass}`;
}

/** The i18n key of a transition kind. @public */
export const transitionKindLabelKey = (
  kind: SlideTransitionState['kind']
): `panel.animations.kinds.${SlideTransitionState['kind']}` => `panel.animations.kinds.${kind}`;

/** Milliseconds as seconds for a field: two decimals at most, no trailing zeros. @public */
export function secondsText(ms: number): string {
  return String(Math.round(ms / 10) / 100);
}

/**
 * Seconds typed in a field, as milliseconds, or null when the text is not a number of seconds
 * from 0 to 3,600. A comma is a decimal separator. @public
 */
export function parseSeconds(text: string): number | null {
  const t = text.trim().replace(',', '.');
  if (!/^\d{0,4}(\.\d{0,3})?$/.test(t) || t === '' || t === '.') return null;
  const n = Number(t);
  return Number.isFinite(n) && n <= 3600 ? Math.round(n * 1000) : null;
}

/** The effects that the pane lists as the main sequence, and the others. @public */
export function animationListParts(effects: readonly AnimationEffect[]): {
  readonly main: readonly AnimationEffect[];
  readonly other: readonly AnimationEffect[];
} {
  return {
    main: effects.filter((e) => e.sequence === 'main'),
    other: effects.filter((e) => e.sequence !== 'main'),
  };
}

/**
 * When an effect starts, for its row: the i18n key and its values. An effect of the main
 * sequence plays at a click step, or with the slide start (step 0). An effect whose step the
 * editor cannot read (in a read-only timing) shows its trigger. An effect of an interactive
 * sequence plays on a click of its trigger shape; `shapeName` is empty for a shape with no
 * name. @public
 */
export type AnimationStartLabel =
  | { readonly key: 'panel.animations.withSlide' }
  | { readonly key: 'panel.animations.step'; readonly step: number }
  | { readonly key: 'panel.animations.onShapeClick'; readonly shapeName: string }
  | { readonly key: `panel.animations.triggers.${AnimationTrigger}` };

/** When an effect starts, for its row (see {@link AnimationStartLabel}). @public */
export function animationStartLabel(
  effect: Pick<AnimationEffect, 'sequence' | 'step' | 'trigger' | 'triggerShapeName'>
): AnimationStartLabel {
  if (effect.step !== null)
    return effect.step === 0
      ? { key: 'panel.animations.withSlide' }
      : { key: 'panel.animations.step', step: effect.step };
  if (effect.sequence === 'interactive')
    return { key: 'panel.animations.onShapeClick', shapeName: effect.triggerShapeName ?? '' };
  return { key: `panel.animations.triggers.${effect.trigger}` };
}
