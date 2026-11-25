/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `SlideTransition`, an SDK extension: the transition of a slide (`p:transition`,
// `CT_SlideTransition`, ECMA-376 Part 1, §19.3.1.50). The read and the write are the engine's
// `readTransition` and `setTransition` operations, which the editor's transition command uses
// too. A write that changes only the timing or the direction keeps the effect element and the
// markup-compatibility branches that hold it (Part 3, §7.5 `AlternateContent`).

import {
  MAX_ADVANCE_MS,
  MAX_TRANSITION_MS,
  MIN_TRANSITION_MS,
  type AutomationTransitionRead,
  type AutomationTransitionWrite,
  type TransitionDirection,
} from '@pptx-editor.dev/core/automation';
import { failArgument, nearest, oneOf } from '../runtime/errors.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import type { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import {
  TRANSITION_DIRECTION,
  TRANSITION_EFFECT,
  type SlideTransitionDirection,
  type SlideTransitionEffect,
} from './transition-enums.ts';
import { flag, ModelObject } from './model-object.ts';

/**
 * The settings that `SlideTransition.set` writes in one step. Absent settings keep their
 * values. An SDK extension: the pinned Office.js reference has no such type.
 *
 * @public
 */
export interface SlideTransitionOptions {
  /**
   * The effect: `None`, `Fade`, `Push`, `Wipe`, `Split`, or `Cut`. `None` removes the effect
   * and keeps the advance settings. The other values read only and refuse.
   */
  effect?: SlideTransitionEffect | `${SlideTransitionEffect}` | undefined;
  /** How long the effect plays, from 10 to 59,000 milliseconds. Refused with no effect. */
  durationMs?: number | undefined;
  /** Whether a click moves to the next slide. */
  advanceOnClick?: boolean | undefined;
  /** The time after which the slide moves on by itself, from 0 to 3,600,000 ms, or null. */
  advanceAfterMs?: number | null | undefined;
  /** The direction of `Push`, `Wipe`, or `Split`. Refused for another effect. */
  direction?: SlideTransitionDirection | `${SlideTransitionDirection}` | undefined;
  /** Whether `Fade` or `Cut` goes through black first. Refused for another effect. */
  throughBlack?: boolean | undefined;
}

const NAMES = [
  'effect',
  'durationMs',
  'advanceOnClick',
  'advanceAfterMs',
  'direction',
  'throughBlack',
] as const;

/** A number of milliseconds from `min` to `max`, or `InvalidArgument` naming `target`. */
function milliseconds(value: unknown, min: number, max: number, target: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max)
    failArgument(target, `a number of milliseconds from ${min} to ${max}`);
  return value;
}

function duration(value: unknown, target: string): number {
  return milliseconds(value, MIN_TRANSITION_MS, MAX_TRANSITION_MS, target);
}

function advanceAfter(value: unknown, target: string): number | null {
  return value === null ? null : milliseconds(value, 0, MAX_ADVANCE_MS, target);
}

function effect(value: unknown, target: string): AutomationTransitionWrite['kind'] {
  return TRANSITION_EFFECT.toFile(value, target) as AutomationTransitionWrite['kind'];
}

function direction(value: unknown, target: string): TransitionDirection {
  return TRANSITION_DIRECTION.toFile(value, target) as TransitionDirection;
}

/** The effect of a read: the standard element by name, else `None` or `Other`. */
function effectOf(t: AutomationTransitionRead): `${SlideTransitionEffect}` {
  if (t.kind === 'none') return 'None';
  return TRANSITION_EFFECT.fromFile(t.element ?? 'other') ?? 'Other';
}

/**
 * The transition of a slide: the effect that shows the slide, how long it plays, and when the
 * slide show moves to the next slide. The settings are those of `p:transition`
 * (`CT_SlideTransition`, ECMA-376 Part 1, §19.3.1.50) on the slide that the effect shows.
 *
 * A load reads the effect of the first branch that the file holds, and the duration from
 * `p14:dur`, else from the speed `spd` (`ST_TransitionSpeed`, §19.7.54: `fast` 500 ms, `med`
 * 750 ms, `slow` 1000 ms). A slide with no transition reads `None`, 750 ms, advance on click,
 * and no advance time. A `p:transition` with no `spd` reads 500 ms, the schema default `fast`.
 *
 * Each property write and each `set` or `remove` call queues one write. The writes of one
 * `context.sync()` are one transaction, so a browser runtime records them as one undo step.
 * A write that changes nothing writes nothing.
 *
 * An SDK extension: the pinned Office.js reference has no such type.
 *
 * @public
 */
export class SlideTransition extends ModelObject {
  readonly #slide: ObjectPath;
  /** Queue one `setTransition` write on the slide. A field, so the argument checks skip it. */
  readonly #write = (name: string, transition: AutomationTransitionWrite): void => {
    const slide = this.#slide;
    this.command(name, (ref) => ({ op: 'setTransition', slide: ref(slide), transition }));
  };

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, slide: ObjectPath) {
    super(context, path);
    this.#slide = slide;
  }

  /**
   * The effect, in PascalCase: the name of the effect element. A write takes `None`, `Fade`,
   * `Push`, `Wipe`, `Split`, or `Cut`; any other value refuses with `InvalidArgument`. The other
   * standard effects (such as `Dissolve`) read by name, and an effect from an extension
   * namespace reads `Other`. A write of a new effect replaces the transition markup and keeps
   * the duration, the advance settings, and the transition sound (`p:sndAc`). A new `Push` goes up and a new `Split` is
   * vertical-out, unless the same write sets `direction`. `None` removes the effect element;
   * the advance settings stay (see `remove`).
   */
  get effect(): SlideTransitionEffect | `${SlideTransitionEffect}` {
    return this.loadedProperty('effect');
  }
  set effect(value: SlideTransitionEffect | `${SlideTransitionEffect}`) {
    this.#write('effect', { kind: effect(value, this.at('effect')) });
  }

  /**
   * How long the effect plays, in milliseconds, from 10 to 59,000. A write saves `p14:dur` and
   * the nearest speed (`spd`). A slide with no effect refuses a duration write with
   * `NotSupported`: set `effect` in the same `set` call. An extension effect with no `p14`
   * branch keeps only a speed: it refuses a duration other than 500, 750, or 1000 with
   * `NotSupported`.
   */
  get durationMs(): number {
    return this.loadedProperty('durationMs');
  }
  set durationMs(value: number) {
    this.#write('durationMs', { durationMs: duration(value, this.at('durationMs')) });
  }

  /** Whether a click moves to the next slide (`advClick`; true when absent). */
  get advanceOnClick(): boolean {
    return this.loadedProperty('advanceOnClick');
  }
  set advanceOnClick(value: boolean) {
    const advanceOnClick = flag(value, this.at('advanceOnClick'));
    this.#write('advanceOnClick', { advanceOnClick });
  }

  /**
   * The time after which the slide show moves on by itself, in milliseconds (`advTm`), from 0
   * to 3,600,000, or null for no automatic advance. A write of null removes `advTm`.
   */
  get advanceAfterMs(): number | null {
    return this.loadedProperty('advanceAfterMs');
  }
  set advanceAfterMs(value: number | null) {
    const advanceAfterMs = advanceAfter(value, this.at('advanceAfterMs'));
    this.#write('advanceAfterMs', { advanceAfterMs });
  }

  /**
   * The direction of the effect: the side of `Push` and `Wipe` (`dir`, default `Left`), or the
   * orientation and direction of `Split` (`orient` and `dir`, default `HorizontalOut`). Null for
   * any other effect. A write to another effect refuses with `NotSupported`.
   */
  get direction(): SlideTransitionDirection | `${SlideTransitionDirection}` | null {
    return this.loadedProperty('direction');
  }
  set direction(value: SlideTransitionDirection | `${SlideTransitionDirection}`) {
    this.#write('direction', { direction: direction(value, this.at('direction')) });
  }

  /**
   * Whether `Fade` or `Cut` goes through black first (`thruBlk`). False for any other effect. A
   * write to another effect refuses with `NotSupported`.
   */
  get throughBlack(): boolean {
    return this.loadedProperty('throughBlack');
  }
  set throughBlack(value: boolean) {
    const throughBlack = flag(value, this.at('throughBlack'));
    this.#write('throughBlack', { throughBlack });
  }

  /**
   * Write several settings in one step. Absent settings keep their values, so the order of the
   * settings does not matter. An option name that is not a setting refuses with
   * `InvalidArgument`, and nothing is written.
   *
   * @example
   * ```ts
   * slide.transition.set({ effect: 'Push', direction: 'Left', durationMs: 700 });
   * await context.sync();
   * ```
   */
  set(options: SlideTransitionOptions): void {
    const target = this.at('set');
    if (typeof options !== 'object' || options === null || Array.isArray(options))
      failArgument(target, 'an options object');
    for (const key of Object.keys(options))
      if (!(NAMES as readonly string[]).includes(key))
        failArgument(`${target}.${key}`, `an option name: ${oneOf(NAMES)}`, nearest(key, NAMES));
    const write: {
      -readonly [K in keyof AutomationTransitionWrite]: AutomationTransitionWrite[K];
    } = {};
    if (options.effect !== undefined) write.kind = effect(options.effect, `${target}.effect`);
    if (options.durationMs !== undefined)
      write.durationMs = duration(options.durationMs, `${target}.durationMs`);
    if (options.advanceOnClick !== undefined)
      write.advanceOnClick = flag(options.advanceOnClick, `${target}.advanceOnClick`);
    if (options.advanceAfterMs !== undefined)
      write.advanceAfterMs = advanceAfter(options.advanceAfterMs, `${target}.advanceAfterMs`);
    if (options.direction !== undefined)
      write.direction = direction(options.direction, `${target}.direction`);
    if (options.throughBlack !== undefined)
      write.throughBlack = flag(options.throughBlack, `${target}.throughBlack`);
    this.#write('set', write);
  }

  /**
   * Remove the transition: no effect, advance on click, no advance time, and no sound. The slide
   * then holds no `p:transition`. A slide with no transition writes nothing.
   */
  remove(): void {
    this.#write('remove', { kind: 'none', advanceOnClick: true, advanceAfterMs: null });
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const slide = this.#slide;
    this.loadScalars(
      this.selection(request, NAMES),
      'transition',
      (ref) => ({ op: 'readTransition', slide: ref(slide) }),
      ({ transition: t }) => ({
        effect: effectOf(t),
        durationMs: t.durationMs,
        advanceOnClick: t.advanceOnClick,
        advanceAfterMs: t.advanceAfterMs,
        direction: TRANSITION_DIRECTION.fromFile(t.direction),
        throughBlack: t.throughBlack,
      })
    );
  }
}
