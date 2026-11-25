/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `SlideTransition`, an SDK extension: the transition of a slide (`p:transition`,
// `CT_SlideTransition`, ECMA-376 Part 1, §19.3.1.50). The read and the write are the engine's
// `readTransition` and `setTransition` operations, which the editor's transition command uses
// too. A write that changes only the timing keeps the effect element and the
// markup-compatibility branches that hold it (Part 3, §7.5 `AlternateContent`).

import {
  MAX_ADVANCE_MS,
  MAX_TRANSITION_MS,
  MIN_TRANSITION_MS,
  type AutomationTransitionWrite,
} from '@pptx-editor.dev/core/automation';
import { failArgument, nearest, oneOf } from '../runtime/errors.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import type { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import { TRANSITION_EFFECT, type SlideTransitionEffect } from './enums.ts';
import { flag, ModelObject } from './model-object.ts';

/**
 * The settings that `SlideTransition.set` writes in one step. Absent settings keep their
 * values. An SDK extension: the pinned Office.js reference has no such type.
 *
 * @public
 */
export interface SlideTransitionOptions {
  /** The effect. `None` removes the effect and keeps the advance settings. `Other` refuses. */
  effect?: SlideTransitionEffect | `${SlideTransitionEffect}` | undefined;
  /** How long the effect plays, from 10 to 59,000 milliseconds. Refused with no effect. */
  durationMs?: number | undefined;
  /** Whether a click moves to the next slide. */
  advanceOnClick?: boolean | undefined;
  /** The time after which the slide moves on by itself, from 0 to 3,600,000 ms, or null. */
  advanceAfterMs?: number | null | undefined;
}

const NAMES = ['effect', 'durationMs', 'advanceOnClick', 'advanceAfterMs'] as const;

/** A whole number of milliseconds from `min` to `max`, or `InvalidArgument` naming `target`. */
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
 * Each property write and each `set` call queues one write. The writes of one
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
   * The effect, in PascalCase. `Other` is an effect that this API does not write, such as
   * `p:dissolve` or an effect from an extension namespace; a write of `Other` refuses with
   * `InvalidArgument`. A write of a new effect replaces the transition markup and keeps the
   * duration and the advance settings. `Push` writes the direction up (`dir="u"`), and `Split`
   * writes the vertical orientation (`orient="vert"`). `None` removes the effect element; the
   * advance settings stay.
   */
  get effect(): SlideTransitionEffect | `${SlideTransitionEffect}` {
    return this.loadedProperty('effect');
  }
  set effect(value: SlideTransitionEffect | `${SlideTransitionEffect}`) {
    this.#write('effect', { kind: effect(value, this.at('effect')) });
  }

  /**
   * How long the effect plays, in milliseconds, from 10 to 59,000. A write saves `p14:dur` in
   * the markup-compatibility choice and the nearest speed (`spd`) in the fallback. A slide with
   * no effect refuses a duration write with `NotSupported`: set `effect` in the same `set` call.
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
   * Write several settings in one step. Absent settings keep their values, so the order of the
   * settings does not matter. An option name that is not a setting refuses with
   * `InvalidArgument`, and nothing is written.
   *
   * @example
   * ```ts
   * slide.transition.set({ effect: 'Fade', durationMs: 700, advanceAfterMs: null });
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
    this.#write('set', write);
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const slide = this.#slide;
    this.loadScalars(
      this.selection(request, NAMES),
      'transition',
      (ref) => ({ op: 'readTransition', slide: ref(slide) }),
      (v) => ({
        effect: TRANSITION_EFFECT.fromFile(v.transition.kind) ?? 'Other',
        durationMs: v.transition.durationMs,
        advanceOnClick: v.transition.advanceOnClick,
        advanceAfterMs: v.transition.advanceAfterMs,
      })
    );
  }
}
