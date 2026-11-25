// The transition of a slide (`readTransition`, `setTransition`; the object model's
// `Slide.transition`): `p:transition` (`CT_SlideTransition`, ECMA-376 Part 1, §19.3.1.50) with
// its effect element, `spd` (`ST_TransitionSpeed`, §19.7.54), `advClick`, and `advTm`. The read
// and the write are the editor's own (`commands/transition-commands.ts`), so an effect that the
// editor does not write and its markup-compatibility branches (Part 3, §7.5 `AlternateContent`)
// stay when only the timing changes.
//
// This module checks the write before the edit: an argument of the wrong form or out of range
// refuses with `invalid-argument`, and nothing is written. A write that changes nothing writes
// nothing, as the other property writes do.

import { xmlPartOf } from '../store/index.ts';
import { CommandRefusal } from './commands/op-batch.ts';
import {
  MAX_ADVANCE_MS,
  MAX_TRANSITION_MS,
  MIN_TRANSITION_MS,
  readTransition as readTransitionOf,
  setTransition as setTransitionEdit,
  type TransitionChange,
} from './commands/transition-commands.ts';
import type { AutomationTransitionWrite } from './operations.ts';
import type { AutomationRef, AutomationValue } from './protocol.ts';
import {
  AutomationFailure,
  badArgument,
  edit,
  slidePartOf,
  targetOf,
  type PlanContext,
} from './targets.ts';

/**
 * The effects a write names: the elements the editor writes (ECMA-376 Part 1, §19.5.34 `cut`,
 * §19.5.41 `fade`, §19.5.59 `push`, §19.5.71 `split`, and §19.5.96 `wipe`), or `none`.
 */
const KINDS: ReadonlySet<string> = new Set(['none', 'fade', 'push', 'wipe', 'split', 'cut']);
const FIELDS: ReadonlySet<string> = new Set([
  'kind',
  'durationMs',
  'advanceOnClick',
  'advanceAfterMs',
]);

const inRange = (value: unknown, min: number, max: number): value is number =>
  typeof value === 'number' && Number.isFinite(value) && value >= min && value <= max;

/** The write as an edit change, or `invalid-argument` naming the first field that is wrong. */
function changeOf(write: AutomationTransitionWrite): TransitionChange {
  if (typeof write !== 'object' || write === null || Array.isArray(write))
    badArgument('transition');
  for (const key of Object.keys(write)) if (!FIELDS.has(key)) badArgument('transition');
  const { kind, durationMs, advanceOnClick, advanceAfterMs } = write;
  if (kind !== undefined && (typeof kind !== 'string' || !KINDS.has(kind)))
    badArgument('transitionKind');
  if (durationMs !== undefined && !inRange(durationMs, MIN_TRANSITION_MS, MAX_TRANSITION_MS))
    badArgument('durationMs');
  if (advanceOnClick !== undefined && typeof advanceOnClick !== 'boolean')
    badArgument('advanceOnClick');
  if (
    advanceAfterMs !== undefined &&
    advanceAfterMs !== null &&
    !inRange(advanceAfterMs, 0, MAX_ADVANCE_MS)
  )
    badArgument('advanceAfterMs');
  return {
    ...(kind !== undefined ? { kind } : {}),
    ...(durationMs !== undefined ? { durationMs } : {}),
    ...(advanceOnClick !== undefined ? { advanceOnClick } : {}),
    ...(advanceAfterMs !== undefined ? { advanceAfterMs } : {}),
  };
}

/** The transition of a slide. A slide with no `p:transition` reads `none`. */
export function readTransition(ctx: PlanContext, ref: AutomationRef): AutomationValue {
  const { sldId } = targetOf(ctx, ref, 'slide');
  const part = slidePartOf(ctx.batch.pkg, sldId);
  return { kind: 'transition', transition: readTransitionOf(xmlPartOf(ctx.batch.pkg, part)?.root) };
}

/** Set or clear the transition of a slide, as the editor's `setTransition` command does. */
export function setTransition(
  ctx: PlanContext,
  ref: AutomationRef,
  write: AutomationTransitionWrite
): AutomationValue {
  const { sldId } = targetOf(ctx, ref, 'slide');
  const change = changeOf(write);
  const part = slidePartOf(ctx.batch.pkg, sldId);
  const current = readTransitionOf(xmlPartOf(ctx.batch.pkg, part)?.root);
  // `none` has no effect element, so it has no duration to keep: the edit would drop it.
  if (change.durationMs !== undefined && (change.kind ?? current.kind) === 'none')
    throw new AutomationFailure(
      'unsupported-content',
      'a slide with no transition effect has no duration',
      'transition-no-effect'
    );
  edit(() => {
    try {
      setTransitionEdit(ctx.batch, part, change);
    } catch (e) {
      // The slide has these settings already: nothing to write.
      if (!(e instanceof CommandRefusal && e.reason === 'disabled.noAnimationChange')) throw e;
    }
  });
  return { kind: 'applied' };
}
