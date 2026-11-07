// The transition of a slide (ECMA-376 Part 1, §19.3.1.50 `p:transition`): read it, and set or
// clear it.
//
// A presentation application writes a transition with a duration as markup-compatibility
// content (ECMA-376 Part 3, §10.2.1): an `mc:Choice` that requires the 2010 extension namespace
// (`p14`) holds `p:transition` with `p14:dur`, and the `mc:Fallback` holds the same transition
// with only the speed (`spd`). A later kind can add another `mc:Choice` (`p159`). The read takes
// the duration from the `p14` branch and the kind from the first branch whose kind it knows.
//
// A change of the advance settings or the duration writes into every copy of the transition,
// so a kind that the editor does not write (and its extension markup) stays. A plain
// `p:transition` holds only a speed: a duration that is no speed wraps it in a new
// `mc:AlternateContent` whose `p14` choice holds the duration. A change of the
// kind replaces the whole transition markup with a new `mc:AlternateContent`, and `none`
// removes it. The new markup goes where `CT_Slide` puts the transition: after `p:clrMapOvr`
// and before `p:timing` and `p:extLst`.

import type { SlideTransitionState, TransitionKind } from '../../contracts/neutral.ts';
import { attr, elements, NS, specOf, xmlPartOf, type Element } from '../../store/index.ts';
import { timeOf } from './animation-model.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { specFromXml } from './shape-commands.ts';

const isP = (el: Element | undefined, local: string): boolean =>
  !!el && el.ns === NS.p && el.local === local;
const isAlternate = (el: Element): boolean => el.ns === NS.mc && el.local === 'AlternateContent';

/** The shortest and longest transition, and the longest advance time, in ms. */
export const MIN_TRANSITION_MS = 10;
export const MAX_TRANSITION_MS = 59_000;
export const MAX_ADVANCE_MS = 3_600_000;

// The tables are `Map`s: their keys come from the file and from commands, and an object would
// also answer for the keys of its prototype (`toString`, `constructor`, `__proto__`).

/** The duration of each speed (`ST_TransitionSpeed`), in ms. */
const SPEEDS: ReadonlyMap<string, number> = new Map([
  ['fast', 500],
  ['med', 750],
  ['slow', 1000],
]);

/** The kinds the editor writes, with their element. */
const KINDS: ReadonlyMap<string, string> = new Map<TransitionKind, string>([
  ['fade', '<p:fade/>'],
  ['push', '<p:push dir="u"/>'],
  ['wipe', '<p:wipe/>'],
  ['split', '<p:split orient="vert"/>'],
  ['cut', '<p:cut/>'],
]);

const NO_TRANSITION: SlideTransitionState = Object.freeze({
  kind: 'none',
  durationMs: 750,
  advanceOnClick: true,
  advanceAfterMs: null,
});

/** One copy of the transition: its element, and whether its branch reads `p14`. */
interface Copy {
  readonly el: Element;
  readonly p14: boolean;
}

/** Whether a branch requires the `p14` prefix, bound to the 2010 extension namespace. */
function requiresP14(branch: Element): boolean {
  const prefixes = (attr(branch, 'Requires') ?? '').split(/\s+/);
  return (
    branch.local === 'Choice' &&
    prefixes.includes('p14') &&
    branch.nsDecls.some((d) => d.prefix === 'p14' && d.uri === NS.p14)
  );
}

/** The transition markup of a slide root: the elements that hold it, and each copy. */
function transitionOf(root: Element | undefined): { holders: Element[]; copies: Copy[] } {
  const holders: Element[] = [];
  const copies: Copy[] = [];
  for (const c of elements(root)) {
    if (isP(c, 'transition')) {
      holders.push(c);
      copies.push({ el: c, p14: false });
    } else if (isAlternate(c)) {
      const own: Copy[] = [];
      for (const branch of elements(c))
        for (const t of elements(branch))
          if (isP(t, 'transition')) own.push({ el: t, p14: requiresP14(branch) });
      if (own.length) {
        holders.push(c);
        for (const copy of own) copies.push(copy);
      }
    }
  }
  return { holders, copies };
}

/** The kind of one copy. */
function kindOf(el: Element): SlideTransitionState['kind'] {
  const first = elements(el).find((c) => !(isP(c, 'sndAc') || isP(c, 'extLst')));
  if (!first) return 'none';
  if (first.ns === NS.p && KINDS.has(first.local)) return first.local as TransitionKind;
  return 'other';
}

/** The transition of a slide root. */
export function readTransition(root: Element | undefined): SlideTransitionState {
  const { copies } = transitionOf(root);
  if (copies.length === 0) return NO_TRANSITION;
  const main = copies.find((c) => c.p14) ?? copies[0];
  const p14dur = timeOf(attr(main.el, { ns: NS.p14, local: 'dur' }));
  const speed = SPEEDS.get(attr(main.el, 'spd') ?? 'fast') ?? 500;
  const advTm = timeOf(attr(main.el, 'advTm'));
  const advClick = attr(main.el, 'advClick');
  return {
    // The first copy, as an application reads the first branch it understands.
    kind: kindOf(copies[0].el),
    durationMs: p14dur ?? speed,
    advanceOnClick: advClick !== '0' && advClick !== 'false',
    advanceAfterMs: advTm ?? null,
  };
}

/** What `setTransition` changes; absent fields keep their values. */
export interface TransitionChange {
  readonly kind?: TransitionKind | 'none';
  readonly durationMs?: number;
  readonly advanceOnClick?: boolean;
  readonly advanceAfterMs?: number | null;
}

const speedOf = (ms: number): string => (ms <= 500 ? 'fast' : ms <= 750 ? 'med' : 'slow');

function checkMs(value: number, min: number, max: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max)
    throw new CommandRefusal('errors.unsupported', 'time');
  return Math.round(value);
}

/** The advance attributes of a new transition. */
function advanceAttrs(onClick: boolean, after: number | null): string {
  return (onClick ? '' : ' advClick="0"') + (after !== null ? ` advTm="${after}"` : '');
}

/** The markup of a new transition: a `p14` choice with the duration, and a fallback. */
function newTransitionXml(state: SlideTransitionState): string | null {
  const advance = advanceAttrs(state.advanceOnClick, state.advanceAfterMs);
  if (state.kind === 'none' || state.kind === 'other')
    return advance ? `<p:transition${advance}/>` : null;
  const body = KINDS.get(state.kind);
  if (body === undefined) return null;
  const spd = speedOf(state.durationMs);
  return (
    `<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice xmlns:p14="${NS.p14}" Requires="p14">` +
    `<p:transition spd="${spd}" p14:dur="${state.durationMs}"${advance}>${body}</p:transition></mc:Choice>` +
    `<mc:Fallback><p:transition spd="${spd}"${advance}>${body}</p:transition></mc:Fallback></mc:AlternateContent>`
  );
}

function setAttr(
  batch: OpBatch,
  part: string,
  el: Element,
  name: { ns: string; local: string; prefix?: string },
  value: string | null
): void {
  if ((attr(el, name) ?? null) === value) return;
  batch.push({ op: 'setAttribute', part, node: el.id, name, value });
}

/** Set or clear the transition of the slide part `part`. Refuses a change that changes nothing. */
export function setTransition(batch: OpBatch, part: string, change: TransitionChange): void {
  const root = xmlPartOf(batch.pkg, part)?.root;
  if (!root) throw new CommandRefusal('errors.malformed');
  if (change.kind !== undefined && change.kind !== 'none' && !KINDS.has(change.kind))
    throw new CommandRefusal('errors.unsupported', 'kind');
  const current = readTransition(root);
  const next: SlideTransitionState = {
    kind: change.kind ?? current.kind,
    durationMs:
      change.durationMs !== undefined
        ? checkMs(change.durationMs, MIN_TRANSITION_MS, MAX_TRANSITION_MS)
        : current.durationMs,
    advanceOnClick: change.advanceOnClick ?? current.advanceOnClick,
    advanceAfterMs:
      change.advanceAfterMs === undefined
        ? current.advanceAfterMs
        : change.advanceAfterMs === null
          ? null
          : checkMs(change.advanceAfterMs, 0, MAX_ADVANCE_MS),
  };
  if (typeof next.advanceOnClick !== 'boolean') throw new CommandRefusal('errors.unsupported');
  const { holders, copies } = transitionOf(root);
  const before = batch.ops.length;
  const sameKind = next.kind === current.kind && copies.length > 0;
  const plainKeepsDuration =
    sameKind &&
    change.durationMs !== undefined &&
    next.kind !== 'none' &&
    holders.length === 1 &&
    isP(holders[0], 'transition') &&
    SPEEDS.get(speedOf(next.durationMs)) !== next.durationMs;
  if (sameKind && !plainKeepsDuration) {
    // The same kind: write the settings into every copy, and keep the rest of the markup.
    for (const copy of copies) {
      setAttr(
        batch,
        part,
        copy.el,
        { ns: '', local: 'advClick' },
        next.advanceOnClick ? null : '0'
      );
      setAttr(
        batch,
        part,
        copy.el,
        { ns: '', local: 'advTm' },
        next.advanceAfterMs === null ? null : String(next.advanceAfterMs)
      );
      if (change.durationMs === undefined || next.kind === 'none') continue;
      setAttr(batch, part, copy.el, { ns: '', local: 'spd' }, speedOf(next.durationMs));
      const dur = { ns: NS.p14, local: 'dur', prefix: 'p14' };
      if (copy.p14 || attr(copy.el, dur) !== undefined)
        setAttr(batch, part, copy.el, dur, String(next.durationMs));
    }
  } else if (plainKeepsDuration) {
    // A plain `p:transition` holds a speed only: a duration that is no speed gets the `p14`
    // copy. The plain element moves into the fallback with its markup, and the choice is a
    // copy of it with the duration.
    wrapPlain(batch, part, root, copies[0].el, next);
  } else {
    for (const h of holders) batch.push({ op: 'removeNode', part, node: h.id });
    const xml = newTransitionXml(next);
    if (xml) {
      const kids = elements(root).filter((c) => !holders.includes(c));
      const at = kids.findIndex(
        (c) =>
          isP(c, 'timing') ||
          isP(c, 'extLst') ||
          (isAlternate(c) && elements(c).some((b) => elements(b).some((t) => isP(t, 'timing'))))
      );
      batch.push({
        op: 'insertNode',
        part,
        parent: root.id,
        index: at < 0 ? kids.length : at,
        spec: specFromXml(xml),
      });
    }
  }
  if (batch.ops.length === before) throw new CommandRefusal('disabled.noAnimationChange');
}

/**
 * Wrap the plain transition `el` in `mc:AlternateContent`: a `p14` choice with the duration and
 * the advance settings, and the fallback with the speed. Both copies keep the element's other
 * attributes and children.
 */
function wrapPlain(
  batch: OpBatch,
  part: string,
  root: Element,
  el: Element,
  next: SlideTransitionState
): void {
  const index = root.children.indexOf(el);
  const spd = speedOf(next.durationMs);
  const advance = { advClick: next.advanceOnClick ? null : '0', advTm: next.advanceAfterMs };
  const node = batch.insert({
    op: 'insertNode',
    part,
    parent: root.id,
    index,
    spec: specFromXml(
      `<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice xmlns:p14="${NS.p14}" Requires="p14"/>` +
        '<mc:Fallback/></mc:AlternateContent>'
    ),
  });
  const ac = batch.element(part, node);
  const [choice, fallback] = elements(ac);
  batch.push({
    op: 'insertNode',
    part,
    parent: choice.id,
    index: 0,
    spec: specOf(el),
  });
  batch.push({ op: 'moveNode', part, node: el.id, parent: fallback.id, index: 0 });
  const copy = elements(batch.element(part, choice.id))[0];
  for (const target of [copy, el]) {
    setAttr(batch, part, target, { ns: '', local: 'spd' }, spd);
    setAttr(batch, part, target, { ns: '', local: 'advClick' }, advance.advClick);
    const tm = advance.advTm === null ? null : String(advance.advTm);
    setAttr(batch, part, target, { ns: '', local: 'advTm' }, tm);
  }
  const dur = { ns: NS.p14, local: 'dur', prefix: 'p14' };
  setAttr(batch, part, copy, dur, String(next.durationMs));
}
