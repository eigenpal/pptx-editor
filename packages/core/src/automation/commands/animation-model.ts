// The effects of a slide, as the animation commands read them (ECMA-376 Part 1, §19.5
// `p:timing`).
//
// A presentation application writes the main sequence of a slide as a fixed tree:
//
//   p:timing/p:tnLst/p:par/p:cTn[@nodeType=tmRoot]/p:childTnLst
//     p:seq/p:cTn[@nodeType=mainSeq]/p:childTnLst
//       p:par (a click step: `p:cond delay="indefinite"`)/p:cTn/p:childTnLst
//         p:par (a time group: `p:cond delay="<start>"`)/p:cTn/p:childTnLst
//           p:par (an effect: `p:cTn/@presetClass`, `@nodeType` clickEffect, withEffect, or
//                  afterEffect)/p:cTn/p:childTnLst
//             the behaviors: p:set, p:anim, p:animEffect, p:animScale, p:animRot, …
//
// An interactive sequence (`p:seq` with `nodeType="interactiveSeq"`) has the same click step
// and time group levels under it. `readAnimations` lists the effects of every sequence, and the
// commands edit the main sequence only when it has exactly this shape. A slide whose `p:timing`
// is inside `mc:AlternateContent` is read-only: each branch would need the same edit.
//
// Every value is from the file and untrusted: numbers are read with bounds, names index only
// `Map` tables (never an object, whose prototype has keys of its own), the walks stop at the
// reader's depth cap, and nothing here throws on a malformed tree.

import type {
  AnimationClass,
  AnimationEffect,
  AnimationPreset,
  AnimationTrigger,
} from '../../contracts/neutral.ts';
import { MAX_XML_DEPTH } from '@eigenpal/ooxml-core';
import { contentElements } from '../../cascade/alternate-content.ts';
import { attr, child, elements, NS, type Element } from '../../store/index.ts';
import { presetOf } from './animation-presets.ts';
import { unsignedIntId } from './package-edits.ts';

/** The walks go as deep as the reader builds a tree. */
const MAX_DEPTH = MAX_XML_DEPTH;

const isP = (el: Element | undefined, local: string): boolean =>
  !!el && el.ns === NS.p && el.local === local;

/** The largest time the commands read or write, in ms: about 24 days, inside `xsd:int`. */
export const MAX_TIME_MS = 2_000_000_000;

/** A time attribute (`ST_TLTime`) in ms, or undefined for `indefinite` and malformed values. */
export function timeOf(raw: string | undefined): number | undefined {
  if (raw === undefined || !/^\d{1,10}$/.test(raw)) return undefined;
  const n = Number(raw);
  return n <= MAX_TIME_MS ? n : undefined;
}

/** A time node's `p:cTn`: its own, or the one in its `p:cBhvr` or `p:cMediaNode`. */
export const cTnOf = (node: Element): Element | undefined =>
  child(node, 'p:cTn') ??
  child(child(node, 'p:cBhvr'), 'p:cTn') ??
  child(child(node, 'p:cMediaNode'), 'p:cTn');

/** The `delay` of the first condition of a `p:cTn`'s start list, as written. */
export const startDelayOf = (cTn: Element | undefined): string | undefined =>
  attr(child(child(cTn, 'p:stCondLst'), 'p:cond'), 'delay');

/** The trigger of each effect `nodeType`. A `Map`: the key is from the file. */
const TRIGGERS: ReadonlyMap<string, AnimationTrigger> = new Map([
  ['clickEffect', 'onClick'],
  ['withEffect', 'withPrevious'],
  ['afterEffect', 'afterPrevious'],
]);

/** The trigger of an effect `p:cTn`, from its `nodeType`, or undefined. */
export const triggerOf = (cTn: Element | undefined): AnimationTrigger | undefined =>
  TRIGGERS.get(attr(cTn, 'nodeType') ?? '');

/** The `nodeType` that a trigger writes. */
export const NODE_TYPES: Readonly<Record<AnimationTrigger, string>> = {
  onClick: 'clickEffect',
  withPrevious: 'withEffect',
  afterPrevious: 'afterEffect',
};

/** The class of each `presetClass`. A `Map`: the key is from the file. */
const CLASSES: ReadonlyMap<string, AnimationClass> = new Map([
  ['entr', 'entrance'],
  ['emph', 'emphasis'],
  ['exit', 'exit'],
  ['path', 'motionPath'],
]);

/** Whether a `p:par` is an effect: its `p:cTn` names a preset class or an effect node type. */
export function isEffect(par: Element): boolean {
  if (!isP(par, 'par')) return false;
  const cTn = child(par, 'p:cTn');
  return attr(cTn, 'presetClass') !== undefined || triggerOf(cTn) !== undefined;
}

/** One effect of the main sequence, with the elements that hold it. */
export interface MainEffect {
  /** The effect `p:par`. */
  readonly par: Element;
  readonly cTn: Element;
  readonly trigger: AnimationTrigger;
}

/** The main sequence, when it has the shape that the commands write. */
export interface MainSequence {
  readonly seq: Element;
  /** The `p:childTnLst` of the sequence's `p:cTn`. */
  readonly list: Element;
  /** The `p:cTn/@id` of the sequence, which an automatic first step waits for. */
  readonly id: string;
  /** The click steps, each with its time groups, each with its effects. */
  readonly clicks: readonly {
    readonly par: Element;
    readonly groups: readonly { readonly par: Element; readonly effects: readonly MainEffect[] }[];
  }[];
  /** The effects in play order. */
  readonly effects: readonly MainEffect[];
}

/** What the animation commands know about the timing of a slide. */
export interface TimingModel {
  /** The slide's own `p:timing`. */
  readonly timing: Element | undefined;
  /** The `p:childTnLst` of the root time node (`tmRoot`). */
  readonly rootList: Element | undefined;
  /** The main sequence, when it is there and the commands can edit it. */
  readonly main: MainSequence | undefined;
  /** Whether the commands can edit the effects of this slide. */
  readonly editable: boolean;
}

/** The children of `el` under `p:cTn/p:childTnLst`, or undefined when it has no such list. */
function listOf(el: Element): Element | undefined {
  return child(child(el, 'p:cTn'), 'p:childTnLst');
}

/** The main sequence of a root time node list, or null when it is there but not editable. */
function mainOf(rootList: Element): MainSequence | null | undefined {
  const seq = elements(rootList).find(
    (s) => isP(s, 'seq') && attr(child(s, 'p:cTn'), 'nodeType') === 'mainSeq'
  );
  if (!seq) return undefined;
  // The id goes into new markup (an automatic first step waits for it): an `xsd:unsignedInt`.
  const num = unsignedIntId(attr(child(seq, 'p:cTn'), 'id'));
  const id = num === undefined ? undefined : String(num);
  const list = listOf(seq);
  if (!list || id === undefined) return null;
  const clicks: MainSequence['clicks'][number][] = [];
  const effects: MainEffect[] = [];
  for (const click of elements(list)) {
    const groupList = isP(click, 'par') ? listOf(click) : undefined;
    if (!groupList) return null;
    const groups: MainSequence['clicks'][number]['groups'][number][] = [];
    for (const group of elements(groupList)) {
      const effectList = isP(group, 'par') ? listOf(group) : undefined;
      if (!effectList) return null;
      const own: MainEffect[] = [];
      for (const par of elements(effectList)) {
        const cTn = child(par, 'p:cTn');
        const trigger = triggerOf(cTn);
        if (!isP(par, 'par') || !cTn || !trigger || attr(cTn, 'id') === undefined) return null;
        own.push({ par, cTn, trigger });
      }
      groups.push({ par: group, effects: own });
      for (const e of own) effects.push(e);
    }
    clicks.push({ par: click, groups });
  }
  return { seq, list, id, clicks, effects };
}

/** Whether a click step's start conditions include the slide start (`onBegin`). */
export const startsWithSlide = (click: Element): boolean =>
  elements(child(child(click, 'p:cTn'), 'p:stCondLst')).some((c) => attr(c, 'evt') === 'onBegin');

/** The timing of a slide root, as the commands read it. */
export function timingModel(root: Element | undefined): TimingModel {
  const kids = elements(root);
  const wrapped = kids.some(
    (c) =>
      c.ns === NS.mc &&
      c.local === 'AlternateContent' &&
      elements(c).some((b) => elements(b).some((t) => isP(t, 'timing')))
  );
  const timing = kids.find((c) => isP(c, 'timing'));
  const tnLst = child(timing, 'p:tnLst');
  const rootPar = elements(tnLst)[0];
  const rootList =
    rootPar && attr(child(rootPar, 'p:cTn'), 'nodeType') === 'tmRoot' ? listOf(rootPar) : undefined;
  const main = rootList ? mainOf(rootList) : undefined;
  // A `p:timing` without a root time node list can hold a build list alone: an edit makes one.
  const shapeOk = !tnLst || (rootList !== undefined && elements(tnLst).length === 1);
  return {
    timing,
    rootList,
    main: main ?? undefined,
    editable: !wrapped && shapeOk && main !== null,
  };
}

/** One shape that an effect can name: its node, by `p:cNvPr/@id`. */
interface Named {
  readonly nodeId: string;
  readonly name: string;
}

/** The shapes of a slide by `p:cNvPr/@id`: the shown copy of each, at any depth. */
export function shapesById(root: Element | undefined): Map<string, Named> {
  const out = new Map<string, Named>();
  const tree = child(child(root, 'p:cSld'), 'p:spTree');
  const walk = (parent: Element, depth: number): void => {
    if (depth > MAX_DEPTH) return;
    for (const el of contentElements(parent, root)) {
      const nv = elements(el).find((c) => c.local.startsWith('nv') && c.local.endsWith('Pr'));
      const cNvPr = child(nv, 'p:cNvPr');
      const id = attr(cNvPr, 'id');
      if (id !== undefined && !out.has(id))
        out.set(id, { nodeId: el.id, name: attr(cNvPr, 'name') ?? '' });
      if (isP(el, 'grpSp')) walk(el, depth + 1);
    }
  };
  if (tree) walk(tree, 0);
  return out;
}

/** The first shape target (`p:spTgt`) of an effect's behaviors. */
function targetOf(effect: Element): Element | undefined {
  let hit: Element | undefined;
  const walk = (el: Element, depth: number): void => {
    if (hit || depth > MAX_DEPTH) return;
    if (isP(el, 'spTgt') && attr(el, 'spid') !== undefined) return void (hit = el);
    for (const c of elements(el)) if (!isP(c, 'stCondLst')) walk(c, depth + 1);
  };
  walk(effect, 0);
  return hit;
}

/** The `spid` an effect animates, or undefined. */
export const effectSpid = (effect: Element): string | undefined => attr(targetOf(effect), 'spid');

/** The behaviors of an effect: the time nodes in its `p:childTnLst`. */
export const behaviorsOf = (effect: Element): Element[] => elements(listOf(effect));

/** How long a behavior plays, from the start of its effect, in ms: delay + duration × repeats. */
function behaviorEnd(behavior: Element): { end: number; dur: number } | undefined {
  const cTn = cTnOf(behavior);
  const dur = timeOf(attr(cTn, 'dur'));
  if (dur === undefined) return undefined;
  const delay = timeOf(startDelayOf(cTn)) ?? 0;
  const repeat = timeOf(attr(cTn, 'repeatCount'));
  const times = (repeat !== undefined ? repeat / 1000 : 1) * (attr(cTn, 'autoRev') === '1' ? 2 : 1);
  return { end: delay + dur * times, dur };
}

/**
 * How long an effect plays, in ms: the end of its last behavior. Null for an instant effect,
 * whose behaviors all last 1 ms or less (Appear, Disappear).
 */
export function effectDuration(effect: Element): number | null {
  let end = 0;
  let timed = false;
  for (const b of behaviorsOf(effect)) {
    const t = behaviorEnd(b);
    if (!t) continue;
    if (t.dur > 1) timed = true;
    end = Math.max(end, t.end);
  }
  return timed ? Math.round(end) : null;
}

/** An integer attribute in a small range, or null. */
function smallInt(raw: string | undefined): number | null {
  return raw !== undefined && /^\d{1,9}$/.test(raw) ? Number(raw) : null;
}

/** The trigger shape of an interactive sequence: the `spid` of its start condition. */
function triggerShapeOf(seq: Element): string | undefined {
  for (const cond of elements(child(child(seq, 'p:cTn'), 'p:stCondLst'))) {
    const spid = attr(child(child(cond, 'p:tgtEl'), 'p:spTgt'), 'spid');
    if (spid !== undefined) return spid;
  }
  return undefined;
}

/** The effects of a slide: the main sequence first, then each interactive sequence. */
export function readAnimations(root: Element | undefined): {
  readonly effects: AnimationEffect[];
  readonly editable: boolean;
} {
  const model = timingModel(root);
  const shapes = shapesById(root);
  const effects: AnimationEffect[] = [];
  const add = (
    par: Element,
    sequence: 'main' | 'interactive',
    step: number | null,
    editable: boolean,
    triggerShapeId?: string
  ): void => {
    const cTn = child(par, 'p:cTn');
    const target = targetOf(par);
    const shapeId = attr(target, 'spid') ?? '';
    const shape = shapes.get(shapeId);
    const presetClass = attr(cTn, 'presetClass');
    const presetId = smallInt(attr(cTn, 'presetID'));
    const effectClass = CLASSES.get(presetClass ?? '') ?? 'other';
    const preset: AnimationPreset | null = presetOf(effectClass, presetId);
    const pRg = child(child(target, 'p:txEl'), 'p:pRg');
    const st = smallInt(attr(pRg, 'st'));
    const end = smallInt(attr(pRg, 'end'));
    effects.push({
      id: attr(cTn, 'id') ?? '',
      shapeId,
      nodeId: shape?.nodeId ?? null,
      shapeName: shape?.name ?? '',
      effectClass,
      preset,
      presetId,
      trigger: triggerOf(cTn) ?? 'withPrevious',
      sequence,
      ...(triggerShapeId !== undefined
        ? { triggerShapeId, triggerShapeName: shapes.get(triggerShapeId)?.name ?? '' }
        : {}),
      step,
      delayMs: timeOf(startDelayOf(cTn)) ?? 0,
      durationMs: effectDuration(par),
      ...(st !== null && end !== null ? { paragraphs: { start: st, end } } : {}),
      editable,
    });
  };
  if (model.main) {
    // A click step waits for a click unless it starts with the slide: its conditions, not the
    // triggers of its effects, say when it plays.
    let step = 0;
    for (const click of model.main.clicks) {
      if (!startsWithSlide(click.par)) step += 1;
      for (const group of click.groups)
        for (const e of group.effects) add(e.par, 'main', step, model.editable);
    }
  }
  // The other sequences, and a main sequence that the commands cannot edit: every effect, in
  // document order.
  const collect = (
    el: Element,
    depth: number,
    sequence: 'main' | 'interactive',
    trigger?: string
  ): void => {
    if (depth > MAX_DEPTH) return;
    for (const c of elements(el)) {
      if (isEffect(c)) add(c, sequence, null, false, trigger);
      else if (isP(c, 'seq')) {
        const type = attr(child(c, 'p:cTn'), 'nodeType');
        if (type === 'mainSeq' && model.main) continue;
        collect(
          c,
          depth + 1,
          type === 'mainSeq' ? 'main' : 'interactive',
          type === 'mainSeq' ? undefined : triggerShapeOf(c)
        );
      } else collect(c, depth + 1, sequence, trigger);
    }
  };
  const timings: Element[] = [];
  for (const c of elements(root)) {
    if (isP(c, 'timing')) timings.push(c);
    else if (c.ns === NS.mc && c.local === 'AlternateContent') {
      // The first branch that holds a `p:timing`: the branches are copies of one timing.
      const branch = elements(c).find((b) => elements(b).some((t) => isP(t, 'timing')));
      const t = elements(branch).find((x) => isP(x, 'timing'));
      if (t) timings.push(t);
    }
  }
  for (const t of timings) collect(child(t, 'p:tnLst') ?? t, 0, 'main');
  return { effects, editable: model.editable };
}
