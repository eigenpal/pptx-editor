// The animation commands: add, remove, reorder, and time the effects of a slide's main sequence
// (ECMA-376 Part 1, §19.5). `animation-model.ts` reads the tree they edit.
//
// The main sequence is a flat list of effects, each with its trigger (`p:cTn/@nodeType`). The
// click steps and time groups around the effects follow from that list: an effect on a click
// (and the first effect) starts a click step, an effect after the previous one starts a time
// group in the step, and an effect with the previous one joins the group. A time group starts
// when the group before it ends: the latest end of its effects (delay + duration). After each
// edit, `relayout` makes the steps and groups follow the list again. When they already do, it
// only corrects the group delays; otherwise it writes new steps and groups (reusing the old
// `p:cTn/@id` values) and moves each effect `p:par` into them, so an effect keeps its own
// markup, unknown content included.
//
// Each new `p:cTn/@id` is unique on the slide; in a collaboration session it comes from the
// actor's stripe, as shape ids do. A new effect takes the lowest `grpId` that no effect or build
// entry of its shape uses, and a shape with text gets a build entry (`p:bldP`, §19.5.12).

import type { AnimationPreset, AnimationTrigger } from '../../contracts/neutral.ts';
import {
  attr,
  child,
  elements,
  nodeIndex,
  NS,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import { MAX_DECIMAL_ID } from './actor-scoped-ids.ts';
import {
  cTnOf,
  behaviorsOf,
  effectDuration,
  effectSpid,
  isEffect,
  MAX_TIME_MS,
  NODE_TYPES,
  startDelayOf,
  timeOf,
  timingModel,
  type MainEffect,
  type MainSequence,
} from './animation-model.ts';
import { ANIMATION_PRESETS, effectXml, PRESET_IDS } from './animation-presets.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { freeIds, stripedIds, unsignedIntId } from './package-edits.ts';
import { specFromXml } from './shape-commands.ts';

const MAX_DEPTH = 64;
/** The longest delay a command writes, in ms: one hour. */
export const MAX_DELAY_MS = 3_600_000;
/** The shortest and longest duration a command writes, in ms. */
export const MIN_DURATION_MS = 10;
export const MAX_DURATION_MS = 59_000;

const isP = (el: Element | undefined, local: string): boolean =>
  !!el && el.ns === NS.p && el.local === local;

const TRIGGERS: ReadonlySet<string> = new Set(['onClick', 'withPrevious', 'afterPrevious']);

function rootOf(batch: OpBatch, part: string): Element {
  const root = xmlPartOf(batch.pkg, part)?.root;
  if (!root) throw new CommandRefusal('errors.malformed');
  return root;
}

/** The timing of a slide, refused when the commands cannot edit it. */
function editableModel(batch: OpBatch, part: string) {
  const model = timingModel(rootOf(batch, part));
  if (!model.editable) throw new CommandRefusal('disabled.animationsReadOnly');
  return model;
}

/** A millisecond value from a command: a finite number in range, rounded. */
function ms(value: number, min: number, max: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max)
    throw new CommandRefusal('errors.unsupported', 'time');
  return Math.round(value);
}

/** Every element of the slide's timing markup (its own `p:timing` and wrapped copies). */
function timingTrees(root: Element): Element[] {
  return elements(root).filter(
    (c) => isP(c, 'timing') || (c.ns === NS.mc && c.local === 'AlternateContent')
  );
}

/** Visit every element under `roots`, depth-bounded. */
function visit(roots: readonly Element[], fn: (el: Element) => void): void {
  const walk = (el: Element, depth: number): void => {
    if (depth > MAX_DEPTH) return;
    fn(el);
    for (const c of elements(el)) walk(c, depth + 1);
  };
  for (const r of roots) walk(r, 0);
}

/** A source of new `p:cTn/@id` values, unique among the slide's time nodes. */
function timeNodeIds(batch: OpBatch, root: Element): () => number {
  const used = new Set<number>();
  visit(timingTrees(root), (el) => {
    if (!isP(el, 'cTn')) return;
    const id = unsignedIntId(attr(el, 'id'));
    if (id !== undefined) used.add(id);
  });
  if (batch.allocationActor)
    return stripedIds(batch.allocationActor, used, 1, MAX_DECIMAL_ID, batch.mintedOf('timeNodeId'));
  return freeIds(used, 1, MAX_DECIMAL_ID);
}

/** The effect `p:par` of the main sequence with this `p:cTn/@id`, or a refusal. */
function mainEffect(main: MainSequence | undefined, effectId: string): MainEffect {
  const hit = main?.effects.find((e) => attr(e.cTn, 'id') === effectId);
  if (!hit) throw new CommandRefusal('disabled.noAnimation');
  return hit;
}

// ---------------------------------------------------------------- steps and groups

/** The click steps and time groups that a list of effects plays in. */
function plan(effects: readonly { el: Element; trigger: AnimationTrigger }[]): Element[][][] {
  const clicks: Element[][][] = [];
  effects.forEach((e, i) => {
    if (i === 0 || e.trigger === 'onClick') clicks.push([[e.el]]);
    else if (e.trigger === 'afterPrevious') clicks.at(-1)!.push([e.el]);
    else clicks.at(-1)!.at(-1)!.push(e.el);
  });
  return clicks;
}

/** When a time group ends, from its start: the latest end of its effects. */
function groupEnd(effects: readonly Element[]): number {
  let end = 0;
  for (const e of effects) {
    const delay = timeOf(startDelayOf(child(e, 'p:cTn'))) ?? 0;
    end = Math.max(end, delay + (effectDuration(e) ?? 0));
  }
  return end;
}

/** The start delays of the groups of each step. */
function groupStarts(clicks: readonly (readonly Element[][])[]): number[][] {
  return clicks.map((groups) => {
    let at = 0;
    return groups.map((g, i) => {
      if (i > 0) at = Math.min(MAX_TIME_MS, at + groupEnd(groups[i - 1]));
      return at;
    });
  });
}

/** The conditions of a click step: a click, and for an automatic first step, the slide start. */
const clickConditions = (auto: boolean, seqId: string): string =>
  '<p:cond delay="indefinite"/>' +
  (auto ? `<p:cond evt="onBegin" delay="0"><p:tn val="${seqId}"/></p:cond>` : '');

/** Whether a click step's start conditions include the slide start (`onBegin`). */
const startsWithSlide = (click: Element): boolean =>
  elements(child(child(click, 'p:cTn'), 'p:stCondLst')).some((c) => attr(c, 'evt') === 'onBegin');

function setAttr(batch: OpBatch, part: string, el: Element, local: string, value: string | null) {
  if ((attr(el, local) ?? null) === value) return;
  batch.push({ op: 'setAttribute', part, node: el.id, name: { ns: '', local }, value });
}

/**
 * Make the steps and groups of the main sequence follow its effects (see the file comment).
 * `order` is the effect list in its new order; absent, the order of the tree.
 */
function relayout(batch: OpBatch, part: string, order?: readonly Element[]): void {
  const root = rootOf(batch, part);
  const model = timingModel(root);
  const main = model.main;
  if (!main) return;
  const byId = new Map(main.effects.map((e) => [e.par.id, e]));
  const list = (order ?? main.effects.map((e) => e.par)).map((el) => byId.get(el.id)!);
  if (list.length === 0) return removeEmptyTiming(batch, part, main);
  const clicks = plan(list.map((e) => ({ el: e.par, trigger: e.trigger })));
  const starts = groupStarts(clicks);
  const auto = list[0].trigger !== 'onClick';
  const same =
    clicks.length === main.clicks.length &&
    clicks.every(
      (groups, c) =>
        groups.length === main.clicks[c].groups.length &&
        groups.every(
          (g, k) =>
            g.length === main.clicks[c].groups[k].effects.length &&
            g.every((el, j) => el.id === main.clicks[c].groups[k].effects[j].par.id)
        )
    );
  if (same) {
    main.clicks.forEach((click, c) => {
      click.groups.forEach((g, k) => {
        const cond = child(child(child(g.par, 'p:cTn'), 'p:stCondLst'), 'p:cond');
        if (cond) setAttr(batch, part, cond, 'delay', String(starts[c][k]));
      });
      const want = c === 0 && auto;
      if (startsWithSlide(click.par) === want) return;
      const conds = child(child(click.par, 'p:cTn'), 'p:stCondLst');
      if (!conds) return;
      if (want)
        batch.push({
          op: 'insertNode',
          part,
          parent: conds.id,
          index: elements(conds).length,
          spec: specFromXml(`<p:cond evt="onBegin" delay="0"><p:tn val="${main.id}"/></p:cond>`),
        });
      else
        for (const cond of elements(conds))
          if (attr(cond, 'evt') === 'onBegin')
            batch.push({ op: 'removeNode', part, node: cond.id });
    });
    return;
  }
  // New steps and groups, with the old wrapper ids first. In a collaboration session every id
  // is new, from the actor's stripe: two peers that rebuild at the same time cannot reuse one id.
  const pool: string[] = [];
  for (const click of batch.allocationActor ? [] : main.clicks) {
    // Old ids go into new markup: only the ones that are `xsd:unsignedInt` numbers.
    const id = unsignedIntId(attr(child(click.par, 'p:cTn'), 'id'));
    if (id !== undefined) pool.push(String(id));
    for (const g of click.groups) {
      const gid = unsignedIntId(attr(child(g.par, 'p:cTn'), 'id'));
      if (gid !== undefined) pool.push(String(gid));
    }
  }
  const mint = timeNodeIds(batch, root);
  let taken = 0;
  const nextId = (): string => (taken < pool.length ? pool[taken++] : String(mint()));
  const old = elements(main.list);
  clicks.forEach((groups, c) => {
    const clickId = nextId();
    const inner = groups
      .map(
        (_, k) =>
          `<p:par><p:cTn id="${nextId()}" fill="hold"><p:stCondLst><p:cond delay="${starts[c][k]}"/></p:stCondLst><p:childTnLst/></p:cTn></p:par>`
      )
      .join('');
    const xml =
      `<p:par><p:cTn id="${clickId}" fill="hold"><p:stCondLst>${clickConditions(c === 0 && auto, main.id)}</p:stCondLst>` +
      `<p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`;
    const clickNode = batch.insert({
      op: 'insertNode',
      part,
      parent: main.list.id,
      index: old.length + c,
      spec: specFromXml(xml),
    });
    const groupPars = elements(
      child(child(batch.element(part, clickNode), 'p:cTn'), 'p:childTnLst')
    );
    groups.forEach((g, k) => {
      const target = child(child(groupPars[k], 'p:cTn'), 'p:childTnLst')!;
      g.forEach((el, j) =>
        batch.push({ op: 'moveNode', part, node: el.id, parent: target.id, index: j })
      );
    });
  });
  for (const el of old) batch.push({ op: 'removeNode', part, node: el.id });
}

/**
 * The main sequence has no effect left: remove it, and the lists and the `p:timing` that it
 * leaves empty (`CT_TimeNodeList` and `CT_BuildList` need one child each).
 */
function removeEmptyTiming(batch: OpBatch, part: string, main: MainSequence): void {
  batch.push({ op: 'removeNode', part, node: main.seq.id });
  const model = timingModel(rootOf(batch, part));
  const timing = model.timing;
  if (!timing || (model.rootList && elements(model.rootList).length > 0)) return;
  const tnLst = child(timing, 'p:tnLst');
  const bldLst = child(timing, 'p:bldLst');
  const builds = elements(bldLst).length > 0;
  if (!builds) {
    batch.push({ op: 'removeNode', part, node: timing.id });
    return;
  }
  if (tnLst) batch.push({ op: 'removeNode', part, node: tnLst.id });
}

// ---------------------------------------------------------------- the commands

/**
 * The shape an effect on node `nodeId` animates: the top-level element of the shape tree that
 * holds it (a group animates as a whole), with its `p:cNvPr/@id`.
 */
function animatedShape(
  batch: OpBatch,
  part: string,
  nodeId: string
): { el: Element; spid: string } {
  const xml = xmlPartOf(batch.pkg, part);
  if (!xml) throw new CommandRefusal('errors.malformed');
  const index = nodeIndex(xml);
  const parentOf = (id: string | null | undefined) => (id ? index.get(id) : undefined);
  let at = index.get(nodeId);
  let top: Element | undefined;
  for (let i = 0; at && i < MAX_DEPTH && !top; i += 1) {
    const up = parentOf(at.parent);
    const parent = up?.node;
    if (!up || parent?.type !== 'element' || at.node.type !== 'element') break;
    const branch =
      parent.ns === NS.mc && (parent.local === 'Choice' || parent.local === 'Fallback');
    // A shape in a branch of `mc:AlternateContent` stands for the wrapper's place.
    const holder = branch ? parentOf(parentOf(up.parent)?.parent)?.node : parent;
    if (holder?.type === 'element' && isP(holder, 'spTree')) top = at.node;
    else at = branch ? parentOf(up.parent) : up;
  }
  const nv = elements(top).find((c) => c.local.startsWith('nv') && c.local.endsWith('Pr'));
  const spid = unsignedIntId(attr(child(nv, 'p:cNvPr'), 'id'));
  if (!top || spid === undefined) throw new CommandRefusal('disabled.unsupportedElement');
  return { el: top, spid: String(spid) };
}

/** The `grpId` values that the effects and build entries of shape `spid` use. */
function usedGroupIds(root: Element, spid: string): Set<number> {
  const used = new Set<number>();
  visit(timingTrees(root), (el) => {
    const grp = unsignedIntId(attr(el, 'grpId'));
    if (grp === undefined) return;
    const owner = isP(el, 'cTn') ? undefined : attr(el, 'spid');
    if (owner === spid) used.add(grp);
  });
  // An effect's `grpId` is on its `p:cTn`: read its target from the effect `p:par`.
  visit(timingTrees(root), (el) => {
    if (!isEffect(el) || effectSpid(el) !== spid) return;
    const grp = unsignedIntId(attr(child(el, 'p:cTn'), 'grpId'));
    if (grp !== undefined) used.add(grp);
  });
  return used;
}

/** A new `p:timing` with an empty main sequence, its ids from `nextId`. */
function newTimingXml(nextId: () => number): string {
  const rootId = nextId();
  const seqId = nextId();
  return (
    `<p:timing><p:tnLst><p:par><p:cTn id="${rootId}" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
    newMainXml(seqId) +
    '</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
  );
}

function newMainXml(seqId: number): string {
  return (
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${seqId}" dur="indefinite" nodeType="mainSeq"><p:childTnLst/></p:cTn>` +
    '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
    '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
  );
}

/** Give the slide a main sequence, and the `p:timing` and lists around it, when it has none. */
function ensureMain(batch: OpBatch, part: string): MainSequence {
  let root = rootOf(batch, part);
  const model = timingModel(root);
  if (model.main) return model.main;
  const ids = timeNodeIds(batch, root);
  if (!model.timing) {
    // `CT_Slide`: cSld, clrMapOvr, transition, timing, extLst. Markup-compatibility content
    // before `p:extLst` (a wrapped transition) stays before the timing.
    const kids = elements(root);
    const ext = kids.findIndex((c) => isP(c, 'extLst'));
    batch.push({
      op: 'insertNode',
      part,
      parent: root.id,
      index: ext < 0 ? kids.length : ext,
      spec: specFromXml(newTimingXml(ids)),
    });
  } else if (!child(model.timing, 'p:tnLst')) {
    const tn = `<p:tnLst><p:par><p:cTn id="${ids()}" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${newMainXml(ids())}</p:childTnLst></p:cTn></p:par></p:tnLst>`;
    batch.push({
      op: 'insertNode',
      part,
      parent: model.timing.id,
      index: 0,
      spec: specFromXml(tn),
    });
  } else if (model.rootList) {
    batch.push({
      op: 'insertNode',
      part,
      parent: model.rootList.id,
      index: 0,
      spec: specFromXml(newMainXml(ids())),
    });
  }
  root = rootOf(batch, part);
  const main = timingModel(root).main;
  if (!main) throw new CommandRefusal('disabled.animationsReadOnly');
  return main;
}

/** Add a build entry for shape `spid` and `grpId` (§19.5.12 `p:bldP`). */
function addBuild(batch: OpBatch, part: string, spid: string, grpId: number): void {
  const timing = timingModel(rootOf(batch, part)).timing;
  if (!timing) return;
  const entry = `<p:bldP spid="${spid}" grpId="${grpId}" animBg="1"/>`;
  const bldLst = child(timing, 'p:bldLst');
  if (bldLst) {
    batch.push({
      op: 'insertNode',
      part,
      parent: bldLst.id,
      index: elements(bldLst).length,
      spec: specFromXml(entry),
    });
    return;
  }
  const kids = elements(timing);
  const at = kids.findIndex((c) => isP(c, 'extLst'));
  batch.push({
    op: 'insertNode',
    part,
    parent: timing.id,
    index: at < 0 ? kids.length : at,
    spec: specFromXml(`<p:bldLst>${entry}</p:bldLst>`),
  });
}

/**
 * Add an effect of `preset` to each shape of `nodeIds`, at the end of the main sequence. The
 * first takes `trigger`, the others play with it. Returns the new effect ids.
 */
export function addAnimation(
  batch: OpBatch,
  part: string,
  nodeIds: readonly string[],
  preset: AnimationPreset,
  trigger: AnimationTrigger = 'onClick'
): string[] {
  if (!PRESET_IDS.has(preset)) throw new CommandRefusal('errors.unsupported', 'preset');
  if (!TRIGGERS.has(trigger)) throw new CommandRefusal('errors.unsupported', 'trigger');
  editableModel(batch, part);
  const shapes = new Map<string, Element>();
  for (const id of nodeIds) {
    const s = animatedShape(batch, part, id);
    if (!shapes.has(s.spid)) shapes.set(s.spid, s.el);
  }
  if (shapes.size === 0) throw new CommandRefusal('disabled.noSelection');
  const main = ensureMain(batch, part);
  const root = rootOf(batch, part);
  const nextId = timeNodeIds(batch, root);
  const added: string[] = [];
  let first = true;
  const xmls: string[] = [];
  const builds: [string, number][] = [];
  for (const [spid, el] of shapes) {
    const used = usedGroupIds(root, spid);
    let grpId = 0;
    while (used.has(grpId)) grpId += 1;
    const xml = effectXml(preset, {
      nextId,
      spid,
      grpId,
      nodeType: NODE_TYPES[first ? trigger : 'withPrevious'],
      delayMs: 0,
    });
    added.push(/<p:cTn id="(\d+)" presetID/.exec(xml)![1]);
    xmls.push(xml);
    if (isP(el, 'sp') && child(el, 'p:txBody')) builds.push([spid, grpId]);
    first = false;
  }
  // A new click step at the end holds the effects; `relayout` puts them in their place.
  batch.push({
    op: 'insertNode',
    part,
    parent: main.list.id,
    index: elements(main.list).length,
    spec: specFromXml(
      `<p:par><p:cTn id="${nextId()}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>` +
        `<p:par><p:cTn id="${nextId()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${xmls.join('')}</p:childTnLst></p:cTn></p:par>` +
        '</p:childTnLst></p:cTn></p:par>'
    ),
  });
  for (const [spid, grpId] of builds) addBuild(batch, part, spid, grpId);
  relayout(batch, part);
  return added;
}

/** Remove effect `effectId`, and its build entry when no other effect of its group is left. */
export function removeAnimation(batch: OpBatch, part: string, effectId: string): void {
  const effect = mainEffect(editableModel(batch, part).main, effectId);
  const spid = effectSpid(effect.par);
  const grpId = attr(effect.cTn, 'grpId');
  batch.push({ op: 'removeNode', part, node: effect.par.id });
  const root = rootOf(batch, part);
  if (spid !== undefined && grpId !== undefined) {
    let still = false;
    visit(timingTrees(root), (el) => {
      if (
        !still &&
        isEffect(el) &&
        effectSpid(el) === spid &&
        attr(child(el, 'p:cTn'), 'grpId') === grpId
      )
        still = true;
    });
    const timing = timingModel(root).timing;
    const bldLst = child(timing, 'p:bldLst');
    if (!still && bldLst) {
      const gone = elements(bldLst).filter(
        (b) => attr(b, 'spid') === spid && attr(b, 'grpId') === grpId
      );
      const all = gone.length === elements(bldLst).length;
      for (const b of all ? [bldLst] : gone) batch.push({ op: 'removeNode', part, node: b.id });
    }
  }
  relayout(batch, part);
}

/** Move effect `effectId` to position `to` of the main sequence. */
export function moveAnimation(batch: OpBatch, part: string, effectId: string, to: number): void {
  const main = editableModel(batch, part).main;
  const effect = mainEffect(main, effectId);
  if (!Number.isInteger(to)) throw new CommandRefusal('errors.unsupported', 'index');
  const order = main!.effects.map((e) => e.par);
  const from = order.indexOf(effect.par);
  const at = Math.max(0, Math.min(order.length - 1, to));
  if (at === from) throw new CommandRefusal('disabled.noAnimationChange');
  order.splice(from, 1);
  order.splice(at, 0, effect.par);
  relayout(batch, part, order);
}

/** Change the trigger, delay, or duration of effect `effectId`. */
export function setAnimationTiming(
  batch: OpBatch,
  part: string,
  effectId: string,
  change: {
    readonly trigger?: AnimationTrigger;
    readonly delayMs?: number;
    readonly durationMs?: number;
  }
): void {
  const effect = mainEffect(editableModel(batch, part).main, effectId);
  const before = batch.ops.length;
  if (change.trigger !== undefined) {
    if (!TRIGGERS.has(change.trigger)) throw new CommandRefusal('errors.unsupported', 'trigger');
    setAttr(batch, part, effect.cTn, 'nodeType', NODE_TYPES[change.trigger]);
  }
  if (change.delayMs !== undefined) {
    const delay = String(ms(change.delayMs, 0, MAX_DELAY_MS));
    const conds = child(effect.cTn, 'p:stCondLst');
    const cond = child(conds, 'p:cond');
    if (cond) setAttr(batch, part, cond, 'delay', delay);
    else if (conds)
      batch.push({
        op: 'insertNode',
        part,
        parent: conds.id,
        index: 0,
        spec: specFromXml(`<p:cond delay="${delay}"/>`),
      });
    else
      batch.push({
        op: 'insertNode',
        part,
        parent: effect.cTn.id,
        index: 0,
        spec: specFromXml(`<p:stCondLst><p:cond delay="${delay}"/></p:stCondLst>`),
      });
  }
  if (change.durationMs !== undefined) {
    const target = ms(change.durationMs, MIN_DURATION_MS, MAX_DURATION_MS);
    const current = effectDuration(effect.par);
    if (current === null || current <= 0) throw new CommandRefusal('disabled.instantAnimation');
    scaleBehaviors(batch, part, effect.par, target / current);
  }
  if (batch.ops.length === before) throw new CommandRefusal('disabled.noAnimationChange');
  relayout(batch, part);
}

/** Scale the durations and delays of an effect's behaviors by `factor`. */
function scaleBehaviors(batch: OpBatch, part: string, effect: Element, factor: number): void {
  for (const behavior of behaviorsOf(effect)) {
    const cTn = cTnOf(behavior);
    if (!cTn) continue;
    const dur = timeOf(attr(cTn, 'dur'));
    if (dur === undefined) continue;
    const timed = dur > 1;
    if (timed) setAttr(batch, part, cTn, 'dur', String(Math.max(1, Math.round(dur * factor))));
    const cond = child(child(cTn, 'p:stCondLst'), 'p:cond');
    const delay = timeOf(attr(cond, 'delay'));
    if (!cond || delay === undefined || delay === 0) continue;
    // An instant behavior at the end of the effect (the hide of an exit) stays at its end.
    const next = timed
      ? Math.round(delay * factor)
      : Math.max(0, Math.round((delay + dur) * factor) - dur);
    setAttr(batch, part, cond, 'delay', String(Math.min(MAX_TIME_MS, next)));
  }
}

/** The default duration of a preset, in ms, or null for an instant one. @internal */
export const presetDuration = (preset: AnimationPreset): number | null =>
  ANIMATION_PRESETS[preset].duration;
