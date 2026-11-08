// The writer's repair of a slide's transition (ECMA-376 Part 1, §19.3.1.50 `p:transition`).
//
// `CT_Slide` holds at most one transition, before the timing: cSld, clrMapOvr, transition,
// timing, extLst. A transition is a `p:transition`, or an `mc:AlternateContent` whose branches
// hold one (ECMA-376 Part 3, §10.2.1). Two collaborators who each set the transition of one
// slide at the same time both insert one, so the merge can hold two, and one after the timing.
// The save keeps the first in document order, before the timing, and drops the others. A slide
// with one transition in its place saves unchanged, and so does a slide whose transitions are
// all as the file wrote them (`fileTransitionsOnly`): the repair is for what the session wrote.

import { NS } from './namespaces.ts';
import { elements } from './query.ts';
import { fileTransitionsOnly } from './timing.ts';
import type { Element } from './tree.ts';

const isP = (el: Element, local: string): boolean => el.ns === NS.p && el.local === local;
const isAlternate = (el: Element): boolean => el.ns === NS.mc && el.local === 'AlternateContent';

/** Whether a child of the slide root is a transition, or wraps one. */
const holds = (el: Element, local: string): boolean =>
  isP(el, local) ||
  (isAlternate(el) && elements(el).some((b) => elements(b).some((c) => isP(c, local))));

/** A slide root with at most one transition, before the timing. The root itself when it has. */
export function repairTransitionRoot(root: Element): Element {
  const kids = elements(root);
  const transitions = kids.filter((c) => holds(c, 'transition'));
  if (transitions.length === 0 || fileTransitionsOnly(root)) return root;
  const timing = kids.findIndex((c) => holds(c, 'timing') || isP(c, 'extLst'));
  const first = kids.indexOf(transitions[0]);
  if (transitions.length === 1 && (timing < 0 || first < timing)) return root;
  const keep = transitions[0];
  const children = root.children.filter((c) => c.type !== 'element' || !transitions.includes(c));
  // The place before the timing (or the extension list), else the end.
  const at = children.findIndex(
    (c) => c.type === 'element' && (holds(c, 'timing') || isP(c, 'extLst'))
  );
  children.splice(at < 0 ? children.length : at, 0, keep);
  return { ...root, children };
}
