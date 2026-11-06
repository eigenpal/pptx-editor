// The answers of `can` and the formatting state for toggles, kept until the state a command
// compiles against moves. Toolbars ask for every slot on every change, so each answer is
// computed once per state. Moved out of `pptx-editor.ts` (the 1000-line cap).
//
// A few commands edit the deck's structure and never read a slide's content: New Slide and
// Add Section. Their probe runs the whole command on a copy, which costs more as the deck grows.
// Their answers have a key of their own (`structure`), without the package or the selection,
// so typing on a slide keeps them. New Slide reads only the layout of the current slide, not
// its index (`structureSlideKey`), so moving between slides on one layout keeps its answer.

import type { CanResult, EditorCommand } from '../contracts/index.ts';
import { newSlideLayout } from '../automation/commands/slide-layouts.ts';
import { changedPartNames, presentationInfo, type PptxPackage } from '../store/index.ts';
import type { DocumentSession } from './document-session.ts';
import { replicaBindingOf } from './replica-ids.ts';

/**
 * Commands whose `can` answer reads no slide content and no shape or text selection: only the
 * presentation part, the relationship parts, the layouts and masters, the content types, the
 * part names, the current slide index, and the mode. @internal
 */
export const STRUCTURE_COMMANDS: ReadonlySet<EditorCommand['type']> = new Set([
  'addSlide',
  'addSection',
]);

/** `can` answers and one derived value, both reset when `state()` changes. @internal */
export class CanCache<Active> {
  readonly #answers = new Map<string, CanResult>();
  readonly #structural = new Map<string, CanResult>();
  #state: readonly unknown[] = [];
  #structure: readonly unknown[] = [];
  #active: Active | null = null;

  /**
   * `state` is everything an answer may read; `structure` is what the answers of
   * {@link STRUCTURE_COMMANDS} read.
   */
  constructor(
    private readonly state: () => readonly unknown[],
    private readonly structure: () => readonly unknown[] = state
  ) {}

  /** Drop every answer when the state moved since the last call. */
  sync(): void {
    const now = this.state();
    if (same(now, this.#state)) return;
    this.#state = now;
    this.#answers.clear();
    this.#active = null;
  }

  /** The cached answer for `key`, or `compute()` cached. */
  answer(key: string, compute: () => CanResult, structural = false): CanResult {
    this.sync();
    const answers = structural ? this.#structuralAnswers() : this.#answers;
    const hit = answers.get(key);
    if (hit) return hit;
    const result = compute();
    answers.set(key, result);
    return result;
  }

  #structuralAnswers(): Map<string, CanResult> {
    const now = this.structure();
    if (!same(now, this.#structure)) {
      this.#structure = now;
      this.#structural.clear();
    }
    return this.#structural;
  }

  /** The derived value for the current state, or `compute()` cached. */
  active(compute: () => Active): Active {
    this.sync();
    return (this.#active ??= compute());
  }
}

/**
 * What the answer of a {@link STRUCTURE_COMMANDS} command reads of the current slide. New
 * Slide reads the layout of the slide (the one it copies, and the one `newSlideLayout` picks
 * from it); its index only places the new slide, which no refusal reads. Add Section reads the
 * index. @internal
 */
export function structureSlideKey(
  command: EditorCommand,
  pkg: PptxPackage | undefined,
  slide: number
): string {
  if (command.type !== 'addSlide' || !pkg) return `slide:${slide}`;
  const slides = presentationInfo(pkg).slides;
  const source = slides[Math.max(0, Math.min(slide, slides.length - 1))]?.layoutPart;
  return `layout:${source ?? ''}\u0000${newSlideLayout(pkg, slide) ?? ''}`;
}

const same = (a: readonly unknown[], b: readonly unknown[]): boolean =>
  a.length === b.length && a.every((v, i) => v === b[i]);

/**
 * The state the answers of {@link STRUCTURE_COMMANDS} read, besides the slide index and the
 * mode: the session, a token for its package, the part names the session has held, and the
 * collaboration binding with how many numbers it minted (new part names and ids come from
 * them). The package token changes unless the package differs from the last one only in the
 * XML of slides in the slide list, not in their relationships. @internal
 */
export function createStructureKey(): (session: DocumentSession | null) => readonly unknown[] {
  const tokens = new WeakMap<PptxPackage, object>();
  let last: WeakRef<PptxPackage> | undefined;
  const tokenOf = (pkg: PptxPackage): object => {
    let token = tokens.get(pkg);
    if (token) return token;
    const prev = last?.deref();
    const changed = prev ? changedPartNames(prev, pkg, 32) : null;
    if (prev && changed?.every((name) => slideContentOnly(prev, pkg, name)))
      token = tokens.get(prev);
    token ??= {};
    tokens.set(pkg, token);
    last = new WeakRef(pkg);
    return token;
  };
  return (session) => {
    if (!session) return [];
    const binding = replicaBindingOf(session.store);
    let minted = 0;
    for (const set of binding?.minted.values() ?? []) minted += set.size;
    const pkg = session.store.package;
    return [session, tokenOf(pkg), session.usedPartNames.size, binding, minted];
  };
}

/** Whether part `name` is the XML part of a slide of the list in both packages. */
function slideContentOnly(prev: PptxPackage, pkg: PptxPackage, name: string): boolean {
  const a = prev.parts.get(name);
  const b = pkg.parts.get(name);
  if (a?.kind !== 'xml' || b?.kind !== 'xml') return false;
  const isSlide = (p: PptxPackage) => presentationInfo(p).slides.some((s) => s.partName === name);
  return isSlide(pkg) && isSlide(prev);
}
