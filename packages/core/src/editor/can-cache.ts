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
import {
  changedPartNames,
  nodeIndex,
  presentationInfo,
  xmlPartOf,
  type Node,
  type PptxPackage,
} from '../store/index.ts';
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

/**
 * The size from which a shape selection is large: `can` then checks only the mode, not each
 * command, because a check builds the whole edit (a Duplicate of every selected shape). The
 * toolbar asks about every command on every change, and must answer within a frame. A
 * command chosen for a large selection makes the full check when it runs (`exec`) and shows
 * its reason when it cannot apply.
 */
export const LARGE_SELECTION = 250;

/**
 * The number of XML nodes from which the selected shapes are large for the commands that copy
 * them ({@link COPYING_COMMANDS}): one text box of 50,000 runs holds 150,000 nodes, and a check
 * of Duplicate copies them all. Other commands check such a selection in full: a check of
 * Group on one shape still says that Group needs two.
 */
export const LARGE_SELECTION_NODES = 20_000;

/** Commands whose check copies every node of the selected shapes. */
const COPYING_COMMANDS: ReadonlySet<EditorCommand['type']> = new Set(['duplicate']);

/**
 * Commands that edit every placeholder of a slide, whatever the selection: a check of one that
 * names a slide of more than {@link LARGE_SELECTION_NODES} nodes is large too.
 */
const SLIDE_WIDE_COMMANDS: ReadonlySet<EditorCommand['type']> = new Set([
  'applyLayout',
  'resetSlides',
]);

/**
 * The quick `can` answer for a large shape selection (`LARGE_SELECTION`,
 * `LARGE_SELECTION_NODES`) on slide `slide` of `session`, or for a slide-wide command on a large
 * slide; null to check the command in full. A {@link STRUCTURE_COMMANDS} answer is cached
 * without the selection, so it never takes the quick answer. @internal
 */
export function largeSelectionAnswer(
  selection: { readonly kind: string; readonly nodeIds?: readonly string[] },
  readOnly: boolean,
  session?: DocumentSession,
  slide = 0,
  command?: EditorCommand
): CanResult | null {
  if (command && STRUCTURE_COMMANDS.has(command.type)) return null;
  const pkg = session?.store.package;
  const part = session?.info.slides[slide]?.partName;
  const quick: CanResult = readOnly ? { ok: false, reason: 'disabled.viewing' } : { ok: true };
  if (command && SLIDE_WIDE_COMMANDS.has(command.type)) {
    // The slides the command names, else the current one.
    const named = 'indices' in command ? command.indices : undefined;
    const large = (named ?? [slide]).some((i) => {
      const at = session?.info.slides[i]?.partName;
      const root = pkg && at ? xmlPartOf(pkg, at)?.root : undefined;
      return !!root && holdsManyNodes(pkg, at, [root.id]);
    });
    return large ? quick : null;
  }
  if (selection.kind !== 'shapes') return null;
  const ids = selection.nodeIds ?? [];
  if (ids.length > LARGE_SELECTION) return quick;
  const copies = !command || COPYING_COMMANDS.has(command.type);
  return copies && holdsManyNodes(pkg, part, ids) ? quick : null;
}

/** The answers of `holdsManyNodes` per package, by part and ids: asked once per command. */
const counts = new WeakMap<PptxPackage, Map<string, boolean>>();

/** Whether the shapes `ids` of `part` hold more than {@link LARGE_SELECTION_NODES} nodes. */
function holdsManyNodes(
  pkg: PptxPackage | undefined,
  part: string | undefined,
  ids: readonly string[]
): boolean {
  if (!pkg || !part) return false;
  // At most `LARGE_SELECTION` ids: the key costs little next to the walk it saves.
  const key = `${part}\u0000${ids.join('\u0000')}`;
  let known = counts.get(pkg);
  if (!known) counts.set(pkg, (known = new Map()));
  const was = known.get(key);
  if (was !== undefined) return was;
  const xml = xmlPartOf(pkg, part);
  const index = xml ? nodeIndex(xml) : undefined;
  let left = LARGE_SELECTION_NODES;
  const stack: Node[] = [];
  for (const id of ids) {
    const node = index?.get(id)?.node;
    if (node) stack.push(node);
  }
  // The walk stops at the limit, so it costs at most that many visits.
  while (stack.length > 0 && left > 0) {
    const node = stack.pop()!;
    left -= 1;
    if (node.type === 'element') for (const kid of node.children) stack.push(kid);
  }
  const many = left <= 0;
  known.set(key, many);
  return many;
}

/** How long after the last keystroke the toolbar's formatting answers are exact again. */
export const TYPING_SETTLE_MS = 300;

/**
 * Formatting commands whose `can` answer depends on the text box and the mode, not on the text
 * typed into it: a typing burst keeps their answers (`CanCache.typed`). A link reads the caret
 * position, so Insert Link is not one of them. Typing can start a list ("1. "), so the list
 * commands are not either. @internal
 */
export const TYPING_STABLE_COMMANDS: ReadonlySet<EditorCommand['type']> = new Set([
  'toggleBold',
  'toggleItalic',
  'toggleUnderline',
  'toggleStrike',
  'setFontFamily',
  'setFontSize',
  'setTextColor',
  'setHighlight',
  'clearFormatting',
  'toggleFormatPainter',
  'setAlign',
  'setLineSpacing',
]);

/**
 * The text box a text selection is in, as a key for `CanCache.typed`: the shape, the table cell,
 * and whether the selection is a caret. Null for any other selection. @internal
 */
export function typingBox(selection: {
  kind: string;
  nodeId?: string;
  cell?: unknown;
  anchor?: unknown;
  focus?: unknown;
}): string | undefined {
  if (selection.kind !== 'text') return undefined;
  const caret = JSON.stringify(selection.anchor) === JSON.stringify(selection.focus);
  return JSON.stringify([selection.nodeId, selection.cell ?? null, caret]);
}

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

  /**
   * Note a keystroke in the text of `box` (`typingBox`). Until typing stops for
   * {@link TYPING_SETTLE_MS}, the answers of {@link TYPING_STABLE_COMMANDS} in that text box
   * stay as they were: they do not change while the text grows, and a probe of each one costs
   * as much as the text box is long. When typing stops, `settled` runs, so the toolbar asks
   * again and gets exact answers.
   */
  typed(box: string, settled: () => void): void {
    if (this.#typingBox !== box) this.#typing.clear();
    this.#typingBox = box;
    if (this.#settle) clearTimeout(this.#settle);
    this.#settle = setTimeout(() => {
      this.#settle = null;
      this.#typingBox = null;
      this.#typing.clear();
      settled();
    }, TYPING_SETTLE_MS);
  }

  readonly #typing = new Map<string, CanResult>();
  #typingBox: string | null = null;
  #settle: ReturnType<typeof setTimeout> | null = null;

  /** Stop the typing burst without a `settled` call (the editor is destroyed). */
  dispose(): void {
    if (this.#settle) clearTimeout(this.#settle);
    this.#settle = null;
    this.#typingBox = null;
    this.#typing.clear();
  }

  /**
   * The answer for `command`, cached under `key`: a {@link STRUCTURE_COMMANDS} answer under the
   * structure key, a {@link TYPING_STABLE_COMMANDS} answer through a typing burst in its text box.
   */
  ask(
    command: EditorCommand,
    key: string,
    at: {
      readonly pkg: PptxPackage | undefined;
      readonly slide: number;
      readonly mode: string;
      readonly selection: Parameters<typeof typingBox>[0];
    },
    compute: () => CanResult
  ): CanResult {
    if (STRUCTURE_COMMANDS.has(command.type))
      return this.answer(`${key}@${structureSlideKey(command, at.pkg, at.slide)}`, compute, true);
    const box = TYPING_STABLE_COMMANDS.has(command.type) ? typingBox(at.selection) : undefined;
    return this.answer(key, compute, false, box === undefined ? undefined : `${at.mode}${box}`);
  }

  /**
   * The cached answer for `key`, or `compute()` cached. With `box`, the text box that the
   * selection is in, a {@link TYPING_STABLE_COMMANDS} answer holds through a typing burst.
   */
  answer(key: string, compute: () => CanResult, structural = false, box?: string): CanResult {
    if (box !== undefined && box === this.#typingBox) {
      const hit = this.#typing.get(key);
      if (hit) return hit;
      const result = this.answer(key, compute, structural);
      this.#typing.set(key, result);
      return result;
    }
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
