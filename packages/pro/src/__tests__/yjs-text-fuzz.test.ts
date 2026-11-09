/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Seeded fuzz of concurrent text edits (openspec change `collaboration-text-lineage`,
// "Character identity"). Three editors on three Y.Docs type, type over a selection, delete
// ranges in one paragraph and across two, press Enter, join paragraphs, undo, redo, and format
// (bold, italic, color) in one text box. The relay holds every update back and delivers parts of
// it at random; a round ends with a flush, or keeps the held updates for the next round. The
// harness is `fuzz-support.ts`; its header lists the environment knobs. Two modes, each with its
// prefix: `TEXT` (no formats) and `FORMAT` (with formats). `yjs-text-fuzz-regressions.ts` holds
// fixed histories.
//
// Invariants:
// - Each edit changes only what it may on its own replica: a delete removes exactly its
//   selection, typing adds its letter and removes only the selection it types over, Enter,
//   a join, and a format remove nothing, and an undo or redo removes only characters its peer
//   placed and brings back only characters its peer removed or that a concurrent edit took.
// - After each flush, every session is `ready`, and the replicas hold the same package.
// - Text conservation (`fuzz-ledger.ts`): each character shows at most once, shows unless a
//   peer removed it, and stays removed unless a peer brought it back. This holds also for a
//   character placed at the same time as a peer's removal of its paragraph (a join, a delete
//   across paragraphs, an undo of Enter): no edit is exempt.
// - At the end, each replica's save reopens to the same deck, a second save keeps it, and the
//   saved deck passes the part checks.

import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import {
  CharacterLedger,
  NONE,
  editClock,
  localProblems,
  stateOf,
  withoutRepeated,
  type LocalRule,
} from './fuzz-ledger.ts';
import {
  FuzzLog,
  clientIdsOf,
  convergenceProblems,
  fuzzKnobs,
  fuzzSuite,
  openFuzzPeers,
  roundProblems,
  seededRandom,
  type Delivery,
  type DeliveryTo,
  type FuzzKnownFailure,
  type FuzzRegression,
  type FuzzKnownSeed,
  type Relay,
} from './fuzz-support.ts';
import { IN, paragraphs, selectIn } from './lineage-support.ts';
import {
  FORMAT_HISTORIES,
  LONG_FORMAT_HISTORIES,
  TEXT_HISTORIES,
  UNDO_HISTORIES,
} from './yjs-text-fuzz-regressions.ts';

const FILE = 'packages/pro/src/__tests__/yjs-text-fuzz.test.ts';
const KNOBS = { text: fuzzKnobs('TEXT', 24), format: fuzzKnobs('FORMAT', 24) } as const;
const PEERS = 3;
const ROUNDS = 5;

/** Three runs and a paragraph end, then a second paragraph. */
const deck = () =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: 2 * IN,
          text:
            para(run('ABCD') + run('EF', 'b="1"') + run('GH') + '<a:endParaRPr lang="en-US"/>') +
            para(run('IJKL')),
        }),
      },
    ],
  });
const DECK_TEXT = 'ABCDEFGHIJKL';

export type EditKind =
  | 'type'
  | 'enter'
  | 'delete'
  | 'deleteAcross'
  | 'join'
  | 'bold'
  | 'italic'
  | 'color'
  | 'undo'
  | 'redo';
export type Mode = 'text' | 'format';
const EDITS: Readonly<Record<Mode, readonly EditKind[]>> = {
  text: ['type', 'type', 'type', 'enter', 'delete', 'deleteAcross', 'join', 'undo', 'redo'],
  format: [
    'type',
    'type',
    'enter',
    'delete',
    'deleteAcross',
    'join',
    'bold',
    'italic',
    'color',
    'undo',
    'redo',
  ],
};
const COLORS = ['accent1', 'accent2', 'accent6'] as const;

/** One step, in paragraph coordinates, so a shrunk list stays meaningful. */
export type TextAction =
  | {
      readonly kind: 'edit';
      readonly peer: number;
      readonly edit: EditKind;
      /**
       * The selection: `length` characters from `offset` in `paragraph`. For `deleteAcross`, from
       * `offset` in `paragraph` to offset `length` in the paragraph after it.
       */
      readonly paragraph: number;
      readonly offset: number;
      readonly length: number;
      /** `type` and `enter`: replace the selection, not only insert at its start. */
      readonly over?: true;
      /** The character `type` types: each one is new. */
      readonly letter?: string;
      readonly color?: (typeof COLORS)[number];
    }
  | Delivery
  | DeliveryTo
  /** Deliver every waiting update; the round's checks run. */
  | { readonly kind: 'flush' };

type Edit = Extract<TextAction, { kind: 'edit' }>;

/** Selects from `from` in `paragraph` to `to` in the paragraph after it. */
function selectAcross(e: PptxEditorInstance, paragraph: number, from: number, to: number): void {
  e.setSelection({
    kind: 'text',
    nodeId: e.session!.layout(0)!.items[0]!.element.nodeId,
    anchor: { paragraph, offset: from },
    focus: { paragraph: paragraph + 1, offset: to },
  });
}

/** Each replica's paragraphs: a `p0:` line for each peer. */
const showAll = (editors: readonly PptxEditorInstance[]) =>
  editors.map((e, i) => `    p${i}: ${JSON.stringify(paragraphs(e))}`).join('\n');

/** How often each check ran, per mode. */
const STATS: Record<
  Mode,
  { rounds: number; edits: number; characterChecks: number; exempt: number; blind: number }
> = {
  text: { rounds: 0, edits: 0, characterChecks: 0, exempt: 0, blind: 0 },
  format: { rounds: 0, edits: 0, characterChecks: 0, exempt: 0, blind: 0 },
};

/** Three replicas, and what the conservation check knows. */
class TextRun {
  readonly log: FuzzLog;
  private readonly ledger = new CharacterLedger(DECK_TEXT);

  private constructor(
    private readonly mode: Mode,
    private readonly editors: readonly PptxEditorInstance[],
    readonly relay: Relay,
    private readonly docs: Awaited<ReturnType<typeof openFuzzPeers>>['docs']
  ) {
    this.log = new FuzzLog(KNOBS[mode].trace, () => showAll(editors));
  }

  static async open(clients: readonly number[], mode: Mode): Promise<TextRun> {
    const { editors, relay, docs } = await openFuzzPeers(deck(), clients);
    return new TextRun(mode, editors, relay, docs);
  }

  destroy(): void {
    const { stats } = this.ledger;
    const totals = STATS[this.mode];
    totals.characterChecks += stats.characterChecks;
    totals.exempt += stats.exempt;
    totals.blind += stats.blind;
    for (const e of this.editors) e.destroy();
  }

  perform(action: TextAction): void {
    if (action.kind === 'deliver') {
      const n = this.relay.deliver(action.from, action.to, action.count);
      if (n > 0) this.log.note(`deliver ${n} p${action.from} -> p${action.to}`);
      return;
    }
    if (action.kind === 'deliverTo') {
      const n = this.relay.deliverTo(action.to);
      if (n > 0) this.log.note(`deliver ${n} -> p${action.to}`);
      return;
    }
    if (action.kind === 'flush') {
      this.relay.flush();
      this.log.note('flush');
      this.checkRound();
      return;
    }
    this.edit(action);
  }

  private edit(action: Edit): void {
    const e = this.editors[action.peer]!;
    const ps = paragraphs(e);
    if (ps.length === 0) return;
    let p = Math.min(action.paragraph, ps.length - 1);
    const length = ps[p]!.length;
    let from = Math.min(action.offset, length);
    let to = Math.min(length, from + action.length);
    switch (action.edit) {
      case 'type':
      case 'enter':
        if (!action.over) to = from;
        break;
      case 'delete':
        // A range, so a delete never joins: `join` does that.
        if (length === 0) return;
        if (from === to) [from, to] = from < length ? [from, from + 1] : [from - 1, from];
        break;
      case 'join':
        // Backspace at the start of a paragraph after the first.
        if (ps.length < 2) return;
        p = Math.max(1, p);
        from = to = 0;
        break;
      case 'deleteAcross':
        // A selection that ends in the next paragraph: the delete removes a paragraph.
        if (p + 1 >= ps.length) return;
        to = Math.min(action.length, ps[p + 1]!.length);
        break;
      default:
    }
    const across = action.edit === 'deleteAcross';
    const history = action.edit === 'undo' || action.edit === 'redo';
    if (across) selectAcross(e, p, from, to);
    else if (!history) selectIn(e, p, from, to);
    const before = new Set(ps.join(''));
    const doc = this.docs[action.peer]!;
    const state = stateOf(doc);
    let result: { ok: boolean };
    if (action.edit === 'type') result = e.input('text', action.letter);
    else if (action.edit === 'enter') result = e.input('paragraph');
    else if (action.edit === 'delete' || action.edit === 'join' || across)
      result = e.input('deleteBackward');
    else if (action.edit === 'bold') result = e.exec({ type: 'toggleBold' });
    else if (action.edit === 'italic') result = e.exec({ type: 'toggleItalic' });
    else if (history) result = e.exec({ type: action.edit });
    else result = e.exec({ type: 'setTextColor', color: { theme: action.color ?? 'accent2' } });
    const shown = paragraphs(e);
    const after = new Set(shown.join(''));
    const clock = editClock(doc, state);
    const gone = new Set([...before].filter((ch) => !after.has(ch)));
    const came = new Set([...after].filter((ch) => !before.has(ch)));
    const what = history
      ? `p${action.peer} ${action.edit}`
      : `p${action.peer} ${action.edit}${action.letter ? ` ${JSON.stringify(action.letter)}` : ''}` +
        (across
          ? ` from paragraph ${p} offset ${from} to paragraph ${p + 1} offset ${to}`
          : ` paragraph ${p} [${from}, ${to})`);
    this.log.note(`${what} ${result.ok ? 'ok' : 'refused'}`);
    STATS[this.mode].edits += 1;
    const selected = across ? ps[p]!.slice(from) + ps[p + 1]!.slice(0, to) : ps[p]!.slice(from, to);
    const rule = withoutRepeated(
      this.rule(action, doc.clientID, result.ok, new Set(selected)),
      ps.join('')
    );
    for (const problem of localProblems(what, gone, came, rule)) this.log.fail(problem);
    // The place of a character that showed: the rest of its paragraph.
    const placeOf = (ch: string) =>
      new Set([...(shown.find((x) => x.includes(ch)) ?? '')].filter((c) => !came.has(c)));
    // No edit is risky: a join, a delete across paragraphs, an undo of Enter, and a redo of a
    // join keep what a peer types into the paragraph they remove at the same time.
    this.ledger.record(clock, gone, came, placeOf);
  }

  /** What one edit may change on its own replica. */
  private rule(action: Edit, client: number, ok: boolean, selected: Set<string>): LocalRule {
    switch (action.edit) {
      case 'type': {
        const letter = new Set(action.letter ? [action.letter] : []);
        const over = action.over ? selected : NONE;
        return {
          mayRemove: over,
          mayAdd: letter,
          ...(ok ? { mustAdd: letter, mustRemove: over } : {}),
        };
      }
      case 'enter': {
        const over = action.over ? selected : NONE;
        return { mayRemove: over, mayAdd: NONE, ...(ok ? { mustRemove: over } : {}) };
      }
      case 'delete':
      case 'deleteAcross':
        return { mayRemove: selected, mayAdd: NONE, ...(ok ? { mustRemove: selected } : {}) };
      case 'undo':
      case 'redo':
        return {
          mayRemove: this.ledger.placedBy(client),
          mayAdd: this.ledger.mayBringBack(client),
        };
      default:
        return { mayRemove: NONE, mayAdd: NONE };
    }
  }

  /** After a flush: sessions ready, replicas equal, and text conservation. */
  private checkRound(): void {
    for (const problem of roundProblems(this.editors)) this.log.fail(problem);
    if (this.log.failed) return;
    STATS[this.mode].rounds += 1;
    const shown = paragraphs(this.editors[0]!);
    for (const problem of this.ledger.problems(shown.join('')))
      this.log.fail(`${problem} in ${JSON.stringify(shown)}`);
  }

  /** The end-of-run invariants of `fuzz-support.ts`, after the round checks. */
  finish(): void {
    this.relay.flush();
    this.checkRound();
    if (this.log.failed) return;
    for (const problem of convergenceProblems(this.editors).problems) this.log.fail(problem);
  }

  /** A random edit on `peer`'s paragraphs. */
  randomEdit(peer: number, random: () => number, letter: () => string): TextAction {
    const ps = paragraphs(this.editors[peer]!);
    const paragraph = Math.floor(random() * ps.length);
    const offset = Math.floor(random() * ((ps[paragraph]?.length ?? 0) + 1));
    const length = Math.floor(random() * 4);
    const edits = EDITS[this.mode];
    const edit = edits[Math.floor(random() * edits.length)]!;
    const over = (edit === 'type' || edit === 'enter') && length > 0 && random() < 0.3;
    return {
      kind: 'edit',
      peer,
      edit,
      paragraph,
      offset,
      length,
      ...(over ? { over: true as const } : {}),
      ...(edit === 'type' ? { letter: letter() } : {}),
      ...(edit === 'color' ? { color: COLORS[Math.floor(random() * COLORS.length)]! } : {}),
    };
  }
}

/** Run `actions` on fresh replicas, then the end checks. */
async function perform(run: TextRun, actions: Iterable<TextAction>): Promise<string[]> {
  try {
    for (const action of actions) {
      run.perform(action);
      if (run.log.failed) break;
    }
    if (!run.log.failed) run.finish();
  } catch (error) {
    run.log.fail(`threw ${String(error)}`);
  } finally {
    run.destroy();
  }
  return run.log.problems;
}

/** Generate and perform one seed's actions. */
async function runSeed(seed: number, mode: Mode) {
  const random = seededRandom(seed);
  const run = await TextRun.open(clientIdsOf(PEERS, seed), mode);
  const actions: TextAction[] = [];
  let n = 0;
  const letter = () => String.fromCharCode(0x100 + n++);
  // The generator reads each replica after the actions before it, so it yields as they run.
  function* generate(): Generator<TextAction> {
    for (let round = 0; round < ROUNDS; round += 1) {
      const steps = 3 + Math.floor(random() * 5);
      for (let step = 0; step < steps; step += 1) {
        yield run.randomEdit(Math.floor(random() * PEERS), random, letter);
        while (random() < 0.35) {
          const delivery = run.relay.randomDelivery(random);
          if (delivery) yield delivery;
        }
      }
      // Most rounds end with a flush and its checks; the others keep held updates for later.
      if (round === ROUNDS - 1 || random() < 0.6) yield { kind: 'flush' };
    }
  }
  const recorded = (function* () {
    for (const action of generate()) {
      actions.push(action);
      yield action;
    }
  })();
  const problems = await perform(run, recorded);
  return { problems, actions };
}

/** Histories that once failed, kept as regressions after their issue was fixed. */
const FIXED: Readonly<Record<Mode, readonly FuzzRegression<TextAction>[]>> = {
  text: [
    {
      name: 'text typed into a paragraph that a peer empties and joins stays',
      seed: 163,
      actions: [
        { kind: 'edit', peer: 2, edit: 'enter', paragraph: 0, offset: 5, length: 0 },
        { kind: 'deliver', from: 2, to: 0, count: 1 },
        { kind: 'edit', peer: 2, edit: 'join', paragraph: 1, offset: 3, length: 3 },
        { kind: 'deliver', from: 2, to: 0, count: 1 },
        { kind: 'edit', peer: 0, edit: 'deleteAcross', paragraph: 0, offset: 4, length: 1 },
        { kind: 'edit', peer: 2, edit: 'undo', paragraph: 0, offset: 5, length: 1 },
        { kind: 'flush' },
        { kind: 'edit', peer: 0, edit: 'undo', paragraph: 1, offset: 0, length: 3 },
        { kind: 'edit', peer: 2, edit: 'type', paragraph: 1, offset: 0, length: 3, letter: 'Ă' },
        { kind: 'edit', peer: 0, edit: 'deleteAcross', paragraph: 0, offset: 0, length: 2 },
      ],
    },
    {
      name: 'a run that a join moved into a paragraph that a peer deletes at once shows there',
      seed: 184,
      actions: [
        { kind: 'edit', peer: 2, edit: 'enter', paragraph: 1, offset: 0, length: 2, over: true },
        { kind: 'deliver', from: 2, to: 0, count: 1 },
        { kind: 'deliver', from: 2, to: 1, count: 1 },
        { kind: 'edit', peer: 0, edit: 'join', paragraph: 2, offset: 0, length: 3 },
        { kind: 'edit', peer: 1, edit: 'deleteAcross', paragraph: 0, offset: 0, length: 2 },
      ],
    },
    {
      name: 'text typed into a run that a peer deleted before it removes the paragraph stays',
      seed: 121,
      actions: [
        { kind: 'edit', peer: 1, edit: 'enter', paragraph: 0, offset: 8, length: 0 },
        { kind: 'edit', peer: 2, edit: 'enter', paragraph: 1, offset: 4, length: 0 },
        { kind: 'edit', peer: 1, edit: 'join', paragraph: 2, offset: 1, length: 2 },
        { kind: 'flush' },
        { kind: 'edit', peer: 2, edit: 'deleteAcross', paragraph: 1, offset: 0, length: 3 },
        { kind: 'edit', peer: 2, edit: 'join', paragraph: 0, offset: 0, length: 2 },
        { kind: 'edit', peer: 1, edit: 'type', paragraph: 1, offset: 4, length: 2, letter: 'Ć' },
      ],
    },
    {
      name: 'a deleted character stays deleted after a peer that saw the delete presses Enter',
      seed: 8,
      actions: [
        { kind: 'edit', peer: 0, edit: 'type', paragraph: 0, offset: 8, length: 2, letter: 'ā' },
        { kind: 'flush' },
        { kind: 'edit', peer: 2, edit: 'delete', paragraph: 0, offset: 3, length: 1 },
        { kind: 'deliver', from: 2, to: 1, count: 1 },
        { kind: 'edit', peer: 1, edit: 'enter', paragraph: 0, offset: 1, length: 0 },
        { kind: 'flush' },
      ],
    },
    {
      name: 'a character that a peer types and deletes while another peer splits its run stays deleted',
      seed: 41,
      actions: [
        { kind: 'edit', peer: 1, edit: 'enter', paragraph: 1, offset: 2, length: 3 },
        { kind: 'edit', peer: 0, edit: 'type', paragraph: 1, offset: 4, length: 1, letter: 'Ā' },
        { kind: 'deliver', from: 0, to: 2, count: 1 },
        { kind: 'edit', peer: 2, edit: 'delete', paragraph: 1, offset: 5, length: 1 },
      ],
    },
    {
      name: 'an undo of Enter keeps the text that a peer typed into the new paragraph',
      seed: 29,
      actions: [
        { kind: 'edit', peer: 2, edit: 'enter', paragraph: 0, offset: 5, length: 0 },
        { kind: 'deliver', from: 2, to: 1, count: 1 },
        { kind: 'edit', peer: 1, edit: 'type', paragraph: 1, offset: 0, length: 3, letter: 'ă' },
        { kind: 'deliver', from: 1, to: 2, count: 1 },
        { kind: 'edit', peer: 2, edit: 'undo', paragraph: 1, offset: 2, length: 1 },
      ],
    },
    {
      name: 'an undo of a join after a concurrent Enter keeps the joined text',
      seed: 88,
      actions: [
        { kind: 'edit', peer: 0, edit: 'enter', paragraph: 2, offset: 0, length: 0 },
        { kind: 'edit', peer: 2, edit: 'join', paragraph: 1, offset: 9, length: 0 },
        { kind: 'flush' },
        {
          kind: 'edit',
          peer: 0,
          edit: 'type',
          paragraph: 1,
          offset: 0,
          length: 1,
          over: true,
          letter: 'ċ',
        },
        { kind: 'flush' },
        { kind: 'edit', peer: 2, edit: 'undo', paragraph: 5, offset: 1, length: 3 },
      ],
    },
    {
      name: 'two concurrent joins of neighboring paragraphs keep both paragraphs',
      seed: 253,
      actions: [
        { kind: 'edit', peer: 2, edit: 'enter', paragraph: 0, offset: 2, length: 1 },
        { kind: 'flush' },
        { kind: 'edit', peer: 0, edit: 'join', paragraph: 2, offset: 4, length: 0 },
        { kind: 'edit', peer: 1, edit: 'join', paragraph: 1, offset: 5, length: 0 },
        { kind: 'flush' },
      ],
    },
    {
      name: 'an undo of Enter keeps the text that a peer typed and joined back',
      seed: 99,
      actions: [
        { kind: 'edit', peer: 1, edit: 'enter', paragraph: 0, offset: 5, length: 3 },
        { kind: 'flush' },
        { kind: 'edit', peer: 0, edit: 'type', paragraph: 1, offset: 1, length: 3, letter: 'Ā' },
        { kind: 'edit', peer: 0, edit: 'join', paragraph: 1, offset: 4, length: 1 },
        { kind: 'flush' },
        { kind: 'edit', peer: 1, edit: 'undo', paragraph: 0, offset: 1, length: 0 },
      ],
    },
    {
      name: 'an undo of Enter after an undo of a join keeps the text a peer typed over',
      seed: 1263,
      actions: [
        { kind: 'edit', peer: 0, edit: 'enter', paragraph: 0, offset: 2, length: 0 },
        { kind: 'edit', peer: 0, edit: 'join', paragraph: 1, offset: 6, length: 0 },
        { kind: 'edit', peer: 0, edit: 'undo', paragraph: 0, offset: 4, length: 1 },
        { kind: 'flush' },
        {
          kind: 'edit',
          peer: 2,
          edit: 'type',
          paragraph: 1,
          offset: 1,
          length: 3,
          over: true,
          letter: 'ć',
        },
        { kind: 'flush' },
        { kind: 'edit', peer: 0, edit: 'undo', paragraph: 0, offset: 5, length: 3 },
      ],
    },
    {
      name: "an undo of typing reaches a peer's Enter that copies the text at the same time",
      seed: 61,
      actions: [
        { kind: 'edit', peer: 1, edit: 'deleteAcross', paragraph: 0, offset: 4, length: 3 },
        { kind: 'edit', peer: 1, edit: 'delete', paragraph: 0, offset: 2, length: 3 },
        {
          kind: 'edit',
          peer: 2,
          edit: 'type',
          paragraph: 1,
          offset: 3,
          length: 3,
          over: true,
          letter: 'ā',
        },
        { kind: 'edit', peer: 0, edit: 'type', paragraph: 1, offset: 4, length: 0, letter: 'Ă' },
        { kind: 'flush' },
        { kind: 'edit', peer: 0, edit: 'undo', paragraph: 0, offset: 1, length: 2 },
        { kind: 'edit', peer: 1, edit: 'undo', paragraph: 0, offset: 2, length: 3 },
        { kind: 'edit', peer: 1, edit: 'enter', paragraph: 0, offset: 6, length: 0 },
      ],
    },
    {
      name: "an undo of a peer's Enter and a delete across it keep text typed beside them",
      seed: 62,
      actions: [
        { kind: 'edit', peer: 2, edit: 'enter', paragraph: 1, offset: 1, length: 2, over: true },
        { kind: 'edit', peer: 1, edit: 'type', paragraph: 0, offset: 7, length: 0, letter: 'Ă' },
        { kind: 'edit', peer: 0, edit: 'deleteAcross', paragraph: 0, offset: 0, length: 3 },
        { kind: 'flush' },
        { kind: 'edit', peer: 1, edit: 'type', paragraph: 0, offset: 1, length: 2, letter: 'Ć' },
        { kind: 'flush' },
        { kind: 'edit', peer: 2, edit: 'enter', paragraph: 0, offset: 1, length: 3 },
        { kind: 'edit', peer: 1, edit: 'enter', paragraph: 0, offset: 2, length: 0 },
        { kind: 'deliver', from: 1, to: 2, count: 1 },
        { kind: 'edit', peer: 2, edit: 'enter', paragraph: 0, offset: 0, length: 3 },
        { kind: 'edit', peer: 0, edit: 'enter', paragraph: 0, offset: 1, length: 0 },
        { kind: 'edit', peer: 0, edit: 'undo', paragraph: 0, offset: 0, length: 0 },
        { kind: 'edit', peer: 2, edit: 'type', paragraph: 3, offset: 0, length: 3, letter: 'ĉ' },
        { kind: 'deliver', from: 2, to: 0, count: 5 },
        { kind: 'edit', peer: 0, edit: 'deleteAcross', paragraph: 0, offset: 1, length: 2 },
      ],
    },
    {
      name: 'an undo of Enter after a redo of a join keeps text typed into the joined run',
      seed: 73,
      actions: [
        { kind: 'edit', peer: 2, edit: 'join', paragraph: 0, offset: 5, length: 2 },
        { kind: 'flush' },
        { kind: 'edit', peer: 0, edit: 'enter', paragraph: 0, offset: 1, length: 3, over: true },
        { kind: 'edit', peer: 2, edit: 'undo', paragraph: 0, offset: 9, length: 3 },
        { kind: 'edit', peer: 2, edit: 'redo', paragraph: 1, offset: 2, length: 0 },
        { kind: 'edit', peer: 0, edit: 'enter', paragraph: 1, offset: 0, length: 2 },
        { kind: 'flush' },
        { kind: 'edit', peer: 2, edit: 'type', paragraph: 0, offset: 0, length: 0, letter: 'ĉ' },
        { kind: 'edit', peer: 2, edit: 'deleteAcross', paragraph: 0, offset: 2, length: 0 },
        { kind: 'edit', peer: 1, edit: 'type', paragraph: 0, offset: 2, length: 1, letter: 'Ċ' },
        { kind: 'edit', peer: 0, edit: 'undo', paragraph: 3, offset: 5, length: 1 },
      ],
    },
    {
      name: "a redo of a delete across paragraphs reaches a peer's Enter copy of the text",
      seed: 637,
      actions: [
        { kind: 'edit', peer: 0, edit: 'deleteAcross', paragraph: 0, offset: 1, length: 2 },
        { kind: 'edit', peer: 0, edit: 'undo', paragraph: 0, offset: 0, length: 2 },
        { kind: 'edit', peer: 0, edit: 'redo', paragraph: 0, offset: 1, length: 2 },
        { kind: 'edit', peer: 2, edit: 'enter', paragraph: 3, offset: 1, length: 2 },
        { kind: 'flush' },
      ],
    },
    {
      name: "a delete of text an undo put back reaches a peer's Enter copy of the text",
      seed: 1767,
      actions: [
        { kind: 'edit', peer: 0, edit: 'enter', paragraph: 0, offset: 1, length: 0 },
        { kind: 'edit', peer: 1, edit: 'enter', paragraph: 0, offset: 3, length: 3 },
        { kind: 'edit', peer: 0, edit: 'undo', paragraph: 0, offset: 1, length: 2 },
        {
          kind: 'edit',
          peer: 0,
          edit: 'type',
          paragraph: 0,
          offset: 3,
          length: 1,
          over: true,
          letter: 'ā',
        },
        { kind: 'flush' },
      ],
    },
  ],
  format: [
    {
      name: 'an undo of a format keeps the text that a peer typed over the formatted text deleted',
      seed: 20,
      actions: [
        { kind: 'edit', peer: 0, edit: 'bold', paragraph: 1, offset: 2, length: 3 },
        {
          kind: 'edit',
          peer: 1,
          edit: 'type',
          paragraph: 1,
          offset: 1,
          length: 3,
          over: true,
          letter: 'Ĉ',
        },
        { kind: 'flush' },
        { kind: 'edit', peer: 0, edit: 'undo', paragraph: 2, offset: 6, length: 2 },
      ],
    },
    {
      name: 'an undo of a format keeps the text that a peer deletes at the same time deleted',
      seed: 9,
      actions: [
        { kind: 'edit', peer: 2, edit: 'italic', paragraph: 0, offset: 2, length: 3 },
        { kind: 'edit', peer: 2, edit: 'undo', paragraph: 1, offset: 4, length: 3 },
        { kind: 'edit', peer: 1, edit: 'delete', paragraph: 0, offset: 2, length: 2 },
      ],
    },
    {
      name: 'a redo of a format keeps the text that a peer deleted before the undo deleted',
      seed: 425,
      actions: [
        { kind: 'edit', peer: 0, edit: 'delete', paragraph: 1, offset: 1, length: 2 },
        { kind: 'edit', peer: 2, edit: 'bold', paragraph: 1, offset: 2, length: 2 },
        { kind: 'flush' },
        { kind: 'edit', peer: 2, edit: 'undo', paragraph: 1, offset: 1, length: 1 },
        { kind: 'edit', peer: 2, edit: 'redo', paragraph: 0, offset: 3, length: 0 },
      ],
    },
    {
      name: "an undo of a format keeps a peer's delete when a losing split holds a hidden copy",
      seed: 1913,
      actions: [
        { kind: 'edit', peer: 0, edit: 'enter', paragraph: 0, offset: 0, length: 3 },
        { kind: 'edit', peer: 1, edit: 'italic', paragraph: 0, offset: 2, length: 2 },
        { kind: 'edit', peer: 0, edit: 'bold', paragraph: 1, offset: 3, length: 1 },
        { kind: 'flush' },
        { kind: 'edit', peer: 2, edit: 'delete', paragraph: 0, offset: 2, length: 1 },
        { kind: 'deliver', from: 2, to: 1, count: 1 },
        { kind: 'edit', peer: 1, edit: 'undo', paragraph: 2, offset: 0, length: 3 },
      ],
    },
    {
      name: 'a redo of a format after its undo moved the text keeps the text',
      seed: 1359,
      actions: [
        { kind: 'edit', peer: 2, edit: 'type', paragraph: 1, offset: 3, length: 0, letter: 'Ā' },
        { kind: 'edit', peer: 2, edit: 'type', paragraph: 0, offset: 4, length: 0, letter: 'ā' },
        {
          kind: 'edit',
          peer: 2,
          edit: 'color',
          paragraph: 1,
          offset: 4,
          length: 1,
          color: 'accent2',
        },
        { kind: 'edit', peer: 1, edit: 'join', paragraph: 0, offset: 7, length: 0 },
        { kind: 'flush' },
        { kind: 'edit', peer: 0, edit: 'type', paragraph: 0, offset: 6, length: 2, letter: 'Ą' },
        {
          kind: 'edit',
          peer: 0,
          edit: 'color',
          paragraph: 0,
          offset: 13,
          length: 3,
          color: 'accent1',
        },
        { kind: 'edit', peer: 2, edit: 'undo', paragraph: 0, offset: 6, length: 0 },
        { kind: 'flush' },
        { kind: 'edit', peer: 2, edit: 'redo', paragraph: 0, offset: 0, length: 2 },
      ],
    },
    {
      name: 'an undo that Yjs applies to an older step keeps text a peer deleted deleted',
      seed: 1146,
      actions: [
        { kind: 'edit', peer: 2, edit: 'deleteAcross', paragraph: 0, offset: 2, length: 0 },
        { kind: 'edit', peer: 1, edit: 'type', paragraph: 0, offset: 2, length: 0, letter: 'Ā' },
        { kind: 'edit', peer: 1, edit: 'join', paragraph: 0, offset: 8, length: 0 },
        { kind: 'edit', peer: 1, edit: 'delete', paragraph: 0, offset: 2, length: 2 },
        { kind: 'edit', peer: 1, edit: 'bold', paragraph: 0, offset: 7, length: 2 },
        { kind: 'flush' },
        { kind: 'edit', peer: 1, edit: 'bold', paragraph: 0, offset: 2, length: 2 },
        { kind: 'edit', peer: 0, edit: 'bold', paragraph: 0, offset: 4, length: 0 },
        { kind: 'edit', peer: 2, edit: 'enter', paragraph: 0, offset: 4, length: 1, over: true },
        {
          kind: 'edit',
          peer: 1,
          edit: 'color',
          paragraph: 0,
          offset: 3,
          length: 3,
          color: 'accent2',
        },
        { kind: 'edit', peer: 1, edit: 'undo', paragraph: 0, offset: 0, length: 0 },
        { kind: 'flush' },
        { kind: 'edit', peer: 1, edit: 'undo', paragraph: 1, offset: 3, length: 2 },
        { kind: 'edit', peer: 1, edit: 'undo', paragraph: 0, offset: 0, length: 0 },
      ],
    },
    {
      name: "a redo of a delete reaches a peer's format of the deleted text",
      seed: 237,
      actions: [
        { kind: 'edit', peer: 1, edit: 'italic', paragraph: 0, offset: 2, length: 2 },
        { kind: 'edit', peer: 2, edit: 'delete', paragraph: 0, offset: 3, length: 1 },
        { kind: 'edit', peer: 2, edit: 'undo', paragraph: 0, offset: 8, length: 0 },
        { kind: 'edit', peer: 2, edit: 'redo', paragraph: 0, offset: 5, length: 3 },
        { kind: 'flush' },
      ],
    },
    {
      name: "a redo of a delete across paragraphs reaches a peer's format of the text",
      seed: 1885,
      actions: [
        {
          kind: 'edit',
          peer: 2,
          edit: 'color',
          paragraph: 0,
          offset: 1,
          length: 1,
          color: 'accent2',
        },
        { kind: 'edit', peer: 1, edit: 'deleteAcross', paragraph: 0, offset: 1, length: 2 },
        { kind: 'edit', peer: 1, edit: 'undo', paragraph: 0, offset: 2, length: 1 },
        { kind: 'edit', peer: 1, edit: 'redo', paragraph: 1, offset: 2, length: 0 },
        { kind: 'flush' },
      ],
    },
    {
      name: "a delete of text an undo put back reaches a peer's format of the text",
      seed: 1919,
      actions: [
        { kind: 'edit', peer: 1, edit: 'delete', paragraph: 0, offset: 3, length: 3 },
        { kind: 'edit', peer: 1, edit: 'undo', paragraph: 0, offset: 0, length: 3 },
        {
          kind: 'edit',
          peer: 1,
          edit: 'type',
          paragraph: 0,
          offset: 1,
          length: 3,
          over: true,
          letter: 'ă',
        },
        { kind: 'edit', peer: 0, edit: 'italic', paragraph: 0, offset: 1, length: 1 },
        { kind: 'flush' },
      ],
    },
  ],
};

/** Histories that fail today, each with its issue and the problem it pins. */
const KNOWN_FAILURES: Readonly<Record<Mode, readonly FuzzKnownFailure<TextAction>[]>> = {
  text: [
    {
      name: 'an undo of a delete across a run that peers moved keeps text it did not delete',
      issue: '#3',
      seed: 1185,
      problem: 'p0 undo removed characters it may not remove: "L"',
      actions: [
        { kind: 'edit', peer: 0, edit: 'type', paragraph: 0, offset: 2, length: 0, letter: 'Ă' },
        {
          kind: 'edit',
          peer: 0,
          edit: 'type',
          paragraph: 0,
          offset: 9,
          length: 1,
          over: true,
          letter: 'ă',
        },
        { kind: 'deliver', from: 0, to: 1, count: 1 },
        { kind: 'edit', peer: 2, edit: 'enter', paragraph: 0, offset: 15, length: 2 },
        { kind: 'deliver', from: 0, to: 1, count: 1 },
        { kind: 'edit', peer: 1, edit: 'join', paragraph: 0, offset: 14, length: 0 },
        { kind: 'edit', peer: 1, edit: 'enter', paragraph: 0, offset: 10, length: 2 },
        { kind: 'flush' },
        { kind: 'edit', peer: 1, edit: 'join', paragraph: 2, offset: 3, length: 0 },
        { kind: 'flush' },
        { kind: 'edit', peer: 1, edit: 'join', paragraph: 0, offset: 10, length: 1 },
        { kind: 'edit', peer: 2, edit: 'deleteAcross', paragraph: 0, offset: 6, length: 1 },
        { kind: 'edit', peer: 0, edit: 'type', paragraph: 1, offset: 1, length: 1, letter: 'ć' },
        { kind: 'edit', peer: 0, edit: 'deleteAcross', paragraph: 0, offset: 0, length: 3 },
        { kind: 'edit', peer: 2, edit: 'undo', paragraph: 0, offset: 13, length: 3 },
        { kind: 'flush' },
        { kind: 'edit', peer: 1, edit: 'delete', paragraph: 0, offset: 1, length: 2 },
        { kind: 'deliver', from: 1, to: 0, count: 3 },
        { kind: 'edit', peer: 0, edit: 'undo', paragraph: 0, offset: 2, length: 3 },
      ],
    },
  ],
  format: [
    {
      name: 'two formats of one run after a peer typed into it show each character once',
      issue: '#16',
      seed: 42,
      problem: 'characters show more than once: "JKL" in ["ABCDEFGH","IJKĀLJKL"]',
      actions: [
        { kind: 'edit', peer: 0, edit: 'type', paragraph: 1, offset: 3, length: 1, letter: 'Ā' },
        { kind: 'edit', peer: 2, edit: 'bold', paragraph: 1, offset: 0, length: 1 },
        { kind: 'edit', peer: 0, edit: 'bold', paragraph: 1, offset: 1, length: 1 },
        { kind: 'flush' },
      ],
    },
    {
      name: 'an edit at a run that a peer splits at the same time lands in place',
      issue: '#17',
      seed: 71,
      problem: 'characters show more than once: "Ă" in ["ABĂ","CĂDEFGH","IJKL"]',
      actions: [
        { kind: 'edit', peer: 2, edit: 'type', paragraph: 0, offset: 3, length: 0, letter: 'Ă' },
        { kind: 'edit', peer: 0, edit: 'enter', paragraph: 0, offset: 2, length: 3 },
        { kind: 'deliver', from: 2, to: 0, count: 1 },
        {
          kind: 'edit',
          peer: 0,
          edit: 'color',
          paragraph: 0,
          offset: 0,
          length: 2,
          color: 'accent6',
        },
        { kind: 'flush' },
      ],
    },
    {
      name: 'an undo of a format keeps text that a peer formatted at the same time',
      issue: '#3',
      seed: 1178,
      problem: 'characters that no peer removed are lost: "C" in ["A","BKL"]',
      actions: [
        { kind: 'edit', peer: 2, edit: 'deleteAcross', paragraph: 0, offset: 3, length: 2 },
        { kind: 'flush' },
        { kind: 'edit', peer: 2, edit: 'bold', paragraph: 1, offset: 2, length: 1 },
        { kind: 'edit', peer: 0, edit: 'enter', paragraph: 1, offset: 1, length: 2 },
        { kind: 'edit', peer: 0, edit: 'bold', paragraph: 1, offset: 1, length: 3 },
        { kind: 'edit', peer: 0, edit: 'undo', paragraph: 2, offset: 0, length: 0 },
      ],
    },
  ],
};

/**
 * Default seeds that fail today, each with the issue that its shrunk history shows and the
 * problem it pins. Remove a seed when its issue is fixed: the test then says so.
 */
const KNOWN_SEEDS: Readonly<Record<Mode, Readonly<Record<number, FuzzKnownSeed>>>> = {
  text: {},
  format: {
    4: {
      issue: '#16',
      problem: 'characters show more than once: "KĂ" in ["āABąFG","ĄăIJKĂKĂ",""]',
    },
    7: {
      issue: '#16',
      problem: 'characters show more than once: "Ā" in ["ąăĀĄā","Ā",""]',
    },
    17: {
      issue: '#16',
      problem: 'characters show more than once: "Ā" in ["ABCĂEFH","","IĀJāĀ",""]',
    },
    20: {
      issue: '#3',
      problem: 'p1 join paragraph 1 [0, 0) removed characters it may not remove: "Ąă"',
    },
  },
};

const TITLES: Readonly<Record<Mode, string>> = {
  text: 'typing, deletes, Enter, joins, undo, and redo by three peers keep each character once',
  format: 'with bold, italic, and color, three peers keep each character once',
};

for (const mode of ['text', 'format'] as const)
  fuzzSuite<TextAction>({
    name: `${mode} fuzz`,
    title: TITLES[mode],
    knobs: KNOBS[mode],
    file: FILE,
    peers: PEERS,
    run: (seed) => runSeed(seed, mode),
    replay: async (clients, actions) => perform(await TextRun.open(clients, mode), actions),
    regressions: [
      ...(mode === 'text'
        ? [...TEXT_HISTORIES, ...UNDO_HISTORIES]
        : [...FORMAT_HISTORIES, ...LONG_FORMAT_HISTORIES]),
      ...FIXED[mode],
    ],
    knownFailures: mode === 'text' ? KNOWN_FAILURES.text : KNOWN_FAILURES.format,
    knownSeeds: KNOWN_SEEDS[mode],
    summary: () => `${mode} fuzz checks: ${JSON.stringify(STATS[mode])}`,
  });
