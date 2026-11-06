/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Seeded fuzz of concurrent table edits. Three editors
// on three Y.Docs insert and delete rows and columns, merge and split cells, type in cells,
// delete text, undo, and redo in one 3 x 3 table. The relay holds every update back and
// delivers parts of it at random; a round ends with a flush, or keeps the held updates for the
// next round. The harness is `fuzz-support.ts`; its header lists the environment knobs (prefix
// `TABLE`).
//
// Invariants:
// - Each edit changes only what it may on its own replica: typing adds its letter, a text
//   delete removes exactly its selection, a row or column delete removes only the text of the
//   cells it covers, an insert, a merge, and a split remove nothing, and an undo or redo
//   removes only characters its peer placed or the cells it removes held, and brings back only
//   characters its peer removed, that a concurrent edit took, or that the cells it restores or
//   uncovers hold. A caret in a covered position goes to the merged cell that covers it.
// - After each flush, every session is `ready`, and the replicas hold the same package.
// - Text conservation (`fuzz-ledger.ts`): each character shows at most once, shows unless a
//   peer removed it, and stays removed unless a peer brought it back. A character placed at the
//   same time as a peer's delete of its row or column may go with it. A merge copies the text
//   of the cells it covers, so a concurrent delete of that text may lose.
// - After each flush, the canvas grid (the table that layout draws) equals the grid of the
//   saved deck reopened in a fresh editor, on every replica.
// At the end, each replica's save reopens to the same deck, a second save keeps it, and the
// saved deck passes the part checks (`table-merge` included).

import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { textOf, type Element } from '@pptx-editor.dev/core/store';
import { buildPptx, cell, table } from '../../../../scripts/fixtures/builder.ts';
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
  type FuzzKnownFailure,
  type FuzzKnownSeed,
  type FuzzRegression,
} from './fuzz-support.ts';
import { IN, all, kids, root } from './lineage-support.ts';

const FILE = 'packages/pro/src/__tests__/yjs-table-fuzz.test.ts';
const KNOBS = fuzzKnobs('TABLE', 16);
const PEERS = 3;
const ROUNDS = 4;

const DECK_ROWS = ['ABC', 'DEF', 'GHI'];

/**
 * The deck of one fuzz mode: a 3 x 3 table with a letter in each cell. The cells of `bodiless`
 * have no text body (`a:txBody` is optional, ECMA-376 Part 1, 21.1.3.16 tc), so peers that type
 * into one at once each give it a body.
 */
interface FuzzDeck {
  readonly bodiless: string;
}
const FULL: FuzzDeck = { bodiless: '' };
const BODILESS: FuzzDeck = { bodiless: 'E' };

const deckText = (d: FuzzDeck): string =>
  [...DECK_ROWS.join('')].filter((c) => !d.bodiless.includes(c)).join('');

const deck = (d: FuzzDeck): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: table({
          id: 2,
          x: IN,
          y: IN,
          cx: 3 * IN,
          cy: 3 * 370_840,
          columns: [IN, IN, IN],
          tblPr: '<a:tblPr firstRow="1" bandRow="1"/>',
          rows: DECK_ROWS.map(
            (r) =>
              `<a:tr h="370840">${[...r]
                .map((c) => (d.bodiless.includes(c) ? '<a:tc><a:tcPr/></a:tc>' : cell(c)))
                .join('')}</a:tr>`
          ).join(''),
        }),
      },
    ],
  });

type EditKind =
  | 'insertRow'
  | 'insertColumn'
  | 'deleteRow'
  | 'deleteColumn'
  | 'merge'
  | 'split'
  | 'type'
  | 'deleteText'
  | 'undo'
  | 'redo';
const EDITS: readonly EditKind[] = [
  'insertRow',
  'insertColumn',
  'deleteRow',
  'deleteColumn',
  'merge',
  'merge',
  'split',
  'type',
  'type',
  'type',
  'deleteText',
  'undo',
  'redo',
];
const SPLITS = [
  { rows: 1, cols: 2 },
  { rows: 2, cols: 1 },
  { rows: 2, cols: 2 },
] as const;

/** One step, in grid coordinates, so a shrunk list stays meaningful. */
export type TableAction =
  | {
      readonly kind: 'edit';
      readonly peer: number;
      readonly edit: EditKind;
      readonly row: number;
      readonly column: number;
      /** The other corner of a merge. */
      readonly row2?: number;
      readonly column2?: number;
      /** `insertRow`: above; `insertColumn`: left. */
      readonly before?: boolean;
      readonly split?: (typeof SPLITS)[number];
      readonly offset?: number;
      readonly length?: number;
      /** The character `type` types: each one is new. */
      readonly letter?: string;
    }
  | Delivery
  /** Deliver every waiting update; the round's checks run. */
  | { readonly kind: 'flush' };

/** The slide's graphic frame (PresentationML, so `all` does not find it). */
function frameOf(e: PptxEditorInstance): Element | undefined {
  const find = (el: Element): Element | undefined => {
    for (const c of kids(el)) {
      const hit = c.kind === 'graphicFrame' ? c : find(c);
      if (hit) return hit;
    }
    return undefined;
  };
  return find(root(e));
}

/** The table that layout draws for `e`: its size and its cells. */
function layoutTable(e: PptxEditorInstance) {
  const item = e.session!.layout(0)?.items.find((i) => i.element.kind === 'graphicFrame');
  return item?.element.kind === 'graphicFrame' ? item.element.table : undefined;
}

/**
 * The table as `e` draws it: per grid row, each drawn cell as its text with its spans when
 * they are not 1 x 1, and `·` for a covered position. A blank cell that the save adds to a short
 * row has no node in the store and draws no text; `?` marks a drawn node that the store lacks.
 */
function drawn(e: PptxEditorInstance): string[] {
  const t = layoutTable(e);
  if (!t) return [];
  const byId = new Map(all(root(e), 'tc').map((tc) => [tc.id, tc]));
  const out: string[] = [];
  for (const c of t.cells) {
    if (c.col === 0) out.push('');
    const tc = byId.get(c.nodeId);
    const text = tc ? textOf(tc) : c.nodeId ? '?' : '';
    const spans = c.gridSpan > 1 || c.rowSpan > 1 ? `[${c.rowSpan}x${c.gridSpan}]` : '';
    out[c.row] += `${c.col ? '|' : ''}${c.covered ? '·' : text + spans}`;
  }
  return out;
}

/**
 * The text of the cells that the canvas draws. A covered cell is left out: in a room, a merge
 * keeps a copy of the text it moved in each covered cell, and the save drops the copies.
 */
function tableText(e: PptxEditorInstance): string {
  const shown = new Set(
    (layoutTable(e)?.cells ?? []).filter((c) => !c.covered).map((c) => c.nodeId)
  );
  return all(root(e), 'tc')
    .filter((tc) => shown.has(tc.id))
    .map((tc) => textOf(tc))
    .join('');
}

/** The text of the first paragraph of the cell at a grid position. */
function cellParagraph(e: PptxEditorInstance, row: number, column: number): string {
  const t = layoutTable(e)!;
  const c = t.cells[row * t.columns.length + column];
  const tc = c && all(root(e), 'tc').find((x) => x.id === c.nodeId);
  const p = tc && all(tc, 'p')[0];
  return p ? textOf(p) : '';
}

/** The text of each drawn cell; a covered cell is left out (see `tableText`). */
function cellsOf(e: PptxEditorInstance): string[] {
  const shown = new Set(
    (layoutTable(e)?.cells ?? []).filter((c) => !c.covered).map((c) => c.nodeId)
  );
  return all(root(e), 'tc')
    .filter((tc) => shown.has(tc.id))
    .map((tc) => textOf(tc));
}

/** The text of the drawn cells that a delete of rows (or columns) `from`..`to` covers. */
function lineText(e: PptxEditorInstance, column: boolean, from: number, to: number): Set<string> {
  const byId = new Map(all(root(e), 'tc').map((tc) => [tc.id, tc]));
  const out = new Set<string>();
  for (const c of layoutTable(e)?.cells ?? []) {
    if (c.covered) continue;
    const [start, span] = column ? [c.col, c.gridSpan] : [c.row, c.rowSpan];
    if (start + span <= from || start >= to) continue;
    const tc = byId.get(c.nodeId);
    for (const ch of tc ? textOf(tc) : '') out.add(ch);
  }
  return out;
}

/** The text that a merge of the rectangle from (`row`, `column`) copies into its first cell. */
function mergedText(
  e: PptxEditorInstance,
  action: { row2?: number; column2?: number },
  row: number,
  column: number
): Set<string> {
  const t = layoutTable(e)!;
  const row2 = Math.min(action.row2 ?? row, t.rows.length - 1);
  const column2 = Math.min(action.column2 ?? column, t.columns.length - 1);
  const [r0, r1] = [Math.min(row, row2), Math.max(row, row2)];
  const [c0, c1] = [Math.min(column, column2), Math.max(column, column2)];
  const byId = new Map(all(root(e), 'tc').map((tc) => [tc.id, tc]));
  const out = new Set<string>();
  for (const c of t.cells) {
    if (
      c.covered ||
      c.row > r1 ||
      c.col > c1 ||
      c.row + c.rowSpan <= r0 ||
      c.col + c.gridSpan <= c0
    )
      continue;
    const tc = byId.get(c.nodeId);
    for (const ch of tc ? textOf(tc) : '') out.add(ch);
  }
  return out;
}

/** The text of every cell, covered ones too, by node id. */
const cellTexts = (e: PptxEditorInstance): Map<string, string> =>
  new Map(all(root(e), 'tc').map((tc) => [tc.id, textOf(tc)]));

/** How often each check ran. */
const STATS = { rounds: 0, edits: 0, characterChecks: 0, exempt: 0, blind: 0, gridChecks: 0 };

type TableEdit = Extract<TableAction, { kind: 'edit' }>;

/** Three replicas, and what the conservation check knows. */
class TableRun {
  readonly log: FuzzLog;
  private readonly ledger: CharacterLedger;

  private constructor(
    d: FuzzDeck,
    private readonly editors: readonly PptxEditorInstance[],
    readonly relay: Awaited<ReturnType<typeof openFuzzPeers>>['relay'],
    private readonly docs: Awaited<ReturnType<typeof openFuzzPeers>>['docs']
  ) {
    this.ledger = new CharacterLedger(deckText(d));
    this.log = new FuzzLog(KNOBS.trace, () =>
      editors.map((e, i) => `    p${i}: ${JSON.stringify(drawn(e))}`).join('\n')
    );
  }

  static async open(d: FuzzDeck, clients: readonly number[]): Promise<TableRun> {
    const { editors, relay, docs } = await openFuzzPeers(deck(d), clients);
    return new TableRun(d, editors, relay, docs);
  }

  destroy(): void {
    STATS.characterChecks += this.ledger.stats.characterChecks;
    STATS.exempt += this.ledger.stats.exempt;
    STATS.blind += this.ledger.stats.blind;
    for (const e of this.editors) e.destroy();
  }

  async perform(action: TableAction): Promise<void> {
    if (action.kind === 'deliver') {
      const n = this.relay.deliver(action.from, action.to, action.count);
      if (n > 0) this.log.note(`deliver ${n} p${action.from} -> p${action.to}`);
      return;
    }
    if (action.kind === 'flush') {
      this.relay.flush();
      this.log.note('flush');
      await this.checkRound();
      return;
    }
    this.edit(action);
  }

  private edit(action: TableEdit): void {
    const e = this.editors[action.peer]!;
    const frame = frameOf(e);
    const t = layoutTable(e);
    if (!frame || !t) return;
    const nR = t.rows.length;
    const nC = t.columns.length;
    // A caret in a covered position goes to the merged cell that covers it, so the edit
    // addresses the cell it changes: the drawn cell, its text, and all the rows or columns of
    // its span.
    const at = { row: Math.min(action.row, nR - 1), column: Math.min(action.column, nC - 1) };
    const target =
      t.cells.find(
        (c) =>
          !c.covered &&
          c.row <= at.row &&
          at.row < c.row + c.rowSpan &&
          c.col <= at.column &&
          at.column < c.col + c.gridSpan
      ) ?? t.cells[at.row * nC + at.column]!;
    const caretAt = action.edit !== 'merge';
    const row = caretAt ? target.row : at.row;
    const column = caretAt ? target.col : at.column;
    const lines =
      action.edit === 'deleteColumn'
        ? { from: column, to: column + target.gridSpan, of: nC }
        : { from: row, to: row + target.rowSpan, of: nR };
    const caret = (from: number, to = from) =>
      e.setSelection({
        kind: 'text',
        nodeId: frame.id,
        cell: { row, column },
        anchor: { paragraph: 0, offset: from },
        focus: { paragraph: 0, offset: to },
      });
    const paragraph = cellParagraph(e, row, column);
    let what = `${action.edit} (${row}, ${column})`;
    let selected = '';
    let run: () => { ok: boolean };
    switch (action.edit) {
      case 'insertRow':
        caret(0);
        what = `insertRow ${action.before ? 'above' : 'below'} (${row}, ${column})`;
        run = () => e.exec({ type: 'insertRow', where: action.before ? 'above' : 'below' });
        break;
      case 'insertColumn':
        caret(0);
        what = `insertColumn ${action.before ? 'left' : 'right'} (${row}, ${column})`;
        run = () => e.exec({ type: 'insertColumn', where: action.before ? 'left' : 'right' });
        break;
      case 'deleteRow':
      case 'deleteColumn':
        // Keep the table: a delete of all its rows or columns would delete it.
        if (lines.to - lines.from >= lines.of) return;
        caret(0);
        run = () => e.exec({ type: action.edit as 'deleteRow' | 'deleteColumn' });
        break;
      case 'merge': {
        const row2 = Math.min(action.row2 ?? row, nR - 1);
        const column2 = Math.min(action.column2 ?? column, nC - 1);
        if (row2 === row && column2 === column) return;
        what = `merge (${row}, ${column})..(${row2}, ${column2})`;
        e.setSelection({
          kind: 'cells',
          nodeId: frame.id,
          anchor: { row, column },
          focus: { row: row2, column: column2 },
        });
        run = () => e.exec({ type: 'mergeCells' });
        break;
      }
      case 'split': {
        const split = action.split ?? SPLITS[0];
        what += ` into ${split.rows} x ${split.cols}`;
        caret(0);
        run = () => e.exec({ type: 'splitCell', ...split });
        break;
      }
      case 'type': {
        const at = Math.min(action.offset ?? 0, paragraph.length);
        what = `type ${JSON.stringify(action.letter)} at (${row}, ${column}) ${at}`;
        caret(at);
        run = () => e.input('text', action.letter);
        break;
      }
      case 'deleteText': {
        if (paragraph.length === 0) return;
        const from = Math.min(action.offset ?? 0, paragraph.length - 1);
        const to = Math.min(paragraph.length, from + Math.max(1, action.length ?? 1));
        what = `delete text at (${row}, ${column}) [${from}, ${to})`;
        selected = paragraph.slice(from, to);
        caret(from, to);
        run = () => e.input('deleteBackward');
        break;
      }
      case 'undo':
      case 'redo':
        what = action.edit;
        run = () => e.exec({ type: action.edit as 'undo' | 'redo' });
        break;
    }
    const lineDelete = action.edit === 'deleteRow' || action.edit === 'deleteColumn';
    const line = lineDelete
      ? lineText(e, action.edit === 'deleteColumn', lines.from, lines.to)
      : NONE;
    const before = new Set(tableText(e));
    const cellsBefore = cellTexts(e);
    const moved = action.edit === 'merge' ? mergedText(e, action, row, column) : NONE;
    const doc = this.docs[action.peer]!;
    const state = stateOf(doc);
    const result = run();
    const clock = editClock(doc, state);
    const cells = cellsOf(e);
    const after = new Set(cells.join(''));
    const gone = new Set([...before].filter((ch) => !after.has(ch)));
    const came = new Set([...after].filter((ch) => !before.has(ch)));
    what = `p${action.peer} ${what}`;
    this.log.note(`${what} ${result.ok ? 'ok' : 'refused'}`);
    STATS.edits += 1;
    // An undo or a redo of a structure edit may remove the cells it removes, with their text,
    // and may uncover cells that a merge covered, with theirs.
    const cellsAfter = cellTexts(e);
    const structure = {
      vanished: new Set(
        [...cellsBefore].filter(([id]) => !cellsAfter.has(id)).flatMap(([, x]) => [...x])
      ),
      uncovered: new Set(
        [...cellsBefore.values()].flatMap((x) => [...x]).filter((ch) => !before.has(ch))
      ),
      appeared: new Set(
        [...cellsAfter].filter(([id]) => !cellsBefore.has(id)).flatMap(([, x]) => [...x])
      ),
    };
    const rule = withoutRepeated(
      this.rule(action, doc.clientID, result.ok, new Set(selected), line, structure),
      [...cellsBefore.values()].join('')
    );
    for (const problem of localProblems(what, gone, came, rule)) this.log.fail(problem);
    // The place of a character that showed: the rest of its cell.
    const placeOf = (ch: string) =>
      new Set([...(cells.find((x) => x.includes(ch)) ?? '')].filter((c) => !came.has(c)));
    this.ledger.record(clock, gone, came, placeOf);
    if (lineDelete && result.ok) this.ledger.risk({ clock, region: line, view: before });
    if (action.edit === 'merge' && result.ok) this.ledger.move(clock, moved);
    // An undo of an insert, or a redo of a delete, removes cells as a delete does.
    const history = action.edit === 'undo' || action.edit === 'redo';
    if (history && result.ok && gone.size > 0)
      this.ledger.risk({ clock, region: gone, view: before });
  }

  /** What one edit may change on its own replica. */
  private rule(
    action: TableEdit,
    client: number,
    ok: boolean,
    selected: ReadonlySet<string>,
    line: ReadonlySet<string>,
    structure: {
      vanished: ReadonlySet<string>;
      uncovered: ReadonlySet<string>;
      appeared: ReadonlySet<string>;
    }
  ): LocalRule {
    switch (action.edit) {
      case 'type': {
        const letter = new Set(action.letter ? [action.letter] : []);
        return { mayRemove: NONE, mayAdd: letter, ...(ok ? { mustAdd: letter } : {}) };
      }
      case 'deleteText':
        return { mayRemove: selected, mayAdd: NONE, ...(ok ? { mustRemove: selected } : {}) };
      case 'deleteRow':
      case 'deleteColumn':
        return { mayRemove: line, mayAdd: NONE };
      case 'undo':
      case 'redo':
        return {
          mayRemove: new Set([...this.ledger.placedBy(client), ...structure.vanished]),
          mayAdd: new Set([
            ...this.ledger.mayBringBack(client),
            ...structure.uncovered,
            ...structure.appeared,
          ]),
        };
      default:
        return { mayRemove: NONE, mayAdd: NONE };
    }
  }

  /** After a flush: sessions ready, replicas equal, text conservation, and the grids. */
  private async checkRound(): Promise<void> {
    for (const problem of roundProblems(this.editors)) this.log.fail(problem);
    if (this.log.failed) return;
    STATS.rounds += 1;
    const first = this.editors[0]!;
    for (const problem of this.ledger.problems(tableText(first)))
      this.log.fail(`${problem} in ${JSON.stringify(drawn(first))}`);
    if (!this.log.failed) await this.checkGrids();
  }

  /** On every replica, the canvas grid equals the grid of its saved deck in a fresh editor. */
  private async checkGrids(): Promise<void> {
    for (const [i, e] of this.editors.entries()) {
      const reopened = createPptxEditor({ fixedMeasurer: true, substitutes: false });
      try {
        const bytes = e.save();
        if (!bytes || !(await reopened.load(bytes)).ok) {
          this.log.fail(`replica ${i} saved a deck that does not open`);
          return;
        }
        STATS.gridChecks += 1;
        const canvas = drawn(e);
        const saved = drawn(reopened);
        if (JSON.stringify(canvas) !== JSON.stringify(saved))
          this.log.fail(
            `replica ${i} draws another grid than its saved deck: canvas ` +
              `${JSON.stringify(canvas)}, saved ${JSON.stringify(saved)}`
          );
      } finally {
        reopened.destroy();
      }
    }
  }

  /** The end-of-run invariants of `fuzz-support.ts`, after the round checks. */
  async finish(): Promise<void> {
    this.relay.flush();
    await this.checkRound();
    if (this.log.failed) return;
    for (const problem of convergenceProblems(this.editors).problems) this.log.fail(problem);
  }

  /** A random edit on `peer`'s table. */
  randomEdit(peer: number, random: () => number, letter: () => string): TableAction | null {
    const t = layoutTable(this.editors[peer]!);
    if (!t) return null;
    const pick = (n: number) => Math.floor(random() * n);
    const nR = t.rows.length;
    const nC = t.columns.length;
    const edit = EDITS[pick(EDITS.length)]!;
    const base = { kind: 'edit', peer, edit, row: pick(nR), column: pick(nC) } as const;
    switch (edit) {
      case 'insertRow':
      case 'insertColumn':
        return { ...base, before: random() < 0.5 };
      case 'merge':
        return { ...base, row2: pick(nR), column2: pick(nC) };
      case 'split':
        return { ...base, split: SPLITS[pick(SPLITS.length)]! };
      case 'type':
        return { ...base, offset: pick(3), letter: letter() };
      case 'deleteText':
        return { ...base, offset: pick(3), length: 1 + pick(3) };
      default:
        return base;
    }
  }
}

/** Run `actions` on fresh replicas, then the end checks. */
async function perform(run: TableRun, actions: Iterable<TableAction>): Promise<string[]> {
  try {
    for (const action of actions) {
      await run.perform(action);
      if (run.log.failed) break;
    }
    if (!run.log.failed) await run.finish();
  } catch (error) {
    run.log.fail(`threw ${String(error)}`);
  } finally {
    run.destroy();
  }
  return run.log.problems;
}

/** Generate and perform one seed's actions. */
async function runSeed(d: FuzzDeck, seed: number) {
  const random = seededRandom(seed);
  const run = await TableRun.open(d, clientIdsOf(PEERS, seed));
  const actions: TableAction[] = [];
  let n = 0;
  const letter = () => String.fromCharCode(0x100 + n++);
  // The generator reads each replica after the actions before it, so it yields as they run.
  function* generate(): Generator<TableAction> {
    for (let round = 0; round < ROUNDS; round += 1) {
      const steps = 3 + Math.floor(random() * 4);
      for (let step = 0; step < steps; step += 1) {
        const edit = run.randomEdit(Math.floor(random() * PEERS), random, letter);
        if (edit) yield edit;
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
  return { problems: await perform(run, recorded), actions };
}

/** Fixed histories that once failed. Add a shrunk seed here when its issue is fixed. */
const REGRESSIONS: readonly FuzzRegression<TableAction>[] = [
  {
    name: 'concurrent deletes of the last two columns after a row insert save a valid table',
    seed: 291,
    actions: [
      { kind: 'edit', peer: 2, edit: 'deleteColumn', row: 3, column: 1 },
      { kind: 'flush' },
      { kind: 'edit', peer: 0, edit: 'insertRow', row: 2, column: 0, before: true },
      { kind: 'edit', peer: 0, edit: 'deleteColumn', row: 1, column: 1 },
      { kind: 'edit', peer: 2, edit: 'deleteColumn', row: 2, column: 0 },
    ],
  },
  {
    name: 'a row insert and a concurrent column insert draw the grid that the save writes',
    seed: 241,
    actions: [
      { kind: 'edit', peer: 0, edit: 'insertRow', row: 2, column: 2, before: false },
      { kind: 'edit', peer: 1, edit: 'insertColumn', row: 2, column: 2, before: false },
    ],
  },
  {
    name: 'a merge and a concurrent delete of its first column keep the other merged text',
    seed: 18,
    actions: [
      { kind: 'edit', peer: 1, edit: 'deleteColumn', row: 0, column: 0 },
      { kind: 'edit', peer: 2, edit: 'merge', row: 0, column: 1, row2: 2, column2: 0 },
      { kind: 'flush' },
    ],
  },
  {
    name: 'a merge and a concurrent row insert inside it show the merged text once',
    seed: 2,
    actions: [
      { kind: 'edit', peer: 0, edit: 'merge', row: 2, column: 1, row2: 0, column2: 1 },
      { kind: 'edit', peer: 2, edit: 'insertRow', row: 2, column: 0, before: true },
      { kind: 'flush' },
    ],
  },
  {
    name: 'text typed into a cell that a peer merges at the same time is kept',
    seed: 27,
    actions: [
      { kind: 'edit', peer: 0, edit: 'insertRow', row: 0, column: 2, before: true },
      { kind: 'deliver', from: 0, to: 1, count: 1 },
      { kind: 'edit', peer: 1, edit: 'merge', row: 1, column: 1, row2: 0, column2: 0 },
      { kind: 'edit', peer: 0, edit: 'type', row: 0, column: 0, offset: 0, letter: 'ā' },
      { kind: 'flush' },
    ],
  },
  {
    name: 'an undo of a merge keeps the merged text after concurrent row and column deletes',
    seed: 16,
    actions: [
      { kind: 'edit', peer: 1, edit: 'deleteRow', row: 2, column: 1 },
      { kind: 'edit', peer: 2, edit: 'merge', row: 1, column: 2, row2: 0, column2: 2 },
      { kind: 'edit', peer: 1, edit: 'deleteRow', row: 0, column: 2 },
      { kind: 'edit', peer: 0, edit: 'deleteColumn', row: 0, column: 0 },
      { kind: 'deliver', from: 1, to: 0, count: 3 },
      { kind: 'deliver', from: 2, to: 0, count: 1 },
      { kind: 'edit', peer: 0, edit: 'merge', row: 1, column: 1, row2: 1, column2: 0 },
      { kind: 'flush' },
      { kind: 'edit', peer: 2, edit: 'undo', row: 1, column: 1 },
    ],
  },
  {
    name: 'a split after a merge that moved typed text leaves the same table on every replica',
    seed: 460,
    actions: [
      { kind: 'edit', peer: 0, edit: 'split', row: 3, column: 1, split: { rows: 2, cols: 1 } },
      { kind: 'edit', peer: 1, edit: 'deleteColumn', row: 2, column: 2 },
      { kind: 'deliver', from: 0, to: 1, count: 2 },
      { kind: 'edit', peer: 1, edit: 'split', row: 4, column: 0, split: { rows: 1, cols: 2 } },
      { kind: 'flush' },
      { kind: 'edit', peer: 2, edit: 'merge', row: 3, column: 4, row2: 1, column2: 1 },
      { kind: 'edit', peer: 1, edit: 'type', row: 1, column: 4, offset: 0, letter: 'ā' },
      { kind: 'flush' },
      { kind: 'edit', peer: 2, edit: 'merge', row: 0, column: 4, row2: 2, column2: 3 },
      { kind: 'edit', peer: 2, edit: 'undo', row: 3, column: 3 },
      { kind: 'edit', peer: 2, edit: 'split', row: 1, column: 2, split: { rows: 1, cols: 2 } },
    ],
  },
];
/** Histories that fail today, each with its issue and the problem it pins. */
const KNOWN_FAILURES: readonly FuzzKnownFailure<TableAction>[] = [];

/**
 * Default seeds that fail today, each with the issue that its shrunk history shows and the
 * problem it pins. Remove a seed when its issue is fixed: the test then says so.
 */
const KNOWN_SEEDS: Readonly<Record<number, FuzzKnownSeed>> = {
  1: {
    issue: '#1',
    problem: 'characters that no peer removed are lost: "G" in ["B|C[1x2]|·","H|I|"]',
  },
  10: {
    issue: '#2',
    problem:
      'characters that no peer removed are lost: "Ă" in ["||[1x2]|·","A|CF[2x2]|·|","D|·|·|","[1x2]|·|I|"]',
  },
};

fuzzSuite<TableAction>({
  name: 'table fuzz',
  title: 'concurrent structure and text edits by three peers',
  knobs: KNOBS,
  file: FILE,
  peers: PEERS,
  run: (seed) => runSeed(FULL, seed),
  replay: async (clients, actions) => perform(await TableRun.open(FULL, clients), actions),
  regressions: REGRESSIONS,
  knownFailures: KNOWN_FAILURES,
  knownSeeds: KNOWN_SEEDS,
  summary: () => `table fuzz checks: ${JSON.stringify(STATS)}`,
});

// The same fuzz on a table whose center cell has no text body. It has its own seeds and
// pins, so the pins of the mode above stay valid. Its knobs have the prefix `TABLE_BODILESS`.
const BODILESS_KNOBS = fuzzKnobs('TABLE_BODILESS', 16);

/** Default seeds of the bodiless deck that fail today (as `KNOWN_SEEDS`). */
const BODILESS_KNOWN_SEEDS: Readonly<Record<number, FuzzKnownSeed>> = {
  1: {
    issue: '#1',
    problem: 'characters that no peer removed are lost: "G" in ["B|C[1x2]|·","H|I|"]',
  },
  10: {
    issue: '#2',
    problem:
      'characters that no peer removed are lost: "Ă" in ["||[1x2]|·","A|CF[2x2]|·|","D|·|·|","[1x2]|·|I|"]',
  },
};

fuzzSuite<TableAction>({
  name: 'table fuzz, cell without a text body',
  title: 'concurrent structure and text edits by three peers',
  knobs: BODILESS_KNOBS,
  file: FILE,
  peers: PEERS,
  run: (seed) => runSeed(BODILESS, seed),
  replay: async (clients, actions) => perform(await TableRun.open(BODILESS, clients), actions),
  knownSeeds: BODILESS_KNOWN_SEEDS,
});
