/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The shared harness of the seeded collaboration fuzzers (`*-fuzz.test.ts`). Adapted from DOCX
// Editor `packages/pro/src/collaboration/__tests__/document-undo-groups-fuzz.test.ts` (commit
// 5993bd223): `seededRandom` is a copy, and the environment knobs, the trail, the shrink, and
// `kindOf` keep its shape. The peer harness is PPTX-only: N editors on N Y.Docs
// (`openPeer` in `lineage-support.ts`) and a relay that holds every update back until an action
// delivers it. The text conservation oracle is `fuzz-ledger.ts`.
//
// A fuzzer records each step as a plain action in deck coordinates (a paragraph and an offset,
// a row and a column), so a replay of fewer actions stays meaningful. Deliveries are actions
// too: a replay of one action list is deterministic.
//
// Environment, for a fuzzer whose prefix is `TEXT`:
// - `TEXT_FUZZ_SEEDS` and `TEXT_FUZZ_FROM` choose the seeds.
// - `TEXT_FUZZ_SHRINK=<seed>` reduces that seed's failure to the fewest actions and prints them
//   as an entry for the fuzzer's known failures (and, once fixed, its regressions).
// - `TEXT_FUZZ_TRACE=1` prints every action with each replica's state.
// See CONTRIBUTING.md, "Fuzz tests".
//
// Known failures are pinned, not skipped: a known seed or history must still fail with exactly
// the problem it pins. A new problem fails the test, and so does a pass (the issue is fixed:
// move the entry to the regressions).

import { afterAll, describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  readPptx,
  semanticDigest,
  type PptxPackage,
} from '@pptx-editor.dev/core/store';
import { packageViolations } from '../../../core/src/store/store/part-checks.ts';
import { differingParts } from './convergence-harness.ts';
import { openPeer, sessionsOf } from './lineage-support.ts';

/** A small deterministic generator (mulberry32), so a failure names its seed. */
export function seededRandom(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let value = state;
    value = Math.imul(value ^ (value >>> 15), value | 1);
    value ^= value + Math.imul(value ^ (value >>> 7), value | 61);
    return ((value ^ (value >>> 14)) >>> 0) / 4_294_967_296;
  };
}

/** The environment knobs of one fuzzer. */
export interface FuzzKnobs {
  /** The prefix of the variables, for example `TEXT` for `TEXT_FUZZ_SEEDS`. */
  readonly prefix: string;
  readonly seeds: number;
  readonly from: number;
  /** The seed to shrink, from `<PREFIX>_FUZZ_SHRINK`. */
  readonly shrink: number | undefined;
  readonly trace: boolean;
}

function whole(name: string, fallback: number | undefined): number | undefined {
  const raw = process.env[name];
  if (raw === undefined || raw === '') return fallback;
  const value = Number(raw);
  if (!Number.isInteger(value) || value < 0)
    throw new Error(`${name}=${JSON.stringify(raw)}: give a whole number of 0 or more.`);
  return value;
}

/** Read `<prefix>_FUZZ_SEEDS`, `_FROM`, `_SHRINK`, and `_TRACE`. */
export function fuzzKnobs(prefix: string, defaultSeeds: number): FuzzKnobs {
  const name = (knob: string) => `${prefix}_FUZZ_${knob}`;
  return {
    prefix,
    seeds: whole(name('SEEDS'), defaultSeeds)!,
    from: whole(name('FROM'), 1)!,
    shrink: whole(name('SHRINK'), undefined),
    trace: process.env[name('TRACE')] === '1',
  };
}

/** A delivery of the first `count` updates that wait on the link `from` → `to`. */
export interface Delivery {
  readonly kind: 'deliver';
  readonly from: number;
  readonly to: number;
  readonly count: number;
}

/** A delivery of every update that waits for `to`, in the order they were sent. */
export interface DeliveryTo {
  readonly kind: 'deliverTo';
  readonly to: number;
}

/** An update that waits on a link, with its place in the order of all sent updates. */
interface Held {
  readonly sent: number;
  readonly update: Uint8Array;
}

/**
 * Links between every pair of N Y.Docs that hold each update back until an action delivers it.
 * Each link keeps its updates in order, as a transport does; the order across links and the
 * delays are the fuzzer's choice. The relay passes updates at once until `hold`, so the peers
 * can open the room. Updates that a delivery makes (a session's repair) wait like any other.
 */
export class Relay {
  /** The updates that wait on each link, keyed `from:to`. */
  private readonly queues = new Map<string, Held[]>();
  private holding = false;
  private sent = 0;

  constructor(private readonly docs: readonly Y.Doc[]) {
    for (const [from, doc] of docs.entries())
      doc.on('update', (update: Uint8Array, origin: unknown) => {
        if (origin === this) return;
        const sent = (this.sent += 1);
        for (const to of docs.keys()) {
          if (to === from) continue;
          if (this.holding) this.queue(from, to).push({ sent, update });
          else Y.applyUpdate(docs[to]!, update, this);
        }
      });
  }

  private queue(from: number, to: number): Held[] {
    const key = `${from}:${to}`;
    let queue = this.queues.get(key);
    if (!queue) this.queues.set(key, (queue = []));
    return queue;
  }

  /** Hold every later update back until an action delivers it. */
  hold(): void {
    this.holding = true;
  }

  /** The links with updates waiting, in a fixed order. */
  waiting(): { from: number; to: number; count: number }[] {
    return [...this.queues.entries()]
      .filter(([, queue]) => queue.length > 0)
      .map(([key, queue]) => {
        const [from, to] = key.split(':').map(Number) as [number, number];
        return { from, to, count: queue.length };
      })
      .sort((x, y) => x.from - y.from || x.to - y.to);
  }

  /** Deliver up to `count` waiting updates of one link, oldest first. */
  deliver(from: number, to: number, count: number): number {
    const held = this.queue(from, to).splice(0, count);
    for (const { update } of held) Y.applyUpdate(this.docs[to]!, update, this);
    return held.length;
  }

  /** The oldest waiting update of the links to `to`. */
  private oldest(to: number): { from: number; to: number } | undefined {
    let best: { from: number; to: number; sent: number } | undefined;
    for (const link of this.waiting()) {
      if (link.to !== to) continue;
      const sent = this.queue(link.from, link.to)[0]!.sent;
      if (!best || sent < best.sent) best = { ...link, sent };
    }
    return best;
  }

  /** Deliver every update that waits for `to`, in the order they were sent. */
  deliverTo(to: number): number {
    let n = 0;
    for (let link = this.oldest(to); link; link = this.oldest(to))
      n += this.deliver(link.from, link.to, 1);
    return n;
  }

  /** Deliver every waiting update, link by link. Links stay held for the updates after. */
  flush(): void {
    for (let waiting = this.waiting(); waiting.length > 0; waiting = this.waiting())
      for (const { from, to, count } of waiting) this.deliver(from, to, count);
  }

  /** A random delivery of part of one link's waiting updates, or null when none wait. */
  randomDelivery(random: () => number): Delivery | null {
    const waiting = this.waiting();
    if (waiting.length === 0) return null;
    const link = waiting[Math.floor(random() * waiting.length)]!;
    return { kind: 'deliver', ...link, count: 1 + Math.floor(random() * link.count) };
  }
}

/** The client ids the peers take, in an order the seed picks, so the tie rule varies. */
const CLIENT_IDS = [1_001, 2_002, 3_003, 4_004, 5_005];

/** The client ids of `count` peers for `seed`. */
export function clientIdsOf(count: number, seed: number): number[] {
  const random = seededRandom(seed ^ 0x5eed);
  return CLIENT_IDS.slice(0, count)
    .map((id) => [random(), id] as const)
    .sort((x, y) => x[0] - y[0])
    .map(([, id]) => id);
}

/** Editors on Y.Docs with the client ids `ids`, joined by a held relay, loaded from `bytes`. */
export async function openFuzzPeers(bytes: Uint8Array, ids: readonly number[]) {
  const docs = ids.map((id) => {
    const doc = new Y.Doc();
    doc.clientID = id;
    return doc;
  });
  const relay = new Relay(docs);
  const editors: PptxEditorInstance[] = [];
  for (const [i, doc] of docs.entries()) editors.push(await openPeer(doc, `p${i}`, bytes));
  relay.hold();
  return { editors, relay, docs };
}

/**
 * The trail of one run: every action, the problems found, and with tracing on, each action
 * printed with each replica's state.
 */
export class FuzzLog {
  readonly problems: string[] = [];
  readonly trail: string[] = [];

  constructor(
    private readonly trace: boolean,
    private readonly show: () => string
  ) {}

  get failed(): boolean {
    return this.problems.length > 0;
  }

  note(entry: string): void {
    this.trail.push(entry);
    if (this.trace) console.warn(`${entry}\n${this.show()}`);
  }

  fail(message: string): void {
    this.problems.push(
      `${message}\n  last actions (of ${this.trail.length}):\n    ${this.trail.slice(-12).join('\n    ')}`
    );
  }
}

/** After a flush: every session is `ready`, and every replica holds the same package. */
export function roundProblems(editors: readonly PptxEditorInstance[]): string[] {
  const problems: string[] = [];
  for (const [i, session] of sessionsOf(editors).entries())
    if (session.status() !== 'ready') problems.push(`replica ${i} session is ${session.status()}`);
  const first = editors[0]!.session!.store.package;
  for (const [i, e] of editors.entries()) {
    if (i === 0) continue;
    const differ = differingParts(first, e.session!.store.package);
    if (differ.length > 0)
      problems.push(`replicas 0 and ${i} differ after a flush in ${differ.join(', ')}`);
  }
  return problems;
}

/** One replica's saved deck, reopened. */
export interface SavedReplica {
  readonly bytes: Uint8Array;
  readonly pkg: PptxPackage;
}

/**
 * The invariants every fuzzer checks at the end of a run:
 * - every session is `ready`: it applied every remote change;
 * - every replica holds the same package (`canonicalPackageFingerprint`);
 * - each replica's save reopens, the reopened decks are equal, and a second save and reopen
 *   keeps them (`semanticDigest`);
 * - the saved deck passes the part checks (`packageViolations`, `table-merge` included).
 * Returns the problems and the reopened saves.
 */
export function convergenceProblems(editors: readonly PptxEditorInstance[]): {
  problems: string[];
  saved: SavedReplica[];
} {
  const problems: string[] = [];
  for (const [i, session] of sessionsOf(editors).entries()) {
    const snapshot = session.statusSnapshot();
    if (snapshot.status !== 'ready')
      problems.push(
        `replica ${i} session is ${snapshot.status}: ${JSON.stringify(snapshot.reason ?? null)}`
      );
  }
  const first = editors[0]!.session!.store.package;
  for (const [i, e] of editors.entries()) {
    if (i === 0) continue;
    const differ = differingParts(first, e.session!.store.package);
    if (differ.length > 0) problems.push(`replicas 0 and ${i} differ in ${differ.join(', ')}`);
  }
  const saved: SavedReplica[] = [];
  for (const [i, e] of editors.entries()) {
    const bytes = e.save();
    const reopened = bytes ? readPptx(bytes) : undefined;
    if (!bytes || !reopened?.ok) {
      problems.push(`replica ${i} saved a deck that does not reopen`);
      continue;
    }
    saved.push({ bytes, pkg: reopened.pkg });
    const violations = packageViolations(reopened.pkg);
    if (violations.length > 0)
      problems.push(
        `replica ${i} saved a deck that fails the part checks: ${JSON.stringify(violations)}`
      );
    if (semanticDigest(reopened.pkg) !== canonicalPackageFingerprint(reopened.pkg))
      problems.push(`replica ${i}: a second save and reopen changes the saved deck`);
  }
  for (const [i, s] of saved.entries()) {
    if (i === 0) continue;
    const differ = differingParts(saved[0]!.pkg, s.pkg);
    if (differ.length > 0)
      problems.push(`the saved decks of replicas 0 and ${i} differ in ${differ.join(', ')}`);
  }
  return { problems, saved };
}

/** The kind of a problem, without its trail and values, to keep a shrink on one defect. */
export function kindOf(problem: string): string {
  return problem
    .replace(/^seed \d+: /, '')
    .replace(/\n[\s\S]*$/, '')
    .replace(/:.*$/, '')
    .replace(/\d+/g, '#')
    .slice(0, 60);
}

/** The first line of a problem, without its trail: what a known failure pins. */
export const headOf = (problem: string): string => problem.replace(/\n[\s\S]*$/, '');

/** Remove actions while the same kind of problem persists (delta debugging, as DOCX Editor). */
export async function shrink<A>(
  actions: readonly A[],
  kind: string,
  replay: (actions: readonly A[]) => Promise<string[]>
): Promise<A[]> {
  let current = [...actions];
  for (let size = Math.ceil(current.length / 2); size >= 1; size = Math.floor(size / 2)) {
    let removed = true;
    while (removed) {
      removed = false;
      for (let start = 0; start < current.length; start += size) {
        const candidate = [...current.slice(0, start), ...current.slice(start + size)];
        const problems = await replay(candidate);
        if (problems.some((problem) => kindOf(problem) === kind)) {
          current = candidate;
          removed = true;
          break;
        }
      }
    }
  }
  return current;
}

/** A fixed action list that once failed, or that an older test ran, kept as a regression. */
export interface FuzzRegression<A> {
  /** What the history shows, and the issue that fixed it. */
  readonly name: string;
  /** The seed whose client-id order the replay uses. */
  readonly seed: number;
  /** The client ids, when they are not the seed's. */
  readonly clients?: readonly number[];
  readonly actions: readonly A[];
}

/**
 * A history that fails today, with its issue. The test replays it and passes only while it
 * fails with exactly `problem`; it fails on another problem, and it fails once the history
 * passes: then move the entry to the regressions.
 */
export interface FuzzKnownFailure<A> extends FuzzRegression<A> {
  /** The GitHub issue, for example `#1`. */
  readonly issue: string;
  /** The first line of the problem the replay reports today. */
  readonly problem: string;
}

/** A default seed that fails today: the issue its shrunk history shows, and its problem. */
export interface FuzzKnownSeed {
  readonly issue: string;
  /** The first line of the problem the seed reports today. */
  readonly problem: string;
}

/** One fuzzer: how to run a seed and how to replay an action list. */
export interface FuzzSuite<A> {
  /** A short name that starts the test names, for `bun test -t`, for example `table fuzz`. */
  readonly name: string;
  readonly title: string;
  readonly knobs: FuzzKnobs;
  /** The test file, for the commands that a failure prints. */
  readonly file: string;
  /** Generate and perform one seed's actions. */
  readonly run: (seed: number) => Promise<{ problems: string[]; actions: A[] }>;
  /** Perform recorded actions on fresh replicas with the client ids `clients`. */
  readonly replay: (clients: readonly number[], actions: readonly A[]) => Promise<string[]>;
  /** The number of peers, for the client ids of a seed. */
  readonly peers: number;
  readonly regressions?: readonly FuzzRegression<A>[];
  readonly knownFailures?: readonly FuzzKnownFailure<A>[];
  /** Default seeds that fail today. Each still runs, and must fail with its pinned problem. */
  readonly knownSeeds?: Readonly<Record<number, FuzzKnownSeed>>;
  /** Prints what the oracles checked, after the seeds ran. */
  readonly summary?: () => string;
}

/** The commands that trace and shrink one seed. */
function howTo<A>(suite: FuzzSuite<A>, seed: number): string {
  const p = `${suite.knobs.prefix}_FUZZ`;
  // The trace runs the seed's test only; the shrink registers one test.
  const one = `${suite.file} -t '${suite.name}.* seed ${seed}( \\(known|$)'`;
  return [
    `  Trace it:  ${p}_FROM=${seed} ${p}_SEEDS=1 ${p}_TRACE=1 bun test ${one}`,
    `  Shrink it: ${p}_SHRINK=${seed} bun test ${suite.file} -t '${suite.name}'`,
    '  See CONTRIBUTING.md, "Fuzz tests".',
  ].join('\n');
}

/** The known failures whose problem is of the same kind as `problem`, for a failure message. */
function sameKind<A>(suite: FuzzSuite<A>, problem: string): string {
  const kind = kindOf(problem);
  const issues = new Set(
    [...(suite.knownFailures ?? []), ...Object.values(suite.knownSeeds ?? {})]
      .filter((known) => kindOf(known.problem) === kind)
      .map((known) => known.issue)
  );
  return issues.size === 0
    ? '  No known failure has this kind of problem: shrink it and file an issue.'
    : `  Known failures of the same kind: ${[...issues].join(', ')}. Shrink it and compare ` +
        'the actions before you file an issue.';
}

/** Register the seeds, the regressions, and the known failures of one fuzzer. */
export function fuzzSuite<A>(suite: FuzzSuite<A>): void {
  const { knobs, file } = suite;
  const replaySeed = (seed: number, actions: readonly A[], clients?: readonly number[]) =>
    suite.replay(clients ?? clientIdsOf(suite.peers, seed), actions);
  describe(`${suite.name}: ${suite.title}`, () => {
    if (knobs.shrink !== undefined) {
      const seed = knobs.shrink;
      test(`shrink seed ${seed}`, async () => {
        const { problems, actions } = await suite.run(seed);
        expect(problems, `seed ${seed} does not fail, so there is nothing to shrink.`).not.toEqual(
          []
        );
        const kind = kindOf(problems[0]!);
        const replay = (candidate: readonly A[]) => replaySeed(seed, candidate);
        const minimal = await shrink(actions, kind, replay);
        const problem = (await replay(minimal))[0] ?? kind;
        const entry = [
          '  {',
          `    name: '<what it shows>',`,
          `    issue: '#<issue>',`,
          `    seed: ${seed},`,
          `    problem: ${JSON.stringify(headOf(problem))},`,
          '    actions: [',
          ...minimal.map((action) => `      ${JSON.stringify(action)},`),
          '    ],',
          '  },',
        ].join('\n');
        console.warn(
          `seed ${seed} shrinks from ${actions.length} to ${minimal.length} actions.\n` +
            `Problem: ${problem}\n${sameKind(suite, problem)}\n` +
            `File an issue with these actions, and add this entry to the known failures of ` +
            `${file}. When the issue is fixed, move it to the regressions (drop \`issue\` and ` +
            `\`problem\`):\n${entry}\n` +
            `If seed ${seed} is a default seed, also add it to the known seeds of ${file}, ` +
            `with its own problem:\n  ${seed}: { issue: '#<issue>', problem: ` +
            `${JSON.stringify(headOf(problems[0]!))} },`
        );
      }, 600_000);
      return;
    }
    const failed: { seed: number; kind: string }[] = [];
    afterAll(() => {
      if (suite.summary) console.warn(suite.summary());
      if (failed.length === 0) return;
      const byKind = new Map<string, number[]>();
      for (const { seed, kind } of failed) byKind.set(kind, [...(byKind.get(kind) ?? []), seed]);
      console.warn(
        `${suite.name}: ${failed.length} of ${knobs.seeds} seeds failed:\n` +
          [...byKind].map(([kind, seeds]) => `  ${kind}: seeds ${seeds.join(', ')}`).join('\n')
      );
    });
    // Determinism: a seed gives the same actions and problems twice, and a replay of its
    // actions gives the same problems, so a failing seed reproduces anywhere. A trace skips it,
    // so it prints each action once.
    const determinism = knobs.trace ? test.skip : test;
    determinism(
      `seed ${knobs.from} runs and replays the same way twice`,
      async () => {
        const first = await suite.run(knobs.from);
        const second = await suite.run(knobs.from);
        expect(second.actions).toEqual(first.actions);
        expect(second.problems.map(headOf)).toEqual(first.problems.map(headOf));
        const replayed = await replaySeed(knobs.from, first.actions);
        expect(replayed.map(headOf)).toEqual(first.problems.map(headOf));
      },
      60_000
    );
    for (let seed = knobs.from; seed < knobs.from + knobs.seeds; seed += 1) {
      const known = suite.knownSeeds?.[seed];
      const label = known ? `seed ${seed} (known failure ${known.issue})` : `seed ${seed}`;
      test(
        label,
        async () => {
          const { problems, actions } = await suite.run(seed);
          const head = problems[0] && headOf(problems[0]);
          if (known && head === known.problem) return;
          if (known && !head)
            throw new Error(
              `seed ${seed} passes now, so ${known.issue} may be fixed. Remove the seed from ` +
                `the known seeds of ${file}.`
            );
          if (problems.length === 0) return;
          failed.push({ seed, kind: kindOf(problems[0]!) });
          const pinned = known
            ? `It pins ${known.issue}, which fails with ${JSON.stringify(known.problem)}.\n`
            : '';
          throw new Error(
            `seed ${seed} failed after ${actions.length} actions:\n${problems.join('\n')}\n` +
              `${pinned}${sameKind(suite, problems[0]!)}\n${howTo(suite, seed)}`
          );
        },
        60_000
      );
    }
    for (const regression of suite.regressions ?? [])
      test(`regression: ${regression.name}`, async () => {
        const problems = await replaySeed(regression.seed, regression.actions, regression.clients);
        if (problems.length === 0) return;
        throw new Error(
          `${regression.name}: the fixed history of ${regression.actions.length} actions ` +
            `fails:\n${problems.join('\n')}\n${sameKind(suite, problems[0]!)}\n` +
            `  Trace it: ${knobs.prefix}_FUZZ_TRACE=1 bun test ${file} -t '${regression.name}'`
        );
      });
    for (const known of suite.knownFailures ?? [])
      test(`known failure ${known.issue}: ${known.name}`, async () => {
        const problems = await replaySeed(known.seed, known.actions, known.clients);
        const head = problems[0] && headOf(problems[0]);
        if (head === known.problem) return;
        const trace =
          `\n  Trace it: ${knobs.prefix}_FUZZ_TRACE=1 bun test ${file} ` +
          `-t 'known failure ${known.issue}: ${known.name}'`;
        throw new Error(
          head
            ? `${known.name} (${known.issue}) fails another way now:\n${problems.join('\n')}\n` +
                `It pins ${JSON.stringify(known.problem)}. File an issue for the new problem, ` +
                'or update the pin when the change explains it.' +
                trace
            : `${known.name} passes now, so ${known.issue} may be fixed. Move the entry to ` +
                `the regressions of ${file} (drop \`issue\` and \`problem\`).` +
                trace
        );
      });
  });
}
