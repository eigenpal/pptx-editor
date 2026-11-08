/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/

// The two-peer convergence harness of the collaboration coverage gate. Two editors on two
// Y.Docs, joined by a link that can pause (the same link as `yjs-collaboration.test.ts`). A
// fixture prepares peer A, the link pauses, peer A runs one command or one text input, and the
// link flushes. Both replicas must then hold the same package, and save and reopen must keep it.

import { expect } from 'bun:test';
import * as Y from 'yjs';
import { PNG } from 'pngjs';
import {
  createPptxEditor,
  type CommandResult,
  type EditorCommand,
  type EditorCommandType,
  type PptxEditorInstance,
  type TextInputKind,
} from '@pptx-editor.dev/core/editor';
import {
  canonicalFingerprint,
  canonicalPackageFingerprint,
  readPptx,
  semanticDigest,
  type PackagePart,
  type PptxPackage,
} from '@pptx-editor.dev/core/store';
import {
  buildPptx,
  cell,
  para,
  run,
  shape,
  solid,
  table,
} from '../../../../scripts/fixtures/builder.ts';
import { clickEntrances } from '../../../../scripts/fixtures/timing.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

/** A small gray PNG, for the picture in the deck and for the image commands. */
export function png(width = 8, height = 6, gray = 180): Uint8Array {
  const img = new PNG({ width, height });
  img.data.fill(gray);
  return new Uint8Array(PNG.sync.write(img));
}

/**
 * The coverage deck. Slide 1: text shapes `A`, `B`, and `C`, a picture, and a 2 × 3 table, with
 * notes, and a click entrance on `A`, `C`, and the table, so the commands that remove a shape
 * also remove its `p:timing` entries. Slide 2: the title placeholder `D`, with its own position and fill. Slide 3: empty. The master and layouts carry the slide-number
 * placeholder.
 */
export function coverageDeck(): Uint8Array {
  const text = (id: number, name: string, x: number, y: number, words: string, ph?: string) =>
    shape({
      id,
      name,
      x,
      y,
      cx: 3 * IN,
      cy: IN,
      fill: solid('4472C4'),
      text: para(run(words)),
      ...(ph ? { ph } : {}),
    });
  const pic =
    '<p:pic><p:nvPicPr><p:cNvPr id="5" name="Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr><a:xfrm><a:off x="${8 * IN}" y="${IN / 2}"/><a:ext cx="${2 * IN}" cy="${1.5 * IN}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
  const grid = table({
    id: 6,
    x: IN,
    y: 4.5 * IN,
    cx: 3 * IN,
    cy: 2 * 370_840,
    columns: [IN, IN, IN],
    tblPr: '<a:tblPr firstRow="1" bandRow="1"/>',
    rows: ['abc', 'def']
      .map((r) => `<a:tr h="370840">${[...r].map((c) => cell(c)).join('')}</a:tr>`)
      .join(''),
  });
  return buildPptx({
    footers: true,
    datesAndFooters: true,
    slides: [
      {
        layout: 'blank',
        rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
        shapes:
          text(2, 'A', IN, IN / 2, 'Hello world') +
          text(3, 'B', IN, 2 * IN, 'Second shape') +
          text(4, 'C', 5 * IN, 2 * IN, 'Third shape') +
          pic +
          grid,
        notes: 'Speaker notes here',
        ext: clickEntrances([2, 4, 6]),
      },
      // `D` is the title placeholder, moved off its layout position, so Reset has work to do.
      { layout: 'title', shapes: text(2, 'D', IN, IN, 'Slide two', '<p:ph type="title"/>') },
      { layout: 'blank', shapes: '' },
    ],
    media: [{ name: 'image1.png', bytes: png(), contentType: 'image/png' }],
  });
}

/** Two docs and a sync link that can pause, to make an edit wait for the flush. */
export function link(a: Y.Doc, b: Y.Doc) {
  const queue: [Y.Doc, Uint8Array][] = [];
  let paused = false;
  const send = (to: Y.Doc) => (update: Uint8Array, origin: unknown) => {
    if (origin === 'link') return;
    if (paused) queue.push([to, update]);
    else Y.applyUpdate(to, update, 'link');
  };
  a.on('update', send(b));
  b.on('update', send(a));
  return {
    pause: () => (paused = true),
    flush: () => {
      paused = false;
      for (const [to, update] of queue.splice(0)) Y.applyUpdate(to, update, 'link');
    },
  };
}

/** One peer: an editor in a collaboration room on `doc`, with the deck `bytes` open. */
export async function replica(doc: Y.Doc, actorId: string, bytes: Uint8Array) {
  const session = createYjsCollaborationSession({
    doc,
    documentId: 'deck',
    identity: { actorId, name: actorId },
  });
  const editor = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    modules: [collaborationModule({ session })],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

export const fingerprint = (e: PptxEditorInstance): string =>
  canonicalPackageFingerprint(e.session!.store.package);

/** One part's fingerprint, or undefined when the package has no such part. */
function partFingerprint(part: PackagePart | undefined): string | undefined {
  if (part === undefined) return undefined;
  if (part.kind === 'xml') return canonicalFingerprint(part.xml);
  return `binary ${Buffer.from(part.bytes).toString('base64')}`;
}

/** The names of the parts that differ between two packages, for the failure messages. */
export function differingParts(x: PptxPackage, y: PptxPackage): string[] {
  const names = new Set([...x.parts.keys(), ...y.parts.keys()]);
  return [...names]
    .filter((n) => partFingerprint(x.parts.get(n)) !== partFingerprint(y.parts.get(n)))
    .sort();
}

const HOW_TO = 'See CONTRIBUTING.md, "Add a collaboration fixture".';

/** What a fixture prepares: the deck both peers open, and the state of peer A. */
export interface FixtureSetup {
  /** The deck both peers open; the coverage deck when absent. */
  readonly deck?: () => Uint8Array;
  /**
   * Prepare peer A: select, or make the edits the command needs. Edits made here reach peer B
   * before the command runs.
   */
  readonly arrange?: (a: PptxEditorInstance) => void;
}

/** A two-peer convergence fixture for one command. */
export interface ConvergenceFixture extends FixtureSetup {
  /** The command, built against peer A after `arrange`. */
  readonly command: (a: PptxEditorInstance) => EditorCommand;
}

/** A two-peer convergence fixture for one kind of text input (`PptxEditorInstance.input`). */
export interface TextInputFixture extends FixtureSetup {
  /** Place the caret or the text selection on peer A. */
  readonly arrange: (a: PptxEditorInstance) => void;
  /** The `data` argument of `input`: the typed text for `text`. */
  readonly data?: string;
}

/** Open both peers on the fixture's deck and arrange peer A. */
async function prepare(label: string, fixture: FixtureSetup) {
  const bytes = (fixture.deck ?? coverageDeck)();
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sync = link(docA, docB);
  const a = await replica(docA, 'a', bytes);
  const b = await replica(docB, 'b', bytes);
  try {
    fixture.arrange?.(a);
    expect(
      differingParts(a.session!.store.package, b.session!.store.package),
      `${label}: the edits of \`arrange\` did not reach peer B. A command that \`arrange\` ` +
        'runs must converge too; fix its own fixture first.'
    ).toEqual([]);
  } catch (error) {
    a.destroy();
    b.destroy();
    throw error;
  }
  return { a, b, sync };
}

/**
 * Run one edit on peer A with the link paused, then flush. Asserts that the edit applies and
 * changes peer A's deck, that peer B does not see it before the flush, that both replicas then
 * hold the same package, and that both saves reopen to one package that a second save and
 * reopen keeps. A failure names the parts that differ.
 */
export async function convergeEdit(
  label: string,
  fixture: FixtureSetup,
  edit: (a: PptxEditorInstance) => CommandResult
): Promise<void> {
  const { a, b, sync } = await prepare(label, fixture);
  try {
    const before = a.session!.store.package;
    sync.pause();
    expect(
      edit(a),
      `${label} was refused on peer A. Change the fixture's \`arrange\` (selection, slide, or ` +
        `state) so that the edit applies. ${HOW_TO}`
    ).toEqual({ ok: true });
    expect(
      differingParts(before, a.session!.store.package),
      `${label} did not change peer A's deck: it is a no-op on this fixture, so the fixture ` +
        `proves nothing. Arrange a state where it writes. ${HOW_TO}`
    ).not.toEqual([]);
    expect(
      differingParts(before, b.session!.store.package),
      `${label}: peer B changed before the flush, so the link did not pause.`
    ).toEqual([]);
    sync.flush();

    expect(
      differingParts(a.session!.store.package, b.session!.store.package),
      `${label}: the replicas differ in these parts after the flush. The collaboration mirror ` +
        '(packages/pro/src/collaboration/yjs/) does not carry this write, or a replica rebuilds ' +
        'it differently.'
    ).toEqual([]);
    expect(semanticDigest(b.session!.store.package)).toBe(semanticDigest(a.session!.store.package));
    // The editor's save, which drops editor-only marks (structure marks), reopens to one package on
    // both peers, and a second save and reopen keeps it.
    const saved = [a, b].map((e) => {
      const reopened = readPptx(e.save()!);
      if (!reopened.ok)
        throw new Error(`${label}: a saved deck does not reopen: ${reopened.error.reason}`);
      return reopened.pkg;
    });
    expect(
      differingParts(saved[0]!, saved[1]!),
      `${label}: the saved decks of the two peers differ in these parts.`
    ).toEqual([]);
    expect(
      semanticDigest(saved[0]!),
      `${label}: a second save and reopen changes the saved deck.`
    ).toBe(canonicalPackageFingerprint(saved[0]!));
  } finally {
    a.destroy();
    b.destroy();
  }
}

/** Run a command fixture: the command on peer A with the link paused, then flush. */
export function converge(type: EditorCommandType, fixture: ConvergenceFixture): Promise<void> {
  return convergeEdit(`Command \`${type}\``, fixture, (a) => {
    const command = fixture.command(a);
    expect(command.type, `The fixture for \`${type}\` builds a \`${command.type}\` command.`).toBe(
      type
    );
    return a.exec(command);
  });
}

/** Run a text input fixture: `input(kind, data)` on peer A with the link paused, then flush. */
export function convergeInput(kind: TextInputKind, fixture: TextInputFixture): Promise<void> {
  return convergeEdit(`Text input \`${kind}\``, fixture, (a) => a.input(kind, fixture.data));
}

/**
 * Prove the reason "writes nothing" of an uncovered command: on peer A it runs, and neither
 * replica's package changes.
 */
export async function writesNothing(
  type: EditorCommandType,
  fixture: ConvergenceFixture
): Promise<void> {
  const label = `Command \`${type}\``;
  const { a, b } = await prepare(label, fixture);
  try {
    const before = a.session!.store.package;
    const command = fixture.command(a);
    expect(command.type, `The probe for \`${type}\` builds a \`${command.type}\` command.`).toBe(
      type
    );
    expect(a.exec(command), `${label} was refused on peer A; fix the probe's \`arrange\`.`).toEqual(
      { ok: true }
    );
    expect(
      differingParts(before, a.session!.store.package),
      `${label} is in COLLABORATION_UNCOVERED because it writes nothing, but it changed these ` +
        `parts. Remove its entry and add a convergence fixture. ${HOW_TO}`
    ).toEqual([]);
    expect(differingParts(before, b.session!.store.package)).toEqual([]);
  } finally {
    a.destroy();
    b.destroy();
  }
}
