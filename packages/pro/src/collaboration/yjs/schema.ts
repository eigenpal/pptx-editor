/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The shared Yjs schema of a PPTX room, after DOCX Editor's `collaboration/document/schema.ts`.
//
//   pptx-meta-v1      Y.Map   schema version
//   pptx-parts-v1     Y.Map   part name → { kind: 'xml', root: key } | { kind: 'binary' }
//   pptx-binaries-v1  Y.Map   part name → Uint8Array
//   pptx-nodes-v1     Y.Map   key → node record (Y.Map)
//   pptx-splits-v1    Y.Map   split index, outside the undo scope (`text-lineage.ts`)
//   pptx-moves-v1     Y.Map   moved key → part name (`text-lineage.ts`)
//   pptx-pieces-v1    Y.Map   split piece key → part name (`text-lineage.ts`)
//   pptx-copies-v1    Y.Map   copy text key → part name (`text-identity.ts`)
//   pptx-undone-v1    Y.Map   key an undo unlisted → where it was listed (`undo-unlisted.ts`)
//   pptx-restored-v1  Y.Map   character an undo or redo put back → the one it stands for
//                             (`text-identity.ts`)
//
// A node record holds `p` (its part name), `k` (its own key, checked on read), and either an
// element (`s` shell: name and namespace declarations, `a` Y.Map of attributes, `c` Y.Array of
// child keys) or a text node (`x` Y.Text). Split and move lineage adds the optional fields of
// `LINEAGE` in `text-lineage.ts` (`sf`, `sb`, `sa`, `sl`, `sg`, `se`, `sk`, `mf`, `mi`, `mq`,
// `mt`, `mo`, `rw`), character identity the fields of `IDENTITY` in `text-identity.ts` (`ci`, `cp`, `rr.<key>`), and
// merged cells the fields of `MERGED` in `merged-cells.ts` (`mh`, `cv.<client>`).
// Keys are logical ids shared by every peer; store node ids stay replica-local. Everything read
// from the room is untrusted: it reaches the store only through `buildReplicaPart`.

import * as Y from 'yjs';

/**
 * The room format. A replica attaches only to a room of its own format, so a replica of another
 * format refuses the room and keeps its own deck (`schema-version-mismatch`).
 *
 * - 7: comment lists and author lists are XML trees in the room, and a save joins the lists that
 *   two peers added at once and points copied threads at their slide (`comment-repair.ts` in
 *   `@pptx-editor.dev/core`). A replica of format 6 saves the room without that step, so its file
 *   could hold two author lists or threads that name another slide.
 * - 6: a read puts back a paragraph that a write removed from any text body while a peer typed
 *   into it, with that text (`withRemovedRuns` in `removed-content.ts`), counts a move mark only
 *   while another record lists the node, and shows a run that the move rescue put back without
 *   any write listing it there. A replica of format 5 shows no such paragraph, so it would
 *   show other text.
 * - 5: an undo records where each node it unlists was (`pptx-undone-v1`) and marks each
 *   character it puts back with the character it stands for (`pptx-restored-v1`), and every
 *   replica repairs by them. A replica of format 4 neither writes nor reads them: it would
 *   show text that the others drop, so the replicas would differ.
 * - 4: a cell merge leaves no hidden copy of the text it moves. A read shows moved text in the
 *   drawn cell over it and hides the empty paragraph a merge left when moved text comes back
 *   (`merged-cells.ts`), puts a node moved twice back into an earlier parent when the last
 *   one is gone (`mo`), and puts back an empty paragraph that a peer typed into while another
 *   removed it (`text-identity.ts`). A replica of format 3 would show other text, so it must
 *   not share a room with format 4.
 * - 3: a table structure command adds a structure mark (`a:ext` under the frame's `p:cNvPr`,
 *   `table-marks.ts` in `@pptx-editor.dev/core`). A replica of format 2 does not know the mark
 *   and would save it into its file, so it must not share a room with format 3.
 * - 2: the seed marks its keys, and a read shows schema singletons with their companions and
 *   sorts children into schema order, and records split and move lineage.
 * - 1: resolved the same conflicts differently.
 *
 * No format before 3 was published.
 */
export const SCHEMA_VERSION = 7;

export const META_KEY = 'pptx-meta-v1';
export const PARTS_KEY = 'pptx-parts-v1';
export const BINARIES_KEY = 'pptx-binaries-v1';
export const NODES_KEY = 'pptx-nodes-v1';

/** Node record fields. */
export const F = Object.freeze({
  part: 'p',
  key: 'k',
  shell: 's',
  attrs: 'a',
  children: 'c',
  text: 'x',
});

/** Separates namespace URI and local name in an attribute key. XML names cannot hold it. */
export const ATTR_SEP = '\u001f';

/** Minted key numbers above this are ignored when a mirror counts on: a peer may write any key. */
export const MAX_MINT_COUNTER = 2 ** 40;

/**
 * The transaction origin of this replica's own edits; remote handling skips it, and the
 * session's undo manager tracks only it.
 */
export const LOCAL_ORIGIN = Object.freeze({ kind: 'pptx-local' });

/**
 * The origin of a write after an undo that gives back what a peer still uses
 * (`undo-repair.ts`). Every replica, this one too, rebuilds from it; it is never undoable.
 */
export const REPAIR_ORIGIN = Object.freeze({ kind: 'pptx-undo-repair' });

/**
 * The origin of a redo's repair: the step's own values written again after an undo repair. The
 * undo manager tracks it into the redo's own step, so one undo reverts both.
 */
export const REDO_REPAIR_ORIGIN = Object.freeze({ kind: 'pptx-redo-repair' });

/**
 * The origin of a write that repairs text a peer's split moved while this replica edited it
 * (`text-repair.ts`). Every replica, this one too, rebuilds from it; it is never undoable.
 */
export const TEXT_REPAIR_ORIGIN = Object.freeze({ kind: 'pptx-text-repair' });

/**
 * The origin of the delete of listings that no read shows, written before a local edit
 * (`PackageMirror.dropHiddenListings`). No read changes, so remote handling skips it; it is
 * never undoable, so no undo lists a node twice again.
 */
export const TIDY_ORIGIN = Object.freeze({ kind: 'pptx-tidy' });

/**
 * The origin of the record an undo writes before it runs, of where each node it unlists was
 * (`undo-unlisted.ts`). No history tracks it, and it changes no part by itself.
 */
export const UNDONE_ORIGIN = Object.freeze({ kind: 'pptx-undone-listing' });

/** The origin of the seed write: skipped like a local edit, but never undoable. */
export const SEED_ORIGIN = Object.freeze({ kind: 'pptx-seed' });

/** Limits on room content, checked before anything is built. */
export const ROOM_LIMITS = Object.freeze({
  /** One binary part. The zip reader caps whole packages; a room holds parts one by one. */
  maxBinaryBytes: 256 * 1024 * 1024,
  /** Nesting while reading node records, above the XML reader's 256 so the check is the core's. */
  maxReadDepth: 300,
  /**
   * Concurrent copies of one schema singleton under one parent. Honest peers leave one
   * copy for each peer that wrote it at once; a part over the cap is refused.
   */
  maxCopiesPerGroup: 256,
});

export type PartEntry =
  { readonly kind: 'xml'; readonly root: string } | { readonly kind: 'binary' };

export interface Room {
  readonly doc: Y.Doc;
  readonly meta: Y.Map<unknown>;
  readonly parts: Y.Map<unknown>;
  readonly binaries: Y.Map<unknown>;
  readonly nodes: Y.Map<unknown>;
}

export function roomOf(doc: Y.Doc): Room {
  return {
    doc,
    meta: doc.getMap(META_KEY),
    parts: doc.getMap(PARTS_KEY),
    binaries: doc.getMap(BINARIES_KEY),
    nodes: doc.getMap(NODES_KEY),
  };
}

/** A part entry read from the room, or null when it is not one this schema writes. */
export function partEntryOf(value: unknown): PartEntry | null {
  if (typeof value !== 'object' || value === null) return null;
  const v = value as Record<string, unknown>;
  if (v.kind === 'binary') return { kind: 'binary' };
  if (v.kind === 'xml' && typeof v.root === 'string') return { kind: 'xml', root: v.root };
  return null;
}
