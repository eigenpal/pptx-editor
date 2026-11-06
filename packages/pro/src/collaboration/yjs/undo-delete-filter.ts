/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Which items an undo may delete, after DOCX Editor's
// `packages/pro/src/collaboration/document/undo-delete-filter.ts` (commit 5993bd22). The rule is
// the same; the record fields are this schema's (`schema.ts`).
//
// Creating a node is `nodes.set(key, record)`, so an undo of that creation would delete the
// record. A peer that wrote into the node (an attribute, a child, a character) wrote inside the
// record's own shared types, and deleting the record deletes that write with it. The write can
// also arrive after the undo, and then lands in a deleted type: Yjs drops it. Nothing is left to
// repair, so the undo keeps the record: its entry in `nodes`, the fields that hold a shared type
// (`a`, `c`, `x`), and the first values of `p`, `k`, and `s`, which say what the record is, and
// of the copy marks `ci` and `cp`, which say what its text copies.
//
// Everything else undoes as before: the key that lists the node under its parent, the attributes
// and children this author wrote, its characters, and the lineage fields. The node leaves the
// document, and what a peer wrote into it stays in a record no list holds. `undo-unlisted.ts`
// shows that record again where the undo took it from.

import * as Y from 'yjs';
import { F } from './schema.ts';
import { IDENTITY } from './text-identity.ts';
import { holderItem, isFirstWrite } from './yjs-items.ts';

/**
 * Record fields whose first value names the record: its part, its key, and its shell, and the
 * copy mark of a text a split made (`text-identity.ts`), so a copy of that copy still leads to
 * the character first typed after the undo.
 */
const IDENTITY_FIELDS: ReadonlySet<string> = new Set([
  F.part,
  F.key,
  F.shell,
  IDENTITY.copyOf,
  IDENTITY.newParagraph,
]);

/** For `Y.UndoManager`'s `deleteFilter`: false keeps the item that the undo would delete. */
export function nodeRecordDeleteFilter(nodes: Y.Map<unknown>): (item: Y.Item) => boolean {
  return (item) => {
    const parent: unknown = item.parent;
    // The record entry itself.
    if (parent === nodes) return false;
    // A field of a record: its holder sits directly in `nodes`.
    const holder = holderItem(parent);
    if (holder === null || holder.parent !== (nodes as unknown)) return true;
    if (item.content instanceof Y.ContentType) return false;
    return !(item.parentSub !== null && IDENTITY_FIELDS.has(item.parentSub) && isFirstWrite(item));
  };
}
