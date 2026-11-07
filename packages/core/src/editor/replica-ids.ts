// Ids that two replicas of one collaboration room can mint at the same time: the PPTX-only part,
// beside DOCX Editor's `actor-scoped-ids.ts` (copied with the omissions its header lists).
//
// Each replica picks "one past the highest" from its own state, and the state-based merge
// then joins both parts, so two peers that add a slide, a picture, a link, or a shape at once
// would save one relationship id, part name, slide id, or shape id twice. While a session is
// attached, `package-edits.ts` `stripedIds` mints in the actor's residue class instead. This
// module binds a document store to its actor and keeps the session's record of minted numbers.
// The actor is the attached collaboration session's id, one per replica (DOCX Editor uses the
// user's actor id), so two tabs of one user also keep apart.
//
// Limits: two actors that hash into one class can still collide (a room of five about 0.02 %,
// twenty about 0.3 %), and nothing detects it after the fact. Shape names are numbered from
// their ids, so a shape added in a session reads "Picture 40212". Still open: entries keyed by a
// value that two peers can each add first (one notes slide per slide, a layout placeholder such
// as the slide number) need a keyed singleton merge.

import { attr, elements, type Element, type PptxPackage } from '../store/index.ts';
import {
  ACTOR_ID_STRIPE,
  actorStripe,
  relationshipNumberOf,
} from '../automation/commands/actor-scoped-ids.ts';
import { warnEmptySessionId } from './dev-warnings.ts';
import type { EditorCollaborationSession } from './modules.ts';
import type { MintedIdKind, MintedIds } from '../automation/commands/op-batch.ts';
import { unsignedIntId } from '../automation/commands/package-edits.ts';

/**
 * Numbers minted in each collaboration session's stripe. The session's id is the actor, so the
 * record belongs to the session and outlives the editors attached to it.
 */
const MINTED_BY_SESSION = new WeakMap<EditorCollaborationSession, MintedIds>();

/** The record of numbers minted in one collaboration session's stripe. */
function mintedIdsOf(session: EditorCollaborationSession): MintedIds {
  let minted = MINTED_BY_SESSION.get(session);
  if (!minted) MINTED_BY_SESSION.set(session, (minted = new Map()));
  return minted;
}

/**
 * Binds a document store to the attached collaboration session, or unbinds it with
 * `undefined`. With a collaboration session, new ids are minted in this replica's stripe. The
 * record of minted numbers belongs to the collaboration session, as the actor does, so a reload
 * and a new editor on the same session (a remount) never mint a number twice. The record also
 * takes every number of the stripe that the store's package holds (`seedMintedIds`). An empty
 * session id binds nothing and prints a development warning: it is not an actor.
 * `attachEditorModules` binds before the attach (a listener's command then mints in the
 * stripe) and again after it (the seed then takes the room's parts).
 */
export function bindReplicaIds(
  store: { readonly package: PptxPackage },
  collaboration: EditorCollaborationSession | undefined
): void {
  if (!collaboration || collaboration.sessionId === '') {
    if (collaboration) warnEmptySessionId();
    BINDINGS.delete(store);
    return;
  }
  const minted = mintedIdsOf(collaboration);
  BINDINGS.set(store, { allocationActor: collaboration.sessionId, minted });
  seedMintedIds(minted, store.package, collaboration.sessionId);
}

/** The binding of one document store: its allocation actor and the record of minted numbers. */
export interface ReplicaBinding {
  readonly allocationActor: string;
  readonly minted: MintedIds;
}

/**
 * Bindings by document store. A WeakMap keeps them off the store's public surface, and
 * `OpBatch` keeps the record private, so only `OpBatch.mintedOf` reads it.
 */
const BINDINGS = new WeakMap<object, ReplicaBinding>();

/** The binding of a document store, or undefined with no collaboration session. */
export function replicaBindingOf(store: object): ReplicaBinding | undefined {
  return BINDINGS.get(store);
}

/** The deepest element the seed scan visits, as the shape-id walk does. */
const SEED_MAX_DEPTH = 64;

/** The trailing number of a part name (`/ppt/slides/slide12.xml` → 12). */
const PART_NUMBER = /(\d{1,10})\.[^./]*$/;

/**
 * Adds to `minted` every number in the actor's stripe that `pkg` holds: shape ids
 * (`p:cNvPr/@id`), slide ids (`p:sldId/@id`), relationship ids (`rIdN`), part-name numbers,
 * slide creation ids (`p14:creationId/@val`), time node ids (`p:cTn/@id`), and build group ids
 * (`p:cTn/@grpId`, `p:bldP/@grpId`), each read as
 * `xsd:unsignedInt` (`unsignedIntId`). Comment indexes need no seed: a new index starts after the author's
 * `p:cmAuthor/@lastIdx`, which never goes down when a comment is deleted.
 *
 * A number from the file that falls in this replica's stripe is then never minted, also after
 * a peer deleted the object that held it: that peer's undo brings the object back, and a
 * second holder would duplicate the id. One pass over the package; numbers only go into sets,
 * never into an allocation size.
 */
export function seedMintedIds(minted: MintedIds, pkg: PptxPackage, actor: string): void {
  const stripe = actorStripe(actor);
  const add = (kind: MintedIdKind, value: number | null | undefined): void => {
    if (value === null || value === undefined || value % ACTOR_ID_STRIPE !== stripe) return;
    let set = minted.get(kind);
    if (!set) minted.set(kind, (set = new Set()));
    set.add(value);
  };
  for (const [name, part] of pkg.parts) {
    const digits = PART_NUMBER.exec(name);
    if (digits) add('part', Number(digits[1]));
    if (part.kind !== 'xml') continue;
    const rels = name.endsWith('.rels');
    const walk = (el: Element, depth: number): void => {
      if (depth > SEED_MAX_DEPTH) return;
      switch (el.local) {
        case 'cNvPr':
          add('shapeId', unsignedIntId(attr(el, 'id')));
          break;
        case 'sldId':
          add('slideId', unsignedIntId(attr(el, 'id')));
          break;
        case 'cTn':
          add('timeNodeId', unsignedIntId(attr(el, 'id')));
          add('buildGroupId', unsignedIntId(attr(el, 'grpId')));
          break;
        case 'bldP':
          add('buildGroupId', unsignedIntId(attr(el, 'grpId')));
          break;
        case 'creationId':
          add('creationId', unsignedIntId(attr(el, 'val')));
          break;
        case 'Relationship':
          if (rels) add('rId', relationshipNumberOf(attr(el, 'Id') ?? ''));
          break;
      }
      for (const c of elements(el)) walk(c, depth + 1);
    };
    walk(part.xml.root, 0);
  }
}
