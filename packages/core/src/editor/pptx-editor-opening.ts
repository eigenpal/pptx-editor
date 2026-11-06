// The editor's side of a progressive open (`progressive-open.ts`), after DOCX Editor's
// `docx-editor-opening-layout.ts`: the open ends when the first slide shows, and the rest of
// the deck parses in background tasks that belong to the session they started for. They stop
// once the editor holds another session. A reader of the whole deck (save) finishes them first.
//
// When the background parse settles, the fonts follow: a slide parsed in the background can
// name a family the parts parsed at the open did not, so the packaged substitutes and the host
// origins are asked for the families still missing.

import type { PackageRejection } from '../store/index.ts';
import { deckFontFamilies } from './deck-font-request.ts';
import type { DocumentSession } from './document-session.ts';
import { continueOpening } from './progressive-open.ts';

type XmlRejection = Extract<PackageRejection, { reason: 'xml' }>;

/** What the opening reads from and asks of the editor. */
export interface OpeningEditor {
  session(): DocumentSession | null;
  /** Publish a snapshot: the pending slide count moved. */
  emit(): void;
  /** Refuse the deck of `owner`: a deferred part did not parse. */
  refuse(owner: DocumentSession, rejection: XmlRejection): void;
  /** `owner`'s faces changed: register them for paint and publish. */
  facesChanged(owner: DocumentSession): Promise<void>;
  /** Ask the host's font origins for families no face covers. */
  requestFamilies(families: readonly string[]): Promise<void>;
}

/** The opening of the sessions one editor holds. */
export function createOpeningLane(editor: OpeningEditor) {
  /** The families each session's open asked the host origins for, folded. */
  const asked = new WeakMap<DocumentSession, ReadonlySet<string>>();
  const settle = (owner: DocumentSession): void => {
    owner.opening = null;
    void owner
      .addMissingSubstitutes()
      .then((added) => {
        if (added && editor.session() === owner) return editor.facesChanged(owner);
      })
      // A face that does not load is in the session's font report; nothing else to do.
      .catch(() => undefined);
    const before = asked.get(owner) ?? new Set<string>();
    const more = deckFontFamilies(owner.store.package).filter(
      (family) => !before.has(family.toLowerCase())
    );
    if (more.length > 0) void editor.requestFamilies(more);
  };
  const refused = new WeakSet<DocumentSession>();
  /** Refuse `owner`'s deck once, now: the caller is not inside a read of the deck. */
  const refuseNow = (owner: DocumentSession, rejection: XmlRejection): void => {
    if (refused.has(owner)) return;
    refused.add(owner);
    editor.refuse(owner, rejection);
  };
  /** Refuse `owner`'s deck once, after the read that found the damage has returned. */
  const refuse = (owner: DocumentSession, rejection: XmlRejection): void =>
    queueMicrotask(() => refuseNow(owner, rejection));
  return {
    /** Parse the rest of `owner`'s deck in background tasks, if any is left. */
    start(owner: DocumentSession): void {
      // A slide that does not parse, wherever it is read, refuses the deck at once.
      owner.onDamage = (rejection) => refuse(owner, rejection);
      if (owner.damage) refuse(owner, owner.damage);
      const open = owner.opening;
      if (!open) return;
      // What the open asked for: the parts parsed so far, as `hostFonts.open` read them.
      const families = deckFontFamilies(owner.store.package).map((f) => f.toLowerCase());
      asked.set(owner, new Set(families));
      continueOpening(open, {
        current: () => editor.session() === owner && owner.opening === open,
        publish: () => editor.emit(),
        settled: () => (settle(owner), editor.emit()),
        failed: (rejection) => refuseNow(owner, rejection),
      });
    },
    /**
     * Finish the open now, for a reader of the whole deck. False when the deck was refused
     * because a slide did not parse.
     */
    finish(): boolean {
      const owner = editor.session();
      const open = owner?.opening;
      if (!owner || !open) return true;
      const rejection = open.complete();
      if (rejection) return (refuseNow(owner, rejection), false);
      settle(owner);
      editor.emit();
      return true;
    },
    /** The slides of the open deck still opening, for `EditorSnapshot.pendingSlides`. */
    pending(): number {
      const owner = editor.session();
      const open = owner?.opening;
      return owner && open ? open.pendingSlides(owner.store.package, owner.info) : 0;
    },
  };
}
