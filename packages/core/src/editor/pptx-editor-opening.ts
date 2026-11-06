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
    void owner.addMissingSubstitutes().then((added) => {
      if (added && editor.session() === owner) return editor.facesChanged(owner);
    });
    const before = asked.get(owner) ?? new Set<string>();
    const more = deckFontFamilies(owner.store.package).filter(
      (family) => !before.has(family.toLowerCase())
    );
    if (more.length > 0) void editor.requestFamilies(more);
  };
  return {
    /** Parse the rest of `owner`'s deck in background tasks, if any is left. */
    start(owner: DocumentSession): void {
      const open = owner.opening;
      if (!open) return;
      // What the open asked for: the parts parsed so far, as `hostFonts.open` read them.
      const families = deckFontFamilies(owner.store.package).map((f) => f.toLowerCase());
      asked.set(owner, new Set(families));
      continueOpening(open, {
        current: () => editor.session() === owner && owner.opening === open,
        publish: () => editor.emit(),
        settled: () => (settle(owner), editor.emit()),
        failed: (rejection) => editor.refuse(owner, rejection),
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
      if (rejection) return (editor.refuse(owner, rejection), false);
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
