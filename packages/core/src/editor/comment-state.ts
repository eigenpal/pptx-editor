// The editor's comment state: the author of new comments (`setAuthor`), the thread a marker or
// the comments pane picked (`selectComment`), and the current slide's threads for the snapshot,
// reference-stable while the parts they come from are the same objects. The threads show only
// with a review module; without one, the snapshot reports null and the commands still work.

import { presentationInfo, relsPartName, type PptxPackage } from '../store/index.ts';
import type { CommentThread, EditorSnapshot, SlideComment } from '../contracts/index.ts';
import type { CommentMarkerHost } from './comment-markers.ts';
import { requestPane } from './pane-requests.ts';
import {
  COMMENTS,
  identityGroup,
  presentationTarget,
  readSlideComments,
  slideCommentParts,
} from '../automation/commands/comment-read.ts';
import type { EditorModuleRegistry } from './modules.ts';

const NONE: readonly CommentThread[] = Object.freeze([]);

export class CommentState {
  #author: string | undefined;
  #active: string | null = null;
  /** The `identity` of the picked thread, once a snapshot has read it. */
  #activeIdentity: string | null = null;
  #key: readonly unknown[] = [];
  #threads: readonly CommentThread[] = NONE;
  /** The slide part of `#threads`. */
  #slide: string | undefined;
  /** The identities that named a comment of this deck that is gone: never shown again. */
  readonly #retired = new Set<string>();
  /** A count for the fresh identities that `#carry` makes. */
  #fresh = 0;

  constructor(
    author: string | undefined,
    private readonly modules: EditorModuleRegistry
  ) {
    this.#author = typeof author === 'string' && author.trim() ? author : undefined;
  }

  /** The author of new comments: the host's, else the collaboration identity's name. */
  get author(): string | undefined {
    return this.#author ?? this.modules.collaboration?.session.identity.name;
  }

  /** Set the host's author. Returns whether it changed. */
  setAuthor(author: string | undefined): boolean {
    const next = typeof author === 'string' && author.trim() ? author : undefined;
    if (next === this.#author) return false;
    this.#author = next;
    return true;
  }

  /** Pick a thread (a marker press, the pane), or none. Returns whether it changed. */
  select(id: string | null): boolean {
    const next = typeof id === 'string' ? id : null;
    if (next === this.#active) return false;
    this.#active = next;
    // The id names a thread of the last threads read, the ones its caller saw.
    this.#activeIdentity = this.#threads.find((t) => t.id === next)?.identity ?? null;
    return true;
  }

  /** The editor's `setAuthor` and `selectComment`: each calls `emit` when it changes state. */
  methods(emit: () => void): {
    setAuthor(author: string | undefined): void;
    selectComment(id: string | null): void;
  } {
    return {
      setAuthor: (author) => void (this.setAuthor(author) && emit()),
      selectComment: (id) => void (this.select(id) && emit()),
    };
  }

  /** The context fields of a command: the author, when there is one. */
  context(): { readonly author?: string } {
    const author = this.author;
    return author ? { author } : {};
  }

  /** The threads of `slidePart`, the same array while their parts are the same objects. */
  threads(pkg: PptxPackage | undefined, slidePart: string | undefined): readonly CommentThread[] {
    if (!pkg || !slidePart) return NONE;
    const { legacy, modern } = slideCommentParts(pkg, slidePart);
    const presentation = presentationInfo(pkg).presentationPart;
    const key = [
      pkg.parts.get(slidePart),
      pkg.parts.get(relsPartName(slidePart)),
      pkg.parts.get(relsPartName(presentation)),
      ...[COMMENTS.legacyAuthorsRel, COMMENTS.modernAuthorsRel].map((type) => {
        const part = presentationTarget(pkg, type);
        return part ? pkg.parts.get(part) : undefined;
      }),
      ...[...legacy, ...modern].map((part) => pkg.parts.get(part)),
    ];
    if (key.length === this.#key.length && key.every((k, i) => k === this.#key[i]))
      return this.#threads;
    this.#key = key;
    const read = legacy.length || modern.length ? readSlideComments(pkg, slidePart) : NONE;
    const threads = this.#carry(read, this.#slide === slidePart ? this.#threads : NONE);
    this.#slide = slidePart;
    this.#threads = threads.length ? Object.freeze(threads) : NONE;
    return this.#threads;
  }

  /**
   * `read` with the identities of `before` carried over. The file alone tells twins (comments with
   * one GUID, author, and date) apart only by their text and place, so an edit of a twin, or a
   * delete of another, changes the identity that the reader gives it. Per group of twins:
   * - With as many comments as before, each keeps the identity of the one in its place: an
   *   edit or a resolve changes no count.
   * - With more or fewer, each text that counts the same as before keeps its identities in
   *   order, and a text that comes or goes alone is a new or a deleted comment.
   * - When that does not tell which twin is which (one of two twins with the same text is
   *   gone), every twin of the group takes a new identity, so no open field or pick moves to
   *   another comment.
   * An identity that is gone is never shown again.
   */
  #carry(
    read: readonly CommentThread[],
    before: readonly CommentThread[]
  ): readonly CommentThread[] {
    if (read.length === 0 && before.length === 0) return read;
    const flat = (threads: readonly CommentThread[]) =>
      threads.flatMap((t) => [t as SlideComment, ...t.replies]);
    const old = new Map<string, SlideComment[]>();
    for (const c of flat(before)) {
      const group = identityGroup(c.identity);
      old.set(group, [...(old.get(group) ?? []), c]);
    }
    const now = new Map<string, SlideComment[]>();
    for (const c of flat(read)) {
      const group = identityGroup(c.identity);
      now.set(group, [...(now.get(group) ?? []), c]);
    }
    const shown = new Map<SlideComment, string>();
    const used = new Set<string>();
    const fresh = (c: SlideComment): string => `${c.identity}~g${(this.#fresh += 1)}`;
    for (const [group, members] of now) {
      const was = old.get(group) ?? [];
      const take = new Map<SlideComment, string>();
      if (was.length === members.length) members.forEach((c, i) => take.set(c, was[i]!.identity));
      else {
        const byText = (list: readonly SlideComment[]) => {
          const m = new Map<string, SlideComment[]>();
          for (const c of list) m.set(c.text, [...(m.get(c.text) ?? []), c]);
          return m;
        };
        const prev = byText(was);
        const next = byText(members);
        const texts = new Set([...prev.keys(), ...next.keys()]);
        const clear = [...texts].every((t) => {
          const a = prev.get(t)?.length ?? 0;
          const b = next.get(t)?.length ?? 0;
          return a === b || a === 0 || b === 0;
        });
        if (!clear) {
          for (const c of members) shown.set(c, fresh(c));
          continue;
        }
        for (const [t, list] of next) {
          const from = prev.get(t);
          if (from?.length === list.length) list.forEach((c, i) => take.set(c, from[i]!.identity));
        }
      }
      for (const c of members) {
        const carried = take.get(c);
        shown.set(c, carried ?? c.identity);
      }
    }
    // A gone identity retires; one that a carry or a retire makes unsafe takes a fresh one.
    const keep = new Set(shown.values());
    for (const c of flat(before)) if (!keep.has(c.identity)) this.#retired.add(c.identity);
    for (const [c, id] of shown) {
      const safe = !used.has(id) && !(this.#retired.has(id) && id === c.identity);
      const final = safe ? id : fresh(c);
      used.add(final);
      shown.set(c, final);
    }
    const rename = (c: SlideComment): SlideComment =>
      shown.get(c) === c.identity ? c : { ...c, identity: shown.get(c)! };
    return read.map((t) => {
      const head = rename(t) as CommentThread;
      const replies = t.replies.map(rename);
      return head === t && replies.every((r, i) => r === t.replies[i]) ? t : { ...head, replies };
    });
  }

  /**
   * The snapshot fields: the threads (null without a review module), the picked thread, and the
   * name that signs a new comment.
   */
  snapshot(
    pkg: PptxPackage | undefined,
    slidePart: string | undefined
  ): {
    comments: readonly CommentThread[] | null;
    activeComment: string | null;
    commentAuthor: string | null;
  } {
    const commentAuthor = this.author ?? null;
    if (!this.modules.review) return { comments: null, activeComment: null, commentAuthor };
    const comments = this.threads(pkg, slidePart);
    return { comments, activeComment: this.#pickedIn(comments), commentAuthor };
  }

  /**
   * The id of the picked thread in `comments`, or null. The pick follows the thread's
   * `identity`, so it stays when a delete of another repeat changes the thread's id, and it ends
   * when another thread takes the id.
   */
  #pickedIn(comments: readonly CommentThread[]): string | null {
    if (this.#active === null) return null;
    const known = this.#activeIdentity;
    const hit = known
      ? comments.find((t) => t.identity === known)
      : comments.find((t) => t.id === this.#active);
    if (!hit) {
      // The thread is gone: a later pick of the comment that takes its id is a change.
      this.#active = null;
      this.#activeIdentity = null;
      return null;
    }
    this.#active = hit.id;
    this.#activeIdentity = hit.identity;
    return hit.id;
  }

  /**
   * What the slide's comment markers read: the snapshot's threads, the review module's marker
   * style, and a press that picks a thread and, unless the module's `opening` is `manual`, asks
   * for the Comments tab (`requestPane`).
   */
  markers(
    snapshot: () => EditorSnapshot,
    editor: { selectComment(id: string): void }
  ): CommentMarkerHost {
    return {
      threads: () => snapshot().comments ?? NONE,
      active: () => snapshot().activeComment ?? null,
      style: () => this.modules.review?.pane?.commentMarkers ?? 'icon',
      pick: (id) => {
        editor.selectComment(id);
        if (this.modules.review?.pane?.opening !== 'manual') requestPane(editor, 'comments');
      },
    };
  }
}

/**
 * The comment threads of a slide part, in file order: its legacy comments (ECMA-376 Part 1,
 * 19.4), then its threaded comments. For hosts and modules that read a package without an
 * editor. Text and names come from the file: show them as text. @public
 */
export function readCommentThreads(pkg: PptxPackage, slidePart: string): readonly CommentThread[] {
  return readSlideComments(pkg, slidePart);
}
