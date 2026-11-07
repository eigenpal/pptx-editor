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
  commentNodes,
  identityGroup,
  presentationTarget,
  readSlideComments,
  slideCommentParts,
} from '../automation/commands/comment-read.ts';
import type { EditorModuleRegistry } from './modules.ts';

const NONE: readonly CommentThread[] = Object.freeze([]);
/** How many slides other than the shown one keep their threads cached. */
const OTHERS = 4;

const same = (a: readonly unknown[], b: readonly unknown[]): boolean =>
  a.length === b.length && a.every((x, i) => x === b[i]);

/** A comment as a read showed it: its identity, twin group, text, and list node. */
interface Seen {
  readonly identity: string;
  readonly group: string;
  readonly text: string;
  readonly node: string | undefined;
  /** Whether it has been in a group of twins. */
  readonly twin: boolean;
}

/**
 * For each comment of `now`, the place in `was` of the same comment, or -1 for a new one: when
 * one change explains the group (see `CommentState#carry`), else null. When equal texts leave
 * more than one place for the change, the place whose mapping keeps each list node wins.
 */
function oneChange(
  was: readonly { readonly text: string; readonly node: string | undefined }[],
  now: readonly { readonly c: { readonly text: string }; readonly node: string | undefined }[]
): number[] | null {
  const a = was.length;
  const b = now.length;
  if (a === 0) return now.map(() => -1);
  if (a === b) {
    let changed = 0;
    for (let i = 0; i < a; i += 1) if (was[i]!.text !== now[i]!.c.text) changed += 1;
    return changed <= 1 ? now.map((_, i) => i) : null;
  }
  if (Math.abs(a - b) !== 1) return null;
  // The longer list less one entry at `at` is the shorter one: `at` must be the only place.
  const long = a > b ? was.map((w) => w.text) : now.map((n) => n.c.text);
  const short = a > b ? now.map((n) => n.c.text) : was.map((w) => w.text);
  let prefix = 0;
  while (prefix < short.length && long[prefix] === short[prefix]) prefix += 1;
  let suffix = 0;
  while (
    suffix < short.length &&
    long[long.length - 1 - suffix] === short[short.length - 1 - suffix]
  )
    suffix += 1;
  // The places from `low` to `prefix` all explain the change: equal texts surround them.
  const low = long.length - 1 - suffix;
  if (prefix < low) return null;
  const mapAt = (at: number): number[] =>
    a > b
      ? now.map((_, j) => (j < at ? j : j + 1))
      : now.map((_, j) => (j === at ? -1 : j < at ? j : j - 1));
  const keeps = (map: number[]): boolean =>
    map.every((from, j) => from < 0 || was[from]!.node === now[j]!.node);
  for (let at = low; at < prefix; at += 1) if (keeps(mapAt(at))) return mapAt(at);
  return mapAt(prefix);
}

export class CommentState {
  #author: string | undefined;
  #active: string | null = null;
  /** The `identity` of the picked thread, once a snapshot has read it. */
  #activeIdentity: string | null = null;
  #key: readonly unknown[] = [];
  #threads: readonly CommentThread[] = NONE;
  /**
   * What the last read of each slide part showed: one entry per slide part, so a slide that
   * comes back carries its identities as an edit does.
   */
  readonly #seen = new Map<string, readonly Seen[]>();
  /**
   * The identity that each comment id last named, and the ids that named another comment
   * before, until a snapshot shows them at their new identity.
   */
  readonly #named = new Map<string, string>();
  readonly #moved = new Set<string>();
  /** A count that moves each time `#moved` changes: the `can` answers depend on it. */
  #version = 0;
  /** The slide part of `#threads`. */
  #slide: string | undefined;
  /**
   * The threads of the last slides other than the shown one that a command checked, kept like
   * `#threads`, so a command on another slide does not read it again each time.
   */
  readonly #others = new Map<
    string,
    { key: readonly unknown[]; threads: readonly CommentThread[] }
  >();
  /** A count for the new identities that `#carry` makes: each is shown once only. */
  #fresh = 0;
  /** How many decks this editor opened before the current one (-1 before the first). */
  #loads = -1;

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

  /** A count that moves each time an id starts or stops naming another comment than before. */
  get version(): number {
    return this.#version;
  }

  /** Forget the threads and the pick of the last deck, when another deck opens. */
  reset(): void {
    this.#seen.clear();
    this.#named.clear();
    this.#moved.clear();
    this.#version += 1;
    this.#others.clear();
    this.#slide = undefined;
    this.#key = [];
    this.#threads = NONE;
    this.#active = null;
    this.#activeIdentity = null;
    this.#loads += 1;
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
  context(pkg: PptxPackage | undefined): {
    readonly author?: string;
    readonly commentAt: (commentId: string) => {
      readonly identity?: string;
      readonly moved: boolean;
    };
  } {
    const author = this.author;
    // The identity at an id, as the snapshot of the id's slide shows it, and whether the id
    // named another comment before.
    const commentAt = (commentId: string): { identity?: string; moved: boolean } => {
      const slidePart = commentId.slice(0, commentId.indexOf('#'));
      // Another slide's threads have a cache of their own, so the cached ones of the slide shown
      // stay, and the pane does not render again.
      const threads =
        !pkg || slidePart === this.#slide
          ? this.threads(pkg, slidePart)
          : this.#other(pkg, slidePart);
      let identity: string | undefined;
      for (const t of threads) {
        if (t.id === commentId) identity = t.identity;
        else identity = t.replies.find((r) => r.id === commentId)?.identity;
        if (identity !== undefined) break;
      }
      return {
        ...(identity !== undefined ? { identity } : {}),
        moved: this.#moved.has(commentId),
      };
    };
    return author ? { author, commentAt } : { commentAt };
  }

  /** The threads of `slidePart`, the same array while their parts are the same objects. */
  threads(pkg: PptxPackage | undefined, slidePart: string | undefined): readonly CommentThread[] {
    if (!pkg || !slidePart) return NONE;
    const key = this.#keyOf(pkg, slidePart);
    if (same(key, this.#key)) return this.#threads;
    this.#key = key;
    this.#slide = slidePart;
    // A slide that a command checked before it showed keeps the threads that the check read.
    const other = this.#others.get(slidePart);
    this.#others.delete(slidePart);
    this.#threads = other && same(other.key, key) ? other.threads : this.#read(pkg, slidePart);
    return this.#threads;
  }

  /** The threads of `slidePart`, a slide other than the shown one, cached like `threads`. */
  #other(pkg: PptxPackage, slidePart: string): readonly CommentThread[] {
    const key = this.#keyOf(pkg, slidePart);
    const hit = this.#others.get(slidePart);
    if (hit && same(hit.key, key)) return hit.threads;
    const threads = this.#read(pkg, slidePart);
    this.#others.delete(slidePart);
    this.#others.set(slidePart, { key, threads });
    // A few slides only: the oldest goes first.
    if (this.#others.size > OTHERS) this.#others.delete(this.#others.keys().next().value!);
    return threads;
  }

  /** The parts that the threads of `slidePart` come from. */
  #keyOf(pkg: PptxPackage, slidePart: string): readonly unknown[] {
    const { legacy, modern } = slideCommentParts(pkg, slidePart);
    const presentation = presentationInfo(pkg).presentationPart;
    return [
      pkg.parts.get(slidePart),
      pkg.parts.get(relsPartName(slidePart)),
      pkg.parts.get(relsPartName(presentation)),
      ...[COMMENTS.legacyAuthorsRel, COMMENTS.modernAuthorsRel].map((type) => {
        const part = presentationTarget(pkg, type);
        return part ? pkg.parts.get(part) : undefined;
      }),
      ...[...legacy, ...modern].map((part) => pkg.parts.get(part)),
    ];
  }

  /** The threads of `slidePart`, read now, with their identities carried. */
  #read(pkg: PptxPackage, slidePart: string): readonly CommentThread[] {
    const { legacy, modern } = slideCommentParts(pkg, slidePart);
    const read = legacy.length || modern.length ? readSlideComments(pkg, slidePart) : NONE;
    const threads = this.#carry(read, slidePart);
    return threads.length ? Object.freeze(threads) : NONE;
  }

  /**
   * `read` with the identities that the slide's last threads showed, carried over where one
   * change explains a group of twins (comments with one GUID, author, and date; a comment
   * without a twin is a group of one):
   * - The same count, with at most one text changed: an edit, a resolve, or a reply elsewhere.
   * - One comment more or less, the other texts the same and in order, at one place only.
   * Each carried twin must also be the same list node as before; a comment that never had a
   * twin needs no node. In every other case the whole group takes new identities, so open fields
   * close and the pick ends. A comment the slide did not show before takes a new identity too,
   * except on the first read of the slide. So a shown identity never names another comment.
   */
  #carry(read: readonly CommentThread[], slidePart: string): readonly CommentThread[] {
    const prior = this.#seen.get(slidePart);
    const items: { c: SlideComment; group: string; node: string | undefined }[] = [];
    for (const t of read) {
      items.push({ c: t, group: identityGroup(t.identity), node: commentNodes.get(t) });
      for (const r of t.replies)
        items.push({ c: r, group: identityGroup(r.identity), node: commentNodes.get(r) });
    }
    const shown = new Map<SlideComment, Seen>();
    const fresh = (c: SlideComment): string => `${c.identity}~g${(this.#fresh += 1)}`;
    const now = new Map<string, typeof items>();
    for (const item of items) {
      const list = now.get(item.group);
      if (list) list.push(item);
      else now.set(item.group, [item]);
    }
    const old = new Map<string, Seen[]>();
    for (const seen of prior ?? []) {
      const list = old.get(seen.group);
      if (list) list.push(seen);
      else old.set(seen.group, [seen]);
    }
    for (const [group, members] of now) {
      const was = old.get(group) ?? [];
      const twin = members.length > 1 || was.length > 1;
      const map = prior ? oneChange(was, members) : null;
      const safe =
        map !== null &&
        map.every((from, i) => {
          if (from < 0) return true;
          const before = was[from]!;
          // A comment that never had a twin is the only one of its group: no node to check.
          if (!before.twin && was.length === 1 && members.length === 1) return true;
          return before.node !== undefined && before.node === members[i]!.node;
        });
      members.forEach(({ c, node }, i) => {
        const from = safe ? map![i]! : -1;
        // A later deck's first read marks its identities, so none names a comment of an
        // earlier deck.
        const first = this.#loads > 0 ? `${c.identity}~l${this.#loads}` : c.identity;
        const identity = !prior ? first : from >= 0 ? was[from]!.identity : fresh(c);
        const wasTwin = from >= 0 && was[from]!.twin;
        shown.set(c, { identity, group, text: c.text, node, twin: twin || wasTwin });
      });
    }
    this.#seen.set(slidePart, [...shown.values()]);
    // An id that names another comment than before (a twin took a deleted twin's place) is
    // ambiguous for a host that read it earlier: `expect` then needs the identity.
    for (const [c, { identity }] of shown) {
      const before = this.#named.get(c.id);
      if (before !== undefined && before !== identity && !this.#moved.has(c.id)) {
        this.#moved.add(c.id);
        this.#version += 1;
      }
      this.#named.set(c.id, identity);
    }
    const rename = (c: SlideComment): SlideComment => {
      const identity = shown.get(c)!.identity;
      return identity === c.identity ? c : { ...c, identity };
    };
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
    // The snapshot shows each moved id at its new identity: the host reads it from here on.
    for (const id of this.#moved)
      if (id.startsWith(`${slidePart}#`) && this.#moved.delete(id)) this.#version += 1;
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
