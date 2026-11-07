// The editor's comment state: the author of new comments (`setAuthor`), the thread a marker or
// the comments pane picked (`selectComment`), and the current slide's threads for the snapshot,
// reference-stable while the parts they come from are the same objects. The threads show only
// with a review module; without one, the snapshot reports null and the commands still work.

import { presentationInfo, relsPartName, type PptxPackage } from '../store/index.ts';
import type { CommentThread, EditorSnapshot } from '../contracts/index.ts';
import type { CommentMarkerHost } from './comment-markers.ts';
import { requestPane } from './pane-requests.ts';
import {
  COMMENTS,
  presentationTarget,
  readSlideComments,
  slideCommentParts,
} from '../automation/commands/comment-read.ts';
import type { EditorModuleRegistry } from './modules.ts';

const NONE: readonly CommentThread[] = Object.freeze([]);

export class CommentState {
  #author: string | undefined;
  #active: string | null = null;
  #key: readonly unknown[] = [];
  #threads: readonly CommentThread[] = NONE;

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
    const threads = legacy.length || modern.length ? readSlideComments(pkg, slidePart) : NONE;
    this.#threads = threads.length ? Object.freeze(threads) : NONE;
    return this.#threads;
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
    const active = this.#active;
    return {
      comments,
      activeComment: active && comments.some((t) => t.id === active) ? active : null,
      commentAuthor,
    };
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
