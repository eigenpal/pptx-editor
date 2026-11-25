/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The comments of a slide: `Comment`, `CommentReply`, and their collections. An SDK extension:
// the pinned Office.js presentation reference has no comment members. The member names follow
// the comment API of DOCX Editor (`authorName`, `content`, `creationDate`, `resolved`, `replies`,
// `reply`, `delete`).
//
// Two file formats hold comments. The legacy format is ECMA-376 Part 1, 19.4: `p:cm`
// (`CT_Comment`, 19.4.1) in a slide's comments part (13.3.2), by a `p:cmAuthor`
// (`CT_CommentAuthor`, 19.4.2), at a `p:pos` (19.4.5). It has no resolved state and no text
// anchor. The threaded format is a part that the slide names in its `p:extLst` (19.3.1.20); it is
// not part of ECMA-376. Reads cover both; a new thread uses the format that the deck uses.
//
// A comment proxy's path resolves to its slide, and the proxy keeps the comment's id key, as a
// `Tag` keeps its key. A write sends what a read showed of the comment (`expect`), so the host
// refuses the write when another comment took the id since that read.

import type {
  AutomationCommentAnchor,
  AutomationCommentExpect,
  AutomationCommentRead,
  AutomationCommentScope,
  AutomationCommentThreadRead,
  AutomationOperation,
} from '@pptx-editor.dev/core/automation';
import { ClientResult } from '../runtime/client-result.ts';
import type { ClientObject } from '../runtime/client-object.ts';
import { fail, failArgument, nearest } from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import { LOAD_REQUEST } from '../runtime/internals.ts';
import { itemRequest, type ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import { flag, ModelObject, pathOf, type Navigation, type RefOf } from './model-object.ts';

/**
 * Where a comment thread points on its slide.
 *
 * - `slide`: the slide as a whole.
 * - `point`: a place on the slide, in points from its top left corner. A legacy comment stores
 *   one (`p:pos`, ECMA-376 Part 1, 19.4.5).
 * - `shape`: a shape of the slide, by its id (`Shape.id`).
 * - `text`: a stretch of a shape's text, by the shape's id and the `start` and `length` that
 *   `TextRange` uses. Only the threaded format holds it.
 *
 * A thread whose shape is gone points at the slide.
 *
 * @public
 */
export type CommentAnchor =
  | { readonly type: 'slide' }
  | { readonly type: 'point'; readonly left: number; readonly top: number }
  | { readonly type: 'shape'; readonly shapeId: string }
  | {
      readonly type: 'text';
      readonly shapeId: string;
      readonly start: number;
      readonly length: number;
    };

/**
 * Who writes a new comment or reply. Each field is optional.
 *
 * @public
 */
export interface CommentAuthorOptions {
  /**
   * The display name, 1 to 255 characters. Default: the `author` option of the runtime, else, in
   * a browser runtime, the name that signs the editor's own comments. With none, the write
   * refuses with `NotSupported`.
   */
  authorName?: string;
  /** The initials, up to 8 characters. Default: the first letters of the first two words. */
  authorInitials?: string;
}

/**
 * Options for `CommentCollection.add`.
 *
 * @public
 */
export interface CommentAddOptions extends CommentAuthorOptions {
  /**
   * The shape the thread points at: a shape on the top level of the same slide, or its id.
   * Default: the slide as a whole. On a deck with only legacy comments, the thread stores a
   * place near the shape (`p:pos`), and `anchor` reads it back as a `point`.
   */
  shape?: ClientObject | string;
  /**
   * With `shape`: the stretch of its text the thread points at, as `TextRange.start` and
   * `TextRange.length` count it. The length is 1 or more. A deck with only legacy comments has
   * no text anchors, and refuses with `NotSupported`.
   */
  textRange?: { start: number; length: number };
}

/** The values a read of the comment list gives one comment. Module-private. */
type Read = AutomationCommentRead | AutomationCommentThreadRead;

/** Fills a comment a read found. Module-private, so no consumer can call it. */
const FILLERS = new WeakMap<CommentBase, (read: Read) => void>();

function fillComment(comment: CommentBase, read: Read): void {
  FILLERS.get(comment)?.(read);
}

/** Gives a comment that a write created its id key. Module-private. */
const KEYS = new WeakMap<CommentBase, (key: string) => void>();

function takeKey(comment: CommentBase, key: string): void {
  KEYS.get(comment)?.(key);
}

/** The comment threads a `readComments` answer holds. */
function threadsOf(value: Parameters<typeof hydrated>[0], label: string) {
  return hydrated(value, 'comments', label).comments;
}

/**
 * A stored `xsd:dateTime` as a `Date`, or null when it is absent or not a date. A date with no
 * zone counts as UTC, as presentation files store comment dates.
 */
function dateOf(raw: string | undefined): Date | null {
  if (raw === undefined) return null;
  const m = /^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})(\.\d{1,9})?(Z|[+-]\d{2}:\d{2})?$/.exec(raw);
  if (!m) return null;
  const fraction = m[2] ? m[2].slice(0, 4) : '';
  const at = Date.parse(`${m[1]}${fraction}${m[3] ?? 'Z'}`);
  return Number.isNaN(at) ? null : new Date(at);
}

function anchorOf(anchor: AutomationCommentAnchor): CommentAnchor {
  switch (anchor.kind) {
    case 'slide':
      return { type: 'slide' };
    case 'point':
      return { type: 'point', left: anchor.left, top: anchor.top };
    case 'shape':
      return { type: 'shape', shapeId: anchor.shapeId };
    case 'text':
      return {
        type: 'text',
        shapeId: anchor.shapeId,
        start: anchor.start,
        length: anchor.length,
      };
  }
}

/** Comment text: a string. The host checks its length. */
function contentOf(value: unknown, target: string): string {
  if (typeof value !== 'string') failArgument(target, 'comment text of 1 to 10000 characters');
  return value;
}

/** The author options of a write, checked, as operation fields. */
function authorFields(
  options: CommentAuthorOptions | undefined,
  target: string
): { author?: string; initials?: string } {
  if (options === undefined) return {};
  if (typeof options !== 'object' || options === null) failArgument(target, 'an options object');
  const { authorName, authorInitials } = options;
  if (authorName !== undefined && typeof authorName !== 'string')
    failArgument(`${target}.authorName`, 'an author name of 1 to 255 characters');
  if (authorInitials !== undefined && typeof authorInitials !== 'string')
    failArgument(`${target}.authorInitials`, 'initials of at most 8 characters');
  return {
    ...(authorName === undefined ? {} : { author: authorName }),
    ...(authorInitials === undefined ? {} : { initials: authorInitials }),
  };
}

/** A comment id: a non-empty string. */
function idOf(value: unknown, target: string): string {
  if (typeof value !== 'string' || value.length === 0)
    failArgument(target, 'the id of a comment, as Comment.id gives it');
  return value;
}

/** What a comment and a reply share: an id, an author, a date, and a text. */
abstract class CommentBase extends ModelObject {
  /** @internal The slide the comment is on. */
  protected readonly slide: ObjectPath;
  /** @internal The comment's id key, once known. */
  protected key: string;
  /** @internal What the last read showed of the comment, for the checks of a write. */
  protected expected: AutomationCommentExpect | undefined;

  /** @internal */
  protected constructor(
    context: RequestContext,
    path: ObjectPath,
    slide: ObjectPath,
    key: string,
    nullable: boolean
  ) {
    super(context, path, { nullable });
    this.slide = slide;
    this.key = key;
    if (key) this.setLoadedProperty('id', key);
    FILLERS.set(this, (read) => this.fillFrom(read));
    KEYS.set(this, (k) => {
      this.key = k;
      this.setLoadedProperty('id', k);
    });
  }

  /** @internal Take the values a read gave. */
  protected fillFrom(read: Read): void {
    this.key = read.id;
    this.expected = {
      author: read.author,
      text: read.text,
      ...(read.date === undefined ? {} : { date: read.date }),
    };
    this.setLoadedProperty('id', read.id);
    this.setLoadedProperty('authorName', read.author);
    this.setLoadedProperty('authorInitials', read.initials);
    this.setLoadedProperty('content', read.text);
    this.setLoadedProperty('creationDate', dateOf(read.date));
  }

  /**
   * The comment's id on its slide: the GUID of a threaded comment, or `cm-<authorId>-<idx>` for
   * a legacy comment, from the numbers `p:cm/@authorId` and `p:cm/@idx` (ECMA-376 Part 1,
   * 19.4.1). It stays the same through save and reopen.
   */
  get id(): string {
    return this.loadedProperty('id');
  }
  /** The author's display name, or an empty string when the file names no author. */
  get authorName(): string {
    return this.loadedProperty('authorName');
  }
  /** The author's initials, or an empty string when the file has none. */
  get authorInitials(): string {
    return this.loadedProperty('authorInitials');
  }
  /**
   * When it was written, or null when the file records no valid date. A stored date with no zone
   * counts as UTC.
   */
  get creationDate(): Date | null {
    return this.loadedProperty('creationDate');
  }

  /**
   * The text, with a line feed between paragraphs. Text from a file is untrusted: show it as
   * text, never as markup.
   *
   * A write replaces the text at the next `context.sync()`. It takes 1 to 10,000 characters; a
   * line feed starts a new paragraph. A comment body that holds more than plain text, such as a
   * mention, refuses with `NotSupported`.
   */
  get content(): string {
    return this.loadedProperty('content');
  }
  set content(value: string) {
    const text = contentOf(value, this.at('content'));
    const expect = this.expected;
    this.command('content', (ref) => ({
      op: 'setCommentText',
      slide: ref(this.path),
      comment: this.key,
      scope: this.writeScope(),
      text,
      ...(expect === undefined ? {} : { expect }),
    }));
    // The host stores the text cleaned and trimmed, so a later write of this run does not
    // check the comment against the text this write sent. A new load reads it again.
    this.expected = undefined;
  }

  /**
   * Delete the comment at the next `context.sync()`. On a thread, this deletes its replies too.
   * On a reply, it deletes only that reply.
   */
  delete(): void {
    const expect = this.expected;
    this.command('delete', (ref) => ({
      op: 'deleteComment',
      slide: ref(this.path),
      comment: this.key,
      scope: this.writeScope(),
      ...(expect === undefined ? {} : { expect }),
    }));
  }

  /**
   * @internal What a write may reach: a thread of the slide, or a reply of the parent thread. A
   * write sends it, so the host refuses a key outside that scope.
   */
  protected abstract writeScope(): AutomationCommentScope;

  /** @internal The values of this comment in a `readComments` answer, or undefined. */
  protected abstract pickFrom(threads: readonly AutomationCommentThreadRead[]): Read | undefined;

  /** @internal The scalar properties a load can name. */
  protected abstract get scalars(): readonly string[];

  /** @internal The navigation properties a load can name. */
  protected get navigations(): Navigation {
    return {};
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, this.scalars, this.navigations);
    if (names.length === 0 || (names.every((n) => n === 'id') && this.hasLoadedProperty('id')))
      return;
    const label = `${this.path.label}.load`;
    this.enqueue(
      'read',
      'load',
      (ref) => ({ op: 'readComments', slide: ref(this.slide) }),
      (value) => {
        if (value.kind === 'null') return;
        const found = this.pickFrom(threadsOf(value, label));
        if (!found) fail({ code: 'ItemNotFound', target: label });
        this.fillFrom(found);
      }
    );
  }
}

const REPLY_SCALARS = ['id', 'authorName', 'authorInitials', 'content', 'creationDate'] as const;
const COMMENT_SCALARS = [...REPLY_SCALARS, 'resolved', 'anchor'] as const;

/**
 * One reply in a comment thread. An SDK extension.
 *
 * @public
 */
export class CommentReply extends CommentBase {
  readonly #thread: () => string;

  /** @internal */
  constructor(
    context: RequestContext,
    path: ObjectPath,
    slide: ObjectPath,
    thread: () => string,
    key: string,
    nullable = false
  ) {
    super(context, path, slide, key, nullable);
    this.#thread = thread;
  }

  /** @internal */
  protected get scalars(): readonly string[] {
    return REPLY_SCALARS;
  }

  /** @internal */
  protected pickFrom(threads: readonly AutomationCommentThreadRead[]): Read | undefined {
    const thread = threads.find((t) => t.id === this.#thread());
    return thread?.replies.find((r) => r.id === this.key);
  }

  /** @internal */
  protected writeScope(): AutomationCommentScope {
    return { replyOf: this.#thread() };
  }
}

/** Fills a thread's state and anchor. Module-private. */
const THREAD_FILLERS = new WeakMap<Comment, (read: AutomationCommentThreadRead) => void>();

/**
 * A comment thread of a slide: its first comment, its state, its anchor, and its replies. An SDK
 * extension: the pinned Office.js reference has no comment members.
 *
 * @example
 * ```ts
 * await runtime.run(async (context) => {
 *   const comments = context.presentation.slides.getItemAt(0).comments;
 *   comments.load('items/authorName,items/content,items/resolved');
 *   await context.sync();
 *   for (const comment of comments.items)
 *     if (!comment.resolved) comment.reply('Fixed in the next draft.');
 *   await context.sync();
 * });
 * ```
 *
 * @public
 */
export class Comment extends CommentBase {
  #replies: CommentReplyCollection | undefined;

  /** @internal */
  constructor(
    context: RequestContext,
    path: ObjectPath,
    slide: ObjectPath,
    key: string,
    nullable = false
  ) {
    super(context, path, slide, key, nullable);
    THREAD_FILLERS.set(this, (read) => {
      this.setLoadedProperty('resolved', read.resolved);
      this.setLoadedProperty('anchor', anchorOf(read.anchor));
    });
  }

  /** @internal */
  protected override fillFrom(read: Read): void {
    super.fillFrom(read);
    if ('replies' in read) THREAD_FILLERS.get(this)?.(read);
  }

  /** @internal */
  protected get scalars(): readonly string[] {
    return COMMENT_SCALARS;
  }

  /** @internal */
  protected override get navigations(): Navigation {
    return { replies: () => this.replies };
  }

  /** @internal */
  protected pickFrom(threads: readonly AutomationCommentThreadRead[]): Read | undefined {
    return threads.find((t) => t.id === this.key);
  }

  /** @internal */
  protected writeScope(): AutomationCommentScope {
    return 'thread';
  }

  /**
   * Whether the thread is resolved. Always false for a legacy comment, which has no state.
   *
   * A write resolves (`true`) or reopens (`false`) the whole thread at the next
   * `context.sync()`. A write of the current state writes nothing. On a legacy comment, the
   * write refuses with `NotSupported`; reply to the thread instead.
   */
  get resolved(): boolean {
    return this.loadedProperty('resolved');
  }
  set resolved(value: boolean) {
    const resolved = flag(value, this.at('resolved'));
    const expect = this.expected;
    this.command('resolved', (ref) => ({
      op: 'setCommentResolved',
      slide: ref(this.path),
      comment: this.key,
      scope: 'thread',
      resolved,
      ...(expect === undefined ? {} : { expect }),
    }));
  }

  /** Where the thread points on its slide. */
  get anchor(): CommentAnchor {
    return this.loadedProperty('anchor');
  }

  /** The replies, oldest first. */
  get replies(): CommentReplyCollection {
    return (this.#replies ??= new CommentReplyCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.replies`, this.path),
      this.slide,
      () => this.key
    ));
  }

  /**
   * Reply to the thread at the next `context.sync()`. The text takes 1 to 10,000 characters; a
   * line feed starts a new paragraph. The reply keeps the format of its thread.
   *
   * The answer is the new reply. Call its members after the sync that adds it.
   */
  reply(content: string, options?: CommentAuthorOptions): CommentReply {
    const label = `${this.path.label}.reply`;
    const text = contentOf(content, label);
    const author = authorFields(options, label);
    const expect = this.expected;
    const path = ObjectPath.pending(`${label}()`);
    const reply = new CommentReply(this.context, path, this.slide, () => this.key, '');
    this.enqueue(
      'write',
      'reply',
      (ref) => ({
        op: 'replyToComment',
        slide: ref(this.path),
        comment: this.key,
        scope: 'thread',
        text,
        ...author,
        ...(expect === undefined ? {} : { expect }),
      }),
      (value) => {
        takeKey(reply, hydrated(value, 'text', label).text);
        path.resolveTo(this.slide.handle());
      },
      path,
      false
    );
    return reply;
  }
}

/** What both comment collections share: the reads of the comment list of one slide. */
abstract class CommentList<T extends CommentBase> extends ModelObject {
  /** @internal */
  protected readonly slide: ObjectPath;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, slide: ObjectPath) {
    super(context, path);
    this.slide = slide;
  }

  /** @internal The members in a `readComments` answer, in order. */
  protected abstract membersOf(threads: readonly AutomationCommentThreadRead[]): readonly Read[];
  /** @internal A proxy for one member. */
  protected abstract make(path: ObjectPath, key: string, nullable: boolean): T;
  /** @internal The scalar properties of a member. */
  protected abstract get scalars(): readonly string[];

  /** The loaded members. */
  get items(): T[] {
    return this.loadedProperty<T[]>('items');
  }

  #listing(ref: RefOf): AutomationOperation {
    return { op: 'readComments', slide: ref(this.slide) };
  }

  /** The number of members, after the next `sync()`. */
  getCount(): ClientResult<number> {
    const label = `${this.path.label}.getCount`;
    const { result, fill } = ClientResult.create<number>(label);
    this.enqueue(
      'read',
      'getCount',
      (ref) => this.#listing(ref),
      (value) => fill(this.membersOf(threadsOf(value, label)).length)
    );
    return result;
  }

  /** The member with this id; `ItemNotFound` at the sync when there is none. */
  getItem(id: string): T {
    const key = idOf(id, `${this.path.label}.getItem`);
    // The id is known now, so the member needs no lookup: a write to it goes in the same batch
    // as the other writes of the sync, and the host refuses the whole batch when no comment
    // has the id. A load refuses with `ItemNotFound` then. Over an owner that is addressable,
    // the member has a path of its own, so tracking it keeps it. Over an owner that a lookup of
    // this sync finds (a slide, or the thread of a reply), it shares the owner's path, so a
    // write waits for that lookup and knows the thread it is in.
    const label = `${this.path.label}.getItem()`;
    const path = this.path.isAddressable
      ? ObjectPath.of(label, this.path.handle())
      : ObjectPath.derived(label, this.path);
    return this.make(path, key, false);
  }

  /** The member at a zero-based position; `ItemNotFound` at the sync when there is none. */
  getItemAt(index: number): T {
    if (typeof index !== 'number' || !Number.isInteger(index) || index < 0)
      failArgument(`${this.path.label}.getItemAt`, 'a whole number of 0 or more');
    return this.find(`getItemAt(${index})`, (all) => all[index], '', false);
  }

  /** The member with this id, or a null object. */
  getItemOrNullObject(id: string): T {
    const key = idOf(id, `${this.path.label}.getItemOrNullObject`);
    return this.find('getItemOrNullObject()', (all) => all.find((c) => c.id === key), key, true);
  }

  /** A member over a pending path that a read of the list resolves to the slide, or null. */
  protected find(
    name: string,
    pick: (all: readonly Read[]) => Read | undefined,
    key: string,
    orNull: boolean
  ): T {
    const path = ObjectPath.pending(`${this.path.label}.${name}`);
    const member = this.make(path, key, orNull);
    const label = path.label;
    this.enqueue(
      'read',
      name,
      (ref) => this.#listing(ref),
      (value) => {
        const found = pick(this.membersOf(threadsOf(value, label)));
        if (!found) {
          if (orNull) return path.resolveNull();
          fail({ code: 'ItemNotFound', target: label });
        }
        fillComment(member, found);
        path.resolveTo(this.slide.handle());
      },
      path,
      false
    );
    return member;
  }

  /** @internal */
  protected override get loadsItemPaths(): boolean {
    return true;
  }

  /** @internal The navigation properties of a member, which `items/<name>/...` may load. */
  protected get memberNavigations(): readonly string[] {
    return [];
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const accepted = ['items', ...this.scalars, ...this.memberNavigations];
    const nested = [...request.nested.keys(), ...(request.items?.nested.keys() ?? [])];
    for (const name of [...request.select, ...request.itemSelect, ...nested])
      if (!accepted.includes(name))
        failArgument(
          `${this.path.label}.${name}`,
          `a property of comments: ${accepted.join(', ')}`,
          nearest(name, accepted)
        );
    this.takeNestedPaths();
    // A navigation of the members loads in a second read, after the listing named them. The
    // listing fills their scalar properties, so the second read asks only for the navigation.
    const navigations = this.memberNavigations;
    const navigates =
      nested.length > 0 || request.itemSelect.some((name) => navigations.includes(name));
    const forItems = navigates ? itemRequest(request) : undefined;
    const deeper = forItems && {
      ...forItems,
      select: forItems.select.filter((name) => navigations.includes(name)),
      all: false,
      itemSelect: [],
    };
    const label = `${this.path.label}.items`;
    this.enqueue(
      'read',
      'items',
      (ref) => this.#listing(ref),
      (value) => {
        const all = this.membersOf(threadsOf(value, label));
        const end = request.top === undefined ? all.length : (request.skip ?? 0) + request.top;
        const loaded = all.slice(request.skip ?? 0, end).map((read, i) => {
          const member = this.make(
            ObjectPath.of(`${label}[${i}]`, this.slide.handle()),
            read.id,
            false
          );
          fillComment(member, read);
          return member;
        });
        this.setLoadedProperty('items', loaded);
        if (deeper)
          this.internals.queue.continuing(() => {
            for (const member of loaded) member[LOAD_REQUEST](deeper);
          });
      },
      undefined,
      true,
      deeper !== undefined
    );
  }
}

/**
 * The comment threads of a slide, in file order: the legacy comments first, then the threaded
 * ones. An SDK extension.
 *
 * @example
 * ```ts
 * await runtime.run(async (context) => {
 *   const slide = context.presentation.slides.getItemAt(0);
 *   const title = slide.shapes.getItemAt(0);
 *   slide.comments.add('Shorten this title.', { shape: title, authorName: 'Review Agent' });
 *   await context.sync();
 * });
 * ```
 *
 * @public
 */
export class CommentCollection extends CommentList<Comment> {
  readonly #shapeById: (id: string) => ClientObject;

  /** @internal */
  constructor(
    context: RequestContext,
    path: ObjectPath,
    slide: ObjectPath,
    shapeById: (id: string) => ClientObject
  ) {
    super(context, path, slide);
    this.#shapeById = shapeById;
  }

  /** @internal */
  protected membersOf(threads: readonly AutomationCommentThreadRead[]): readonly Read[] {
    return threads;
  }

  /** @internal */
  protected make(path: ObjectPath, key: string, nullable: boolean): Comment {
    return new Comment(this.context, path, this.slide, key, nullable);
  }

  /** @internal */
  protected get scalars(): readonly string[] {
    return COMMENT_SCALARS;
  }

  /** @internal */
  protected override get memberNavigations(): readonly string[] {
    return ['replies'];
  }

  /**
   * Add a comment thread at the next `context.sync()`. The text takes 1 to 10,000 characters; a
   * line feed starts a new paragraph.
   *
   * The thread uses the format that the deck uses: threaded, unless the deck has only legacy
   * comments. It is one undo step in a browser runtime. The answer is the new thread. Call its
   * members after the sync that adds it.
   */
  add(content: string, options?: CommentAddOptions): Comment {
    const label = `${this.path.label}.add`;
    const text = contentOf(content, label);
    const author = authorFields(options, label);
    let shape: ClientObject | undefined;
    let span: { start: number; length: number } | undefined;
    if (options?.shape !== undefined) {
      const given = options.shape;
      // A shape id looks the shape up in the same sync; an object must be one of this API.
      if (typeof given === 'string') shape = this.#shapeById(idOf(given, `${label}.shape`));
      else {
        pathOf(given);
        shape = given;
      }
    }
    if (options?.textRange !== undefined) {
      const range = options.textRange as { start?: unknown; length?: unknown } | null;
      const whole = (v: unknown, min: number): v is number =>
        typeof v === 'number' && Number.isSafeInteger(v) && v >= min;
      if (
        typeof range !== 'object' ||
        range === null ||
        !whole(range.start, 0) ||
        !whole(range.length, 1)
      )
        failArgument(`${label}.textRange`, 'a start of 0 or more and a length of 1 or more');
      if (shape === undefined) failArgument(`${label}.textRange`, 'a textRange with a shape');
      span = { start: range.start as number, length: range.length as number };
    }
    const path = ObjectPath.pending(`${label}()`);
    const comment = new Comment(this.context, path, this.slide, '');
    this.enqueue(
      'write',
      'add',
      (ref) => ({
        op: 'addComment',
        slide: ref(this.slide),
        text,
        ...(shape === undefined ? {} : { shape: ref(pathOf(shape)) }),
        ...(span === undefined ? {} : { span }),
        ...author,
      }),
      (value) => {
        takeKey(comment, hydrated(value, 'text', label).text);
        path.resolveTo(this.slide.handle());
      },
      path,
      false
    );
    return comment;
  }
}

/**
 * The replies of a comment thread, oldest first. An SDK extension.
 *
 * @public
 */
export class CommentReplyCollection extends CommentList<CommentReply> {
  readonly #thread: () => string;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, slide: ObjectPath, thread: () => string) {
    super(context, path, slide);
    this.#thread = thread;
  }

  /** @internal */
  protected membersOf(threads: readonly AutomationCommentThreadRead[]): readonly Read[] {
    const thread = threads.find((t) => t.id === this.#thread());
    if (!thread) fail({ code: 'ItemNotFound', target: this.path.label });
    return thread.replies;
  }

  /** @internal */
  protected make(path: ObjectPath, key: string, nullable: boolean): CommentReply {
    return new CommentReply(this.context, path, this.slide, this.#thread, key, nullable);
  }

  /** @internal */
  protected get scalars(): readonly string[] {
    return REPLY_SCALARS;
  }
}
