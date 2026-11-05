/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Office.js `Hyperlink`, `HyperlinkCollection`, and `HyperlinkScopedCollection`: the click
// links of a slide or of a text range, on text runs
// (`a:rPr/a:hlinkClick`) and on whole shapes (`p:cNvPr/a:hlinkClick`). An address goes
// through `sanitizeHref` in the host: only http, https, mailto, tel, and ftp links are written.

import type {
  AutomationHyperlinkRead,
  AutomationOperation,
  AutomationTextSpan,
} from '@pptx-editor.dev/core/automation';
import { ClientResult } from '../runtime/client-result.ts';
import { fail, failArgument, nearest } from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import type { HyperlinkType } from './enums.ts';
import { ModelObject, text, type RefOf } from './model-object.ts';
import { Shape } from './shapes.ts';
import { TextRange } from './text.ts';

/**
 * Options for a new hyperlink.
 *
 * @public
 */
export interface HyperlinkAddOptions {
  /** The address the link opens. */
  address?: string;
  /** The tip that shows over the link. */
  screenTip?: string;
}

/** A span, null for a shape link, or a getter for the range of a text link a write made. */
type LinkSpan = AutomationTextSpan | null | (() => AutomationTextSpan | undefined);

/** Fills a link a lookup found. Module-private, so no consumer can call it. */
const FILLERS = new WeakMap<Hyperlink, (read: AutomationHyperlinkRead) => void>();

/**
 * One hyperlink: on a text range or on a whole shape.
 *
 * @public
 */
export class Hyperlink extends ModelObject {
  readonly #shape: ObjectPath;
  /** The span, or undefined while a lookup has not answered (`getItemAt`). */
  #span: LinkSpan | undefined;
  /** Values this proxy wrote, which the answer of a pending lookup must not overwrite. */
  readonly #written = new Map<'address' | 'screenTip', string>();
  /** Objects that wait for a pending lookup: the linked shape and the linked text range. */
  readonly #waiting: ((read: AutomationHyperlinkRead) => void)[] = [];

  /**
   * @internal `values` absent: the link is a lookup that has not answered, and nothing is
   * loaded until it does.
   */
  constructor(
    context: RequestContext,
    path: ObjectPath,
    shape: ObjectPath,
    span: LinkSpan | undefined,
    values?: { address: string; screenTip: string }
  ) {
    super(context, path);
    this.#shape = shape;
    this.#span = span;
    if (values) {
      this.setLoadedProperty('address', values.address);
      this.setLoadedProperty('screenTip', values.screenTip);
      this.setLoadedProperty('type', span === null ? 'Shape' : 'TextRange');
    }
    FILLERS.set(this, (read) => {
      this.#span = spanOf(read);
      this.setLoadedProperty('address', this.#written.get('address') ?? read.address ?? '');
      this.setLoadedProperty('screenTip', this.#written.get('screenTip') ?? read.screenTip ?? '');
      this.setLoadedProperty('type', read.type === 'shape' ? 'Shape' : 'TextRange');
      for (const settle of this.#waiting.splice(0)) settle(read);
    });
  }

  /** @internal A link a write just made, with the values it wrote. */
  static created(
    context: RequestContext,
    shape: ObjectPath,
    span: LinkSpan,
    address: string,
    screenTip: string
  ): Hyperlink {
    return new Hyperlink(
      context,
      ObjectPath.derived(`${shape.label}.hyperlink`, shape),
      shape,
      span,
      { address, screenTip }
    );
  }

  /** The address the link opens. */
  get address(): string {
    return this.loadedProperty('address');
  }
  set address(value: string) {
    this.#write('address', text(value, this.at('address')));
  }
  /** The tip that shows over the link. */
  get screenTip(): string {
    return this.loadedProperty('screenTip');
  }
  set screenTip(value: string) {
    this.#write('screenTip', text(value, this.at('screenTip')));
  }
  /** Whether the link is on a text range or on a whole shape. */
  get type(): HyperlinkType | `${HyperlinkType}` {
    return this.loadedProperty('type');
  }

  /** Remove the link. The text or shape stays. */
  delete(): void {
    this.command('delete', (ref) => this.#op(ref, null, undefined));
  }

  /** The shape a shape link is on, or a null object for a text link. */
  getLinkedShapeOrNullObject(): Shape {
    const label = `${this.path.label}.getLinkedShapeOrNullObject()`;
    if (this.#span === undefined) {
      // The lookup has not answered: the same read settles the shape.
      const path = this.#shape.sibling(label);
      this.#waiting.push((read) =>
        read.type === 'shape' ? path.resolveTo(read.shape) : path.resolveNull()
      );
      return new Shape(this.context, path, true);
    }
    if (this.#spanNow() !== null) {
      const none = ObjectPath.pending(label);
      none.resolveNull();
      return new Shape(this.context, none, true);
    }
    return new Shape(this.context, ObjectPath.derived(label, this.#shape), true);
  }

  /** The text range a text link is on, or a null object for a shape link. */
  getLinkedTextRangeOrNullObject(): TextRange {
    const label = `${this.path.label}.getLinkedTextRangeOrNullObject()`;
    if (this.#span === undefined) {
      // The lookup has not answered: the same read settles the range and its span.
      const { range, settle } = TextRange.later(this.context, this.#shape.sibling(label));
      this.#waiting.push((read) => {
        const span = spanOf(read);
        settle(span ? { shape: read.shape, span } : null);
      });
      return range;
    }
    if (this.#spanNow() === null) {
      const none = ObjectPath.pending(label);
      none.resolveNull();
      return new TextRange(this.context, none, none, undefined, true);
    }
    return new TextRange(
      this.context,
      ObjectPath.derived(label, this.#shape),
      this.#shape,
      this.#spanNow() ?? undefined,
      true
    );
  }

  /** The value to write for a property this proxy did not set: the loaded one, or empty. */
  #current(name: 'address' | 'screenTip'): string {
    return (
      this.#written.get(name) ??
      (this.hasLoadedProperty(name) ? this.loadedProperty<string>(name) : '')
    );
  }

  #spanNow(): AutomationTextSpan | null | undefined {
    return typeof this.#span === 'function' ? this.#span() : this.#span;
  }

  #op(ref: RefOf, address: string | null, screenTip: string | undefined): AutomationOperation {
    // `ref` first: for a lookup in this batch it waits for the read, which gives the span.
    const shape = ref(this.#shape);
    const tip = screenTip ? { screenTip } : {};
    const span = this.#spanNow();
    if (span === null) return { op: 'setShapeHyperlink', shape, address, ...tip };
    return { op: 'setHyperlink', shape, ...(span ? { span } : {}), address, ...tip };
  }

  /** Write one property. The other keeps its value, read when the batch is planned. */
  #write(name: 'address' | 'screenTip', value: string): void {
    this.#written.set(name, value);
    this.setLoadedProperty(name, value);
    this.command(name, (ref) =>
      this.#op(ref, this.#current('address'), this.#current('screenTip'))
    );
  }

  /** @internal Values come with the object; a load has nothing more to read. */
  protected override onLoad(request: ResolvedLoadOptions): void {
    this.selection(request, ['address', 'screenTip', 'type']);
  }
}

/** A list of hyperlinks that one read answers: of a slide, or of a text range. */
abstract class HyperlinkList extends ModelObject {
  /** @internal The read that lists the hyperlinks. */
  protected abstract listing(ref: RefOf): AutomationOperation;

  /** The loaded hyperlinks. */
  get items(): Hyperlink[] {
    return this.loadedProperty<Hyperlink[]>('items');
  }

  /** The number of hyperlinks, after the next `sync()`. */
  getCount(): ClientResult<number> {
    const label = `${this.path.label}.getCount`;
    const { result, fill } = ClientResult.create<number>(label);
    this.enqueue(
      'read',
      'getCount',
      (ref) => this.listing(ref),
      (value) => fill(hydrated(value, 'hyperlinks', label).hyperlinks.length)
    );
    return result;
  }

  /** The hyperlink at a zero-based position; `ItemNotFound` at the sync when there is none. */
  getItemAt(index: number): Hyperlink {
    if (typeof index !== 'number' || !Number.isInteger(index) || index < 0)
      failArgument(`${this.path.label}.getItemAt`, 'a whole number of 0 or more');
    const path = ObjectPath.pending(`${this.path.label}.getItemAt(${index})`);
    const link = new Hyperlink(this.context, path, path, undefined);
    const label = path.label;
    this.enqueue(
      'read',
      `getItemAt(${index})`,
      (ref) => this.listing(ref),
      (value) => {
        const found = hydrated(value, 'hyperlinks', label).hyperlinks[index];
        if (!found) fail({ code: 'ItemNotFound', target: label });
        FILLERS.get(link)?.(found);
        path.resolveTo(found.shape);
      },
      path,
      false
    );
    return link;
  }

  /** @internal */
  protected override get loadsItemPaths(): boolean {
    return true;
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    for (const name of [...request.select, ...request.itemSelect])
      if (!['items', 'address', 'screenTip', 'type'].includes(name))
        failArgument(
          `${this.path.label}.${name}`,
          'a property of hyperlinks: items, address, screenTip, or type',
          nearest(name, ['items', 'address', 'screenTip', 'type'])
        );
    const label = `${this.path.label}.items`;
    this.enqueue(
      'read',
      'items',
      (ref) => this.listing(ref),
      (value) => {
        const reads = hydrated(value, 'hyperlinks', label).hyperlinks;
        this.setLoadedProperty(
          'items',
          reads.map((read, i) => {
            const shape = ObjectPath.of(`${label}[${i}]`, read.shape);
            return new Hyperlink(this.context, shape, shape, spanOf(read), {
              address: read.address ?? '',
              screenTip: read.screenTip ?? '',
            });
          })
        );
      }
    );
  }
}

/**
 * The hyperlinks of a slide, in shape order.
 *
 * @public
 */
export class HyperlinkCollection extends HyperlinkList {
  readonly #slide: ObjectPath;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, slide: ObjectPath) {
    super(context, path);
    this.#slide = slide;
  }

  /** Link a text range or a whole shape. Answers the new hyperlink. */
  add(target: TextRange | Shape, options?: HyperlinkAddOptions): Hyperlink {
    if (target instanceof TextRange || target instanceof Shape) return target.setHyperlink(options);
    failArgument(`${this.path.label}.add`, 'a TextRange or a Shape');
  }

  /** @internal */
  protected override listing(ref: RefOf): AutomationOperation {
    return { op: 'listHyperlinks', slide: ref(this.#slide) };
  }
}

/**
 * The hyperlinks of a text range: each text link that overlaps the range, in text order. A link
 * keeps its whole span, so one that starts before the range reads from its own start. An empty
 * range holds the link around its point.
 *
 * @public
 */
export class HyperlinkScopedCollection extends HyperlinkList {
  readonly #shape: ObjectPath;
  readonly #span: () => AutomationTextSpan | undefined;

  /** @internal */
  constructor(
    context: RequestContext,
    path: ObjectPath,
    shape: ObjectPath,
    span: () => AutomationTextSpan | undefined
  ) {
    super(context, path);
    this.#shape = shape;
    this.#span = span;
  }

  /** @internal */
  protected override listing(ref: RefOf): AutomationOperation {
    // `ref` first: for a range a read finds, it waits for that read, which gives the span.
    const shape = ref(this.#shape);
    const span = this.#span();
    return { op: 'listTextHyperlinks', shape, ...(span ? { span } : {}) };
  }
}

/** The span of a text link, or null for a shape link. */
function spanOf(read: AutomationHyperlinkRead | undefined): AutomationTextSpan | null {
  if (!read || read.type === 'shape' || read.start === null || read.length === null) return null;
  return { start: read.start, length: read.length };
}
