/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The root of the object model: Office.js `PowerPoint.Presentation`, `DocumentProperties`, and
// `PageSetup`.

import type { AutomationInsertSlidesOptions } from '@pptx-editor.dev/core/automation';
import { failArgument, nearest, oneOf } from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import { BindingCollection } from './bindings.ts';
import { InsertSlideFormatting } from './enums.ts';
import { CustomPropertyCollection } from './custom-properties.ts';
import { CustomXmlPartCollection } from './custom-xml.ts';
import { finite, ModelObject, pathOf, text } from './model-object.ts';
import { ShapeScopedCollection } from './shapes.ts';
import {
  SlideCollection,
  SlideMasterCollection,
  SlideScopedCollection,
  type Slide,
} from './slides.ts';
import { TagCollection } from './tags.ts';
import { TextRange } from './text.ts';

/**
 * Options for `Presentation.insertSlidesFromBase64`.
 *
 * @public
 */
export interface InsertSlideOptions {
  /** How the inserted slides are formatted. */
  formatting?: InsertSlideFormatting | `${InsertSlideFormatting}` | undefined;
  /** The slides of the source presentation to insert, by id. */
  sourceSlideIds?: string[];
  /** The slide the new slides follow, by id. */
  targetSlideId?: string;
}

/**
 * The presentation: its slides, masters, tags, properties, and page setup.
 *
 * @example
 * ```ts
 * await runtime.run(async (context) => {
 *   const slides = context.presentation.slides;
 *   slides.load('items/id');
 *   await context.sync();
 *   console.log(slides.items.map((slide) => slide.id));
 * });
 * ```
 *
 * @public
 */
export class Presentation extends ModelObject {
  #slides: SlideCollection | undefined;
  #masters: SlideMasterCollection | undefined;
  #tags: TagCollection | undefined;
  #customXml: CustomXmlPartCollection | undefined;
  #properties: DocumentProperties | undefined;
  #pageSetup: PageSetup | undefined;
  #bindings: BindingCollection | undefined;

  /** @internal */
  private constructor(context: RequestContext, path: ObjectPath) {
    super(context, path);
  }

  /** @internal The presentation of a request context, with a lookup queued at first use. */
  static open(context: RequestContext): Presentation {
    const path = ObjectPath.pending('presentation');
    const presentation = new Presentation(context, path);
    presentation.enqueue(
      'read',
      '',
      () => ({ op: 'getPresentation' }),
      (value) => path.resolveTo(hydrated(value, 'handle', 'presentation').handle),
      path
    );
    return presentation;
  }

  /** The id of the presentation. */
  get id(): string {
    return this.loadedProperty('id');
  }

  /** The title of the presentation, from the core properties. */
  get title(): string {
    return this.loadedProperty('title');
  }

  /** The slides, in order. */
  get slides(): SlideCollection {
    return (this.#slides ??= new SlideCollection(
      this.context,
      ObjectPath.derived('presentation.slides', this.path)
    ));
  }

  /** The slide masters, in order. */
  get slideMasters(): SlideMasterCollection {
    return (this.#masters ??= new SlideMasterCollection(
      this.context,
      ObjectPath.derived('presentation.slideMasters', this.path)
    ));
  }

  /** The tags of the presentation. */
  get tags(): TagCollection {
    return (this.#tags ??= new TagCollection(
      this.context,
      ObjectPath.derived('presentation.tags', this.path)
    ));
  }

  /** The custom XML parts of the presentation. */
  get customXmlParts(): CustomXmlPartCollection {
    return (this.#customXml ??= new CustomXmlPartCollection(
      this.context,
      ObjectPath.derived('presentation.customXmlParts', this.path)
    ));
  }

  /**
   * The bindings of this runtime: ids that name shapes. They live in memory for the lifetime of
   * the runtime and are never saved with the presentation.
   */
  get bindings(): BindingCollection {
    return (this.#bindings ??= new BindingCollection(
      this.context,
      ObjectPath.derived('presentation.bindings', this.path)
    ));
  }

  /** The document properties (title, author, …). */
  get properties(): DocumentProperties {
    return (this.#properties ??= new DocumentProperties(
      this.context,
      ObjectPath.derived('presentation.properties', this.path)
    ));
  }

  /** The slide size. */
  get pageSetup(): PageSetup {
    return (this.#pageSetup ??= new PageSetup(
      this.context,
      ObjectPath.derived('presentation.pageSetup', this.path)
    ));
  }

  /** The selected slides of the reader. A server runtime refuses with `NotSupported`. */
  getSelectedSlides(): SlideScopedCollection {
    return new SlideScopedCollection(
      this.context,
      ObjectPath.derived('presentation.getSelectedSlides()', this.path),
      'selection'
    );
  }

  /** The selected shapes of the reader. A server runtime refuses with `NotSupported`. */
  getSelectedShapes(): ShapeScopedCollection {
    return new ShapeScopedCollection(
      this.context,
      ObjectPath.derived('presentation.getSelectedShapes()', this.path),
      { kind: 'selection' }
    );
  }

  /** The selected text. `ItemNotFound` at the sync when no text is selected. */
  getSelectedTextRange(): TextRange {
    return TextRange.selected(this, 'getSelectedTextRange()', false);
  }

  /** The selected text, or a null object when no text is selected. */
  getSelectedTextRangeOrNullObject(): TextRange {
    return TextRange.selected(this, 'getSelectedTextRangeOrNullObject()', true);
  }

  /** Select slides by id. A server runtime refuses with `NotSupported`. */
  setSelectedSlides(slideIds: string[]): void {
    if (!Array.isArray(slideIds)) failArgument('setSelectedSlides', 'an array of slide ids');
    const slides = slideIds.map((id) =>
      this.slides.getItem(text(id, this.at('setSelectedSlides')))
    );
    this.command('setSelectedSlides', (ref) => ({
      op: 'selectSlides',
      slides: slides.map((slide: Slide) => ref(pathOf(slide))),
    }));
  }

  /**
   * Insert the slides of another presentation, given as Base64-encoded PPTX bytes. Each slide
   * comes with its notes, pictures, media, charts, and other parts. With `KeepSourceFormatting`
   * (the default), its layout, master, and theme come too; with `UseDestinationTheme`, it goes on
   * the layout of this presentation with the same type. Without `targetSlideId`, the slides go
   * first. Comments come along, with their authors merged by name and initials. An unknown slide
   * id refuses with `ItemNotFound`; a slide with threaded comments or an embedded control refuses
   * with `NotSupported`, naming the slide ids. Nothing changes when the call is refused.
   *
   * @example
   * ```ts
   * await runtime.run(async (context) => {
   *   context.presentation.insertSlidesFromBase64(base64, {
   *     formatting: 'UseDestinationTheme',
   *     targetSlideId: '256',
   *   });
   *   await context.sync();
   * });
   * ```
   */
  insertSlidesFromBase64(base64File: string, options?: InsertSlideOptions): void {
    const target = this.at('insertSlidesFromBase64');
    const base64 = text(base64File, target);
    if (base64.length === 0) failArgument(target, 'the Base64 bytes of a PPTX file');
    const request = importOptions(options, target);
    this.command('insertSlidesFromBase64', () => ({
      op: 'insertSlides',
      base64,
      ...(request ? { options: request } : {}),
    }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, ['id', 'title'], {
      slides: () => this.slides,
      slideMasters: () => this.slideMasters,
      tags: () => this.tags,
      bindings: () => this.bindings,
      customXmlParts: () => this.customXmlParts,
      properties: () => this.properties,
      pageSetup: () => this.pageSetup,
    });
    this.loadScalars(
      names,
      'presentation',
      () => ({ op: 'readPresentation' }),
      (v) => ({
        id: v.presentation.id,
        title: v.presentation.title,
      })
    );
  }
}

/** Import options as the protocol takes them, or `InvalidArgument` naming `target`. */
function importOptions(
  options: InsertSlideOptions | undefined,
  target: string
): AutomationInsertSlidesOptions | undefined {
  if (options === undefined) return undefined;
  if (typeof options !== 'object' || options === null)
    failArgument(target, 'an InsertSlideOptions object');
  const { formatting, sourceSlideIds, targetSlideId } = options;
  if (
    formatting !== undefined &&
    formatting !== InsertSlideFormatting.keepSourceFormatting &&
    formatting !== InsertSlideFormatting.useDestinationTheme
  )
    failArgument(
      `${target}.formatting`,
      oneOf(Object.values(InsertSlideFormatting)),
      nearest(formatting, Object.values(InsertSlideFormatting))
    );
  if (
    sourceSlideIds !== undefined &&
    (!Array.isArray(sourceSlideIds) || sourceSlideIds.some((id) => typeof id !== 'string'))
  )
    failArgument(`${target}.sourceSlideIds`, 'an array of slide id strings');
  if (targetSlideId !== undefined && typeof targetSlideId !== 'string')
    failArgument(`${target}.targetSlideId`, 'a slide id string');
  return {
    ...(formatting === undefined
      ? {}
      : {
          formatting:
            formatting === InsertSlideFormatting.useDestinationTheme
              ? 'useDestinationTheme'
              : 'keepSource',
        }),
    ...(sourceSlideIds === undefined ? {} : { sourceSlideIds: [...sourceSlideIds] }),
    ...(targetSlideId === undefined ? {} : { targetSlideId }),
  };
}

const PROPERTY_NAMES = [
  'author',
  'category',
  'comments',
  'company',
  'creationDate',
  'keywords',
  'lastAuthor',
  'manager',
  'revisionNumber',
  'subject',
  'title',
] as const;

/**
 * The core and extended document properties (`docProps/core.xml`, `docProps/app.xml`).
 *
 * @public
 */
export class DocumentProperties extends ModelObject {
  #custom: CustomPropertyCollection | undefined;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath) {
    super(context, path);
  }

  /** The custom document properties (`docProps/custom.xml`). */
  get customProperties(): CustomPropertyCollection {
    return (this.#custom ??= new CustomPropertyCollection(
      this.context,
      ObjectPath.derived('presentation.properties.customProperties', this.path)
    ));
  }

  get author(): string {
    return this.loadedProperty('author');
  }
  set author(value: string) {
    this.#set('author', text(value, this.at('author')));
  }
  get category(): string {
    return this.loadedProperty('category');
  }
  set category(value: string) {
    this.#set('category', text(value, this.at('category')));
  }
  get comments(): string {
    return this.loadedProperty('comments');
  }
  set comments(value: string) {
    this.#set('comments', text(value, this.at('comments')));
  }
  get company(): string {
    return this.loadedProperty('company');
  }
  set company(value: string) {
    this.#set('company', text(value, this.at('company')));
  }
  /**
   * When the presentation was created, from `dcterms:created`. Without one, or with a value
   * that is not a date, it is an invalid `Date`: check `Number.isNaN(creationDate.getTime())`.
   */
  get creationDate(): Date {
    return this.loadedProperty('creationDate');
  }
  get keywords(): string {
    return this.loadedProperty('keywords');
  }
  set keywords(value: string) {
    this.#set('keywords', text(value, this.at('keywords')));
  }
  get lastAuthor(): string {
    return this.loadedProperty('lastAuthor');
  }
  get manager(): string {
    return this.loadedProperty('manager');
  }
  set manager(value: string) {
    this.#set('manager', text(value, this.at('manager')));
  }
  get revisionNumber(): number {
    return this.loadedProperty('revisionNumber');
  }
  set revisionNumber(value: number) {
    this.#set('revisionNumber', finite(value, this.at('revisionNumber')));
  }
  get subject(): string {
    return this.loadedProperty('subject');
  }
  set subject(value: string) {
    this.#set('subject', text(value, this.at('subject')));
  }
  get title(): string {
    return this.loadedProperty('title');
  }
  set title(value: string) {
    this.#set('title', text(value, this.at('title')));
  }

  #set(name: string, value: string | number): void {
    this.command(name, () => ({ op: 'setProperties', properties: { [name]: value } }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, PROPERTY_NAMES, {
      customProperties: () => this.customProperties,
    });
    this.loadScalars(
      names,
      'properties',
      () => ({ op: 'getProperties' }),
      (v) => ({
        ...v.properties,
        // Office.js types creationDate as Date, never null. A file with no dcterms:created
        // reads an invalid Date (getTime() is NaN), not a made-up time.
        creationDate: new Date(v.properties.creationDate ?? Number.NaN),
      })
    );
  }
}

/**
 * The slide size, in points (`p:sldSz`, ECMA-376 Part 1, 19.2.1.39).
 *
 * @public
 */
export class PageSetup extends ModelObject {
  /** @internal */
  constructor(context: RequestContext, path: ObjectPath) {
    super(context, path);
  }

  get slideHeight(): number {
    return this.loadedProperty('slideHeight');
  }
  set slideHeight(value: number) {
    const slideHeight = finite(value, this.at('slideHeight'));
    this.command('slideHeight', () => ({ op: 'setPageSetup', slideHeight }));
  }
  get slideWidth(): number {
    return this.loadedProperty('slideWidth');
  }
  set slideWidth(value: number) {
    const slideWidth = finite(value, this.at('slideWidth'));
    this.command('slideWidth', () => ({ op: 'setPageSetup', slideWidth }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, ['slideHeight', 'slideWidth']);
    this.loadScalars(
      names,
      'presentation',
      () => ({ op: 'readPresentation' }),
      (v) => ({ slideHeight: v.presentation.slideHeight, slideWidth: v.presentation.slideWidth })
    );
  }
}
