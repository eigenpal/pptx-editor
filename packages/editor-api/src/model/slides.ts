/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Slides, layouts, and masters: Office.js `Slide`, `SlideCollection`, `SlideScopedCollection`,
// `SlideLayout`, `SlideLayoutCollection`, `SlideMaster`, and `SlideMasterCollection`.

import type {
  AutomationHandle,
  AutomationItemKey,
  AutomationOperation,
  AutomationValue,
} from '@pptx-editor.dev/core/automation';
import { ClientResult } from '../runtime/client-result.ts';
import { fail, failArgument } from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import { SlideBackground, SlideLayoutBackground, SlideMasterBackground } from './backgrounds.ts';
import { ItemCollection } from './collection.ts';
import { LAYOUT_TYPE, type SlideLayoutType } from './enums.ts';
import { HyperlinkCollection } from './hyperlinks.ts';
import { imageBox, settleImage, type SlideGetImageOptions } from './images.ts';
import { flag, ModelObject, pathOf, text, type RefOf } from './model-object.ts';
import { replaceOperation, type SearchOptions } from './search.ts';
import { ShapeCollection } from './shapes.ts';
import { CustomXmlPartCollection } from './custom-xml.ts';
import { TagCollection } from './tags.ts';
import { ThemeColorScheme } from './theme.ts';

/**
 * Options for `SlideCollection.add`.
 *
 * @public
 */
export interface AddSlideOptions {
  /** The layout of the new slide, by id. The first layout of the master when absent. */
  layoutId?: string;
  /** The master of the new slide, by id. The master of the last slide when absent. */
  slideMasterId?: string;
}

/**
 * A slide.
 *
 * @public
 */
export class Slide extends ModelObject {
  #shapes: ShapeCollection | undefined;
  #tags: TagCollection | undefined;
  #customXml: CustomXmlPartCollection | undefined;
  #hyperlinks: HyperlinkCollection | undefined;
  #layout: SlideLayout | undefined;
  #master: SlideMaster | undefined;
  #background: SlideBackground | undefined;
  #theme: ThemeColorScheme | undefined;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, nullable = false) {
    super(context, path, { nullable });
  }

  /** The slide id (`p:sldId/@id`). Stable across moves. */
  get id(): string {
    return this.loadedProperty('id');
  }

  /** The zero-based position of the slide. */
  get index(): number {
    return this.loadedProperty('index');
  }

  /**
   * Whether the slide is hidden: `p:sld/@show` is false. The slide element is ECMA-376 Part 1,
   * §19.3.1.38, and the schema declares `CT_Slide/@show` as an `xsd:boolean` with default
   * `true`. The editor's slide show skips a hidden slide. A write of `true` saves `show="0"`, and
   * a write of `false` removes the attribute. A write of the current value writes nothing.
   *
   * This property is an SDK extension. The pinned Office.js reference has no such member.
   *
   * A hidden slide stays in `presentation.slides`, and `getImageAsBase64()` draws it. A PDF
   * from `pptx-to-pdf` leaves it out unless the export sets `includeHiddenSlides`.
   *
   * @example
   * ```ts
   * const slide = context.presentation.slides.getItemAt(1).load('hidden');
   * await context.sync();
   * slide.hidden = !slide.hidden;
   * await context.sync();
   * ```
   */
  get hidden(): boolean {
    return this.loadedProperty('hidden');
  }
  set hidden(value: boolean) {
    const hidden = flag(value, this.at('hidden'));
    this.command('hidden', (ref) => ({ op: 'setSlideHidden', slide: ref(this.path), hidden }));
  }

  /**
   * The speaker notes of the slide as plain text. The notes are the text of the body
   * placeholder (`p:ph` with `type="body"`, ECMA-376 Part 1, 19.3.1.36) of the slide's notes
   * slide (`p:notes`, `CT_NotesSlide`, 19.3.1.26). A read joins the paragraphs (`a:p`,
   * 21.1.2.2.6) with `\n`, and a line break (`a:br`, 21.1.2.2.1) also reads as `\n`. A slide
   * with no notes slide reads `""`.
   *
   * A write makes one paragraph for each line: `\n`, `\r\n`, `\r`, and `\v` each end a line.
   * It rewrites only the text that differs, so unchanged runs keep their formatting, and new
   * text takes the formatting of the place where it goes. A write of the current text writes
   * nothing, and a line break that the write does not change stays a line break. A write of
   * text to a slide with no notes slide creates the notes slide (13.3.5) from the notes master
   * (13.3.4); a write of `""` there creates nothing.
   *
   * A value that is not a string, longer than 1,000,000 characters, or with a character that
   * XML cannot hold refuses with `InvalidArgument`, and nothing changes. A write that must create
   * a notes slide refuses with `NotSupported` and the diagnostic `edit.no-notes-master` when the
   * presentation has no notes master. A notes slide with no body placeholder refuses a write with `NotSupported` and the diagnostic
   * `edit.no-notes-placeholder`.
   *
   * The notes are a part of their own, so a read of them is one more operation for each slide.
   * A load of every scalar property (`load()`) includes them.
   *
   * This property is an SDK extension. The pinned Office.js reference has no such member.
   * Unlike `TextRange.text`, which joins paragraphs with `\r`, the notes text uses `\n`.
   *
   * @example
   * ```ts
   * const slide = context.presentation.slides.getItemAt(0).load('speakerNotes');
   * await context.sync();
   * slide.speakerNotes = `${slide.speakerNotes}\nSource: 2026 annual report`;
   * await context.sync();
   * ```
   */
  get speakerNotes(): string {
    return this.loadedProperty('speakerNotes');
  }
  set speakerNotes(value: string) {
    const notes = text(value, this.at('speakerNotes'));
    this.command('speakerNotes', (ref) => ({
      op: 'setSlideNotes',
      slide: ref(this.path),
      text: notes,
    }));
  }

  /** The shapes of the slide, in z-order. */
  get shapes(): ShapeCollection {
    return (this.#shapes ??= new ShapeCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.shapes`, this.path),
      this.path
    ));
  }

  /** The tags of the slide. */
  get tags(): TagCollection {
    return (this.#tags ??= new TagCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.tags`, this.path)
    ));
  }

  /** The custom XML parts of the slide. */
  get customXmlParts(): CustomXmlPartCollection {
    return (this.#customXml ??= new CustomXmlPartCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.customXmlParts`, this.path)
    ));
  }

  /** The hyperlinks of the slide: text links and shape click links. */
  get hyperlinks(): HyperlinkCollection {
    return (this.#hyperlinks ??= new HyperlinkCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.hyperlinks`, this.path),
      this.path
    ));
  }

  /** The background of the slide. */
  get background(): SlideBackground {
    return (this.#background ??= new SlideBackground(
      this.context,
      ObjectPath.derived(`${this.path.label}.background`, this.path),
      this.path
    ));
  }

  /** The theme colors of the slide: those of the theme of its master. */
  get themeColorScheme(): ThemeColorScheme {
    return (this.#theme ??= new ThemeColorScheme(
      this.context,
      ObjectPath.derived(`${this.path.label}.themeColorScheme`, this.path),
      this.path
    ));
  }

  /** The layout of the slide. */
  get layout(): SlideLayout {
    return (this.#layout ??= this.navigation(
      'layout',
      (path) => new SlideLayout(this.context, path),
      (ref) => ({ op: 'getSlideLayout', slide: ref(this.path) })
    ));
  }

  /** The master of the slide. */
  get slideMaster(): SlideMaster {
    return (this.#master ??= this.navigation(
      'slideMaster',
      (path) => new SlideMaster(this.context, path),
      (ref) => ({ op: 'getSlideMaster', owner: ref(this.path) })
    ));
  }

  /** Change the layout, re-mapping placeholders by type and index. */
  applyLayout(slideLayout: SlideLayout): void {
    const layout = pathOf(slideLayout);
    this.command('applyLayout', (ref) => ({
      op: 'applyLayout',
      slide: ref(this.path),
      layout: ref(layout),
    }));
  }

  /** Delete the slide. */
  delete(): void {
    this.command('delete', (ref) => ({ op: 'deleteSlide', slide: ref(this.path) }));
  }

  /**
   * SDK extension: copy the slide inside this presentation and answer the copy. The copy goes
   * directly after this slide, in its section. It keeps the layout, the master, and the
   * pictures of this slide, and gets its own notes, charts, diagrams, comments, embedded
   * objects, and tags, so an edit to one slide never changes the other. A slide that holds
   * embedded controls refuses with `NotSupported`.
   *
   * The copy is usable at once: a load, a write, or a call on it in the same batch applies to
   * the new slide. Each `sync()` that writes is one undo step in a live editor.
   *
   * The copy is a new `p:sldId` (`CT_SlideIdListEntry`) in `p:sldIdLst` (`CT_SlideIdList`),
   * with a new Slide part: ECMA-376 Part 1, 19.2.1.33, 19.2.1.34, and 13.3.8.
   *
   * @example
   * ```ts
   * const copy = context.presentation.slides.getItemAt(0).duplicate();
   * copy.load('id,index');
   * await context.sync();
   * console.log(copy.id, copy.index); // a new id, at index 1
   * ```
   */
  duplicate(): Slide {
    const path = ObjectPath.pending(`${this.path.label}.duplicate()`);
    const copy = new Slide(this.context, path);
    const label = path.label;
    this.enqueue(
      'write',
      'duplicate',
      (ref) => ({ op: 'duplicateSlide', slide: ref(this.path) }),
      (value) => path.resolveTo(hydrated(value, 'handle', label).handle),
      path
    );
    return copy;
  }

  /** Move the slide to a zero-based position. */
  moveTo(slideIndex: number): void {
    if (typeof slideIndex !== 'number' || !Number.isInteger(slideIndex) || slideIndex < 0)
      failArgument(`${this.path.label}.moveTo`, 'a whole number of 0 or more');
    this.command('moveTo', (ref) => ({
      op: 'moveSlide',
      slide: ref(this.path),
      index: slideIndex,
    }));
  }

  /**
   * A PNG picture of the slide as base64, drawn by the editor's own renderer. A runtime over a
   * live editor in a browser draws pictures, and so does a server runtime created with a
   * `rasterize` option; any other runtime refuses with `NotSupported`. A hidden slide
   * (`hidden`) draws as any other slide.
   *
   * @example
   * ```ts
   * const image = slide.getImageAsBase64({ width: 640 });
   * await context.sync();
   * const png = `data:image/png;base64,${image.value}`;
   * ```
   */
  getImageAsBase64(options?: SlideGetImageOptions): ClientResult<string> {
    const target = this.at('getImageAsBase64');
    const box = imageBox(options, target, false);
    const { result, fill } = ClientResult.create<string>(target);
    this.enqueue(
      'read',
      'getImageAsBase64',
      (ref) => ({ op: 'prepareImage', target: ref(this.path), ...box }),
      settleImage(this.internals, box, target, fill)
    );
    return result;
  }

  /** The slide as a presentation of its own, Base64-encoded PPTX bytes. */
  exportAsBase64(): ClientResult<string> {
    const label = `${this.path.label}.exportAsBase64`;
    const { result, fill } = ClientResult.create<string>(label);
    this.enqueue(
      'read',
      'exportAsBase64',
      (ref) => ({ op: 'exportSlides', slides: [ref(this.path)] }),
      (value) => fill(hydrated(value, 'base64', label).base64)
    );
    return result;
  }

  /**
   * Replace each match of `searchText` on this slide with `replacementText`, and answer the
   * number of matches after the next `sync()`. An SDK extension: Office.js has no such member.
   * The search reads the shapes of the slide (also in groups), the table cells the slide shows,
   * and the speaker notes of the slide. The text rules and refusals are those of
   * `Presentation.replaceAll`.
   *
   * @example
   * ```ts
   * const slide = context.presentation.slides.getItemAt(0);
   * const count = slide.replaceAll('{{client}}', 'Northwind', { matchCase: true });
   * await context.sync();
   * ```
   */
  replaceAll(
    searchText: string,
    replacementText: string,
    options?: SearchOptions
  ): ClientResult<number> {
    const target = this.at('replaceAll');
    const op = replaceOperation(searchText, replacementText, options, target);
    const { result, fill } = ClientResult.create<number>(target);
    this.enqueue(
      'write',
      'replaceAll',
      (ref) => ({ ...op, slide: ref(this.path) }),
      (value) => fill(hydrated(value, 'number', target).value)
    );
    return result;
  }

  /** Select shapes of this slide by id. A server runtime refuses with `NotSupported`. */
  setSelectedShapes(shapeIds: string[]): void {
    if (!Array.isArray(shapeIds)) failArgument('setSelectedShapes', 'an array of shape ids');
    const shapes = shapeIds.map((id) =>
      this.shapes.getItem(text(id, this.at('setSelectedShapes')))
    );
    this.command('setSelectedShapes', (ref) => ({
      op: 'selectShapes',
      shapes: shapes.map((shape) => ref(pathOf(shape))),
    }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, ['id', 'index', 'hidden', 'speakerNotes'], {
      shapes: () => this.shapes,
      tags: () => this.tags,
      hyperlinks: () => this.hyperlinks,
      layout: () => this.layout,
      slideMaster: () => this.slideMaster,
      background: () => this.background,
      themeColorScheme: () => this.themeColorScheme,
      customXmlParts: () => this.customXmlParts,
    });
    this.loadScalars(
      names.filter((name) => name !== 'speakerNotes'),
      'slide',
      (ref) => ({ op: 'readSlide', slide: ref(this.path) }),
      (v) => ({ id: v.slide.id, index: v.slide.index, hidden: v.slide.hidden })
    );
    // The notes live in a part of their own, which a read of its own answers.
    this.loadScalars(
      names.filter((name) => name === 'speakerNotes'),
      'text',
      (ref) => ({ op: 'readSlideNotes', slide: ref(this.path) }),
      (v) => ({ speakerNotes: v.text })
    );
  }
}

/**
 * The slides of the presentation, in order.
 *
 * @public
 */
export class SlideCollection extends ItemCollection<Slide> {
  /** @internal */
  constructor(context: RequestContext, path: ObjectPath) {
    super(context, path);
  }

  /** Add a slide at the end, from a layout and master by id. */
  add(options?: AddSlideOptions): void {
    const layoutId = options?.layoutId;
    const masterId = options?.slideMasterId;
    if (options !== undefined && (typeof options !== 'object' || options === null))
      failArgument('presentation.slides.add', 'an AddSlideOptions object');
    for (const key of Object.keys(options ?? {}))
      if (key !== 'layoutId' && key !== 'slideMasterId')
        fail({ code: 'NotImplemented', target: `presentation.slides.add.${key}` });
    // The host answers the new slide; the object model's `add` answers nothing.
    this.enqueue('write', 'add', () => ({
      op: 'addSlide',
      ...(layoutId === undefined ? {} : { layoutId: text(layoutId, this.at('layoutId')) }),
      ...(masterId === undefined ? {} : { masterId: text(masterId, this.at('slideMasterId')) }),
    }));
  }

  /** The listed slides as a presentation of their own, Base64-encoded PPTX bytes. */
  exportAsBase64Presentation(values: Array<string | Slide>): ClientResult<string> {
    if (!Array.isArray(values) || values.length === 0)
      failArgument(
        'presentation.slides.exportAsBase64Presentation',
        'a non-empty array of slides or slide ids'
      );
    const slides = values.map((v) => (typeof v === 'string' ? this.getItem(v) : v));
    const paths = slides.map((slide) => pathOf(slide));
    const label = `${this.path.label}.exportAsBase64Presentation`;
    const { result, fill } = ClientResult.create<string>(label);
    this.enqueue(
      'read',
      'exportAsBase64Presentation',
      (ref) => ({ op: 'exportSlides', slides: paths.map((p) => ref(p)) }),
      (value) => fill(hydrated(value, 'base64', label).base64)
    );
    return result;
  }

  /** A slide by id; `ItemNotFound` at the sync when there is none. */
  getItem(key: string): Slide {
    return this.find({ id: this.keyArg(key, 'getItem') }, 'getItem()', false);
  }

  /** A slide by zero-based position; `ItemNotFound` at the sync when there is none. */
  getItemAt(index: number): Slide {
    const at = this.indexArg(index);
    return this.find({ index: at }, `getItemAt(${at})`, false);
  }

  /** A slide by id, or a null object. */
  getItemOrNullObject(id: string): Slide {
    return this.find({ id: this.keyArg(id, 'getItemOrNullObject') }, 'getItemOrNullObject()', true);
  }

  /** @internal */
  protected listing(): AutomationOperation {
    return { op: 'listSlides' };
  }

  /** @internal */
  protected itemOf(path: ObjectPath, nullable: boolean): Slide {
    return new Slide(this.context, path, nullable);
  }

  /** @internal */
  protected lookup(_ref: RefOf, key: AutomationItemKey, orNull: boolean): AutomationOperation {
    return { op: 'getSlide', key, ...(orNull ? { orNull: true } : {}) };
  }
}

/**
 * A set of slides: the reader's selection.
 *
 * @public
 */
export class SlideScopedCollection extends ItemCollection<Slide> {
  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, source: 'selection') {
    super(context, path);
    void source;
  }

  /** The slides as a presentation of their own, Base64-encoded PPTX bytes. */
  exportAsBase64Presentation(): ClientResult<string> {
    const label = `${this.path.label}.exportAsBase64Presentation`;
    const { result, fill } = ClientResult.create<string>(label);
    const listing = ObjectPath.pending(`${this.path.label}.items`);
    let slides: readonly AutomationHandle[] = [];
    this.enqueue(
      'read',
      'items',
      () => this.listing(),
      (value) => {
        slides = this.handlesOf(value, label);
        if (slides[0]) listing.resolveTo(slides[0]);
        else listing.resolveNull();
      },
      listing,
      false
    );
    this.enqueue(
      'read',
      'exportAsBase64Presentation',
      (ref) => {
        if (listing.isNull) fail({ code: 'ItemNotFound', target: label });
        ref(listing);
        return { op: 'exportSlides', slides };
      },
      (value) => fill(hydrated(value, 'base64', label).base64)
    );
    return result;
  }

  /** A selected slide by id; `ItemNotFound` at the sync when there is none. */
  getItem(key: string): Slide {
    return this.find({ id: this.keyArg(key, 'getItem') }, 'getItem()', false);
  }

  /** A selected slide by position in the selection. */
  getItemAt(index: number): Slide {
    const at = this.indexArg(index);
    return this.find({ index: at }, `getItemAt(${at})`, false);
  }

  /** A selected slide by id, or a null object. */
  getItemOrNullObject(id: string): Slide {
    return this.find({ id: this.keyArg(id, 'getItemOrNullObject') }, 'getItemOrNullObject()', true);
  }

  /** @internal */
  protected listing(): AutomationOperation {
    return { op: 'readSelection' };
  }

  /** @internal */
  protected override handlesOf(value: AutomationValue, label: string): readonly AutomationHandle[] {
    return hydrated(value, 'selection', label).selection.slides;
  }

  /** @internal */
  protected override get hasLookup(): boolean {
    return false;
  }

  /** @internal */
  protected itemOf(path: ObjectPath, nullable: boolean): Slide {
    return new Slide(this.context, path, nullable);
  }

  /** @internal */
  protected lookup(): undefined {
    return undefined;
  }
}

const LAYOUT_NAMES = ['id', 'name', 'type'] as const;

/**
 * A slide layout (`p:sldLayout`).
 *
 * @public
 */
export class SlideLayout extends ModelObject {
  #shapes: ShapeCollection | undefined;
  #background: SlideLayoutBackground | undefined;
  #theme: ThemeColorScheme | undefined;
  #customXml: CustomXmlPartCollection | undefined;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, nullable = false) {
    super(context, path, { nullable });
  }

  /** The layout id (`p:sldLayoutId/@id`). */
  get id(): string {
    return this.loadedProperty('id');
  }
  /** The layout name (`p:cSld/@name`). */
  get name(): string {
    return this.loadedProperty('name');
  }
  /** The layout type (`p:sldLayout/@type`). */
  get type(): SlideLayoutType | `${SlideLayoutType}` {
    return this.loadedProperty('type');
  }
  /** The shapes of the layout. */
  get shapes(): ShapeCollection {
    return (this.#shapes ??= new ShapeCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.shapes`, this.path),
      this.path
    ));
  }

  /** The background of the layout. */
  get background(): SlideLayoutBackground {
    return (this.#background ??= new SlideLayoutBackground(
      this.context,
      ObjectPath.derived(`${this.path.label}.background`, this.path),
      this.path
    ));
  }

  /** The custom XML parts of the layout. */
  get customXmlParts(): CustomXmlPartCollection {
    return (this.#customXml ??= new CustomXmlPartCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.customXmlParts`, this.path)
    ));
  }

  /** The theme colors of the layout: those of the theme of its master. */
  get themeColorScheme(): ThemeColorScheme {
    return (this.#theme ??= new ThemeColorScheme(
      this.context,
      ObjectPath.derived(`${this.path.label}.themeColorScheme`, this.path),
      this.path
    ));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, LAYOUT_NAMES, {
      shapes: () => this.shapes,
      background: () => this.background,
      themeColorScheme: () => this.themeColorScheme,
      customXmlParts: () => this.customXmlParts,
    });
    this.loadScalars(
      names,
      'layout',
      (ref) => ({ op: 'readLayout', layout: ref(this.path) }),
      (v) => ({
        id: v.layout.id,
        name: v.layout.name,
        type: LAYOUT_TYPE.fromFile(v.layout.type) ?? 'Custom',
      })
    );
  }
}

/**
 * The layouts of a slide master, in `p:sldLayoutIdLst` order.
 *
 * @public
 */
export class SlideLayoutCollection extends ItemCollection<SlideLayout> {
  readonly #master: ObjectPath;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, master: ObjectPath) {
    super(context, path);
    this.#master = master;
  }

  /** A layout by id; `ItemNotFound` at the sync when there is none. */
  getItem(key: string): SlideLayout {
    return this.find({ id: this.keyArg(key, 'getItem') }, 'getItem()', false);
  }
  /** A layout by zero-based position. */
  getItemAt(index: number): SlideLayout {
    const at = this.indexArg(index);
    return this.find({ index: at }, `getItemAt(${at})`, false);
  }
  /** A layout by id, or a null object. */
  getItemOrNullObject(id: string): SlideLayout {
    return this.find({ id: this.keyArg(id, 'getItemOrNullObject') }, 'getItemOrNullObject()', true);
  }

  /** @internal */
  protected listing(ref: RefOf): AutomationOperation {
    return { op: 'listLayouts', master: ref(this.#master) };
  }
  /** @internal */
  protected itemOf(path: ObjectPath, nullable: boolean): SlideLayout {
    return new SlideLayout(this.context, path, nullable);
  }
  /** @internal */
  protected lookup(ref: RefOf, key: AutomationItemKey, orNull: boolean): AutomationOperation {
    return { op: 'getLayout', master: ref(this.#master), key, ...(orNull ? { orNull: true } : {}) };
  }
}

/**
 * A slide master (`p:sldMaster`).
 *
 * @public
 */
export class SlideMaster extends ModelObject {
  #layouts: SlideLayoutCollection | undefined;
  #shapes: ShapeCollection | undefined;
  #background: SlideMasterBackground | undefined;
  #theme: ThemeColorScheme | undefined;
  #customXml: CustomXmlPartCollection | undefined;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, nullable = false) {
    super(context, path, { nullable });
  }

  /** The master id (`p:sldMasterId/@id`). */
  get id(): string {
    return this.loadedProperty('id');
  }
  /** The master name (`p:cSld/@name`). */
  get name(): string {
    return this.loadedProperty('name');
  }
  /** The layouts of the master. */
  get layouts(): SlideLayoutCollection {
    return (this.#layouts ??= new SlideLayoutCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.layouts`, this.path),
      this.path
    ));
  }
  /** The shapes of the master. */
  get shapes(): ShapeCollection {
    return (this.#shapes ??= new ShapeCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.shapes`, this.path),
      this.path
    ));
  }

  /** The background of the master. */
  get background(): SlideMasterBackground {
    return (this.#background ??= new SlideMasterBackground(
      this.context,
      ObjectPath.derived(`${this.path.label}.background`, this.path),
      this.path
    ));
  }

  /** The custom XML parts of the master. */
  get customXmlParts(): CustomXmlPartCollection {
    return (this.#customXml ??= new CustomXmlPartCollection(
      this.context,
      ObjectPath.derived(`${this.path.label}.customXmlParts`, this.path)
    ));
  }

  /** The theme colors of the master: those of the theme of its master. */
  get themeColorScheme(): ThemeColorScheme {
    return (this.#theme ??= new ThemeColorScheme(
      this.context,
      ObjectPath.derived(`${this.path.label}.themeColorScheme`, this.path),
      this.path
    ));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, ['id', 'name'], {
      layouts: () => this.layouts,
      shapes: () => this.shapes,
      background: () => this.background,
      themeColorScheme: () => this.themeColorScheme,
      customXmlParts: () => this.customXmlParts,
    });
    this.loadScalars(
      names,
      'layout',
      (ref) => ({ op: 'readMaster', master: ref(this.path) }),
      (v) => ({ id: v.layout.id, name: v.layout.name })
    );
  }
}

/**
 * The slide masters of the presentation, in `p:sldMasterIdLst` order.
 *
 * @public
 */
export class SlideMasterCollection extends ItemCollection<SlideMaster> {
  /** @internal */
  constructor(context: RequestContext, path: ObjectPath) {
    super(context, path);
  }

  /** A master by id; `ItemNotFound` at the sync when there is none. */
  getItem(key: string): SlideMaster {
    return this.find({ id: this.keyArg(key, 'getItem') }, 'getItem()', false);
  }
  /** A master by zero-based position. */
  getItemAt(index: number): SlideMaster {
    const at = this.indexArg(index);
    return this.find({ index: at }, `getItemAt(${at})`, false);
  }
  /** A master by id, or a null object. */
  getItemOrNullObject(id: string): SlideMaster {
    return this.find({ id: this.keyArg(id, 'getItemOrNullObject') }, 'getItemOrNullObject()', true);
  }

  /** @internal */
  protected listing(): AutomationOperation {
    return { op: 'listMasters' };
  }
  /** @internal */
  protected itemOf(path: ObjectPath, nullable: boolean): SlideMaster {
    return new SlideMaster(this.context, path, nullable);
  }
  /** @internal */
  protected lookup(_ref: RefOf, key: AutomationItemKey, orNull: boolean): AutomationOperation {
    return { op: 'getMaster', key, ...(orNull ? { orNull: true } : {}) };
  }
}
