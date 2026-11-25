/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `Section` and `SectionCollection`: SDK extensions over the sections of a presentation. The
// pinned Office.js reference has no sections. A section is a `p14:section` of the section list
// in the presentation's extension list (`p:extLst`, `CT_ExtensionList`, ECMA-376 Part 1,
// 19.2.1.12). The section elements come from the `p14` extension schema, not from ECMA-376.

import type { AutomationItemKey, AutomationOperation } from '@pptx-editor.dev/core/automation';
import { failArgument, nearest } from '../runtime/errors.ts';
import { hydrated } from '../runtime/hydrate.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import { ItemCollection } from './collection.ts';
import { ModelObject, pathOf, text, type RefOf } from './model-object.ts';
import { Slide, type SlideCollection } from './slides.ts';

/**
 * Options for `Section.delete`. An SDK extension.
 *
 * @public
 */
export interface SectionDeleteOptions {
  /**
   * Delete the slides of the section too. When false or absent, the slides stay and join the
   * section before it, or the section after it for the first section.
   */
  deleteSlides?: boolean;
}

/**
 * A section of the presentation: a named run of slides. An SDK extension: the pinned Office.js
 * reference has no sections.
 *
 * @example
 * ```ts
 * const section = context.presentation.sections.getItemAt(0).load('name,slideIds');
 * await context.sync();
 * console.log(section.name, section.slideIds);
 * ```
 *
 * @public
 */
export class Section extends ModelObject {
  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, nullable = false) {
    super(context, path, { nullable });
  }

  /**
   * The section id (`p14:section/@id`): a GUID in braces. Stable across moves and renames. A
   * section of a file that has no id reads as an empty string; find it with `getItemAt`.
   */
  get id(): string {
    return this.loadedProperty('id');
  }

  /**
   * The name of the section. A write cleans the name first: it removes control characters,
   * turns a line break into a space, and trims spaces at the ends. A cleaned name of 1 to 255
   * UTF-16 code units applies. An empty or longer name refuses with `InvalidArgument` and
   * diagnostic `edit.section-name` at the sync. A name that cleans to the current name writes
   * nothing.
   */
  get name(): string {
    return this.loadedProperty('name');
  }
  set name(value: string) {
    const name = text(value, this.at('name'));
    this.command('name', (ref) => ({ op: 'renameSection', section: ref(this.path), name }));
  }

  /** The zero-based position of the section in `presentation.sections`. */
  get index(): number {
    return this.loadedProperty('index');
  }

  /**
   * The ids of the slides of the section (`Slide.id`), in slide order. Empty for no slides. Pass
   * an id to `presentation.slides.getItem` to work on that slide.
   *
   * @example
   * ```ts
   * const section = context.presentation.sections.getItemAt(1).load('slideIds');
   * await context.sync();
   * for (const id of section.slideIds) context.presentation.slides.getItem(id).hidden = true;
   * await context.sync();
   * ```
   */
  get slideIds(): string[] {
    return this.loadedProperty('slideIds');
  }

  /**
   * Move the section with its slides to a zero-based position among the sections. The slides
   * keep their order. A position past the last section refuses with `InvalidArgument` at the
   * sync.
   */
  moveTo(index: number): void {
    if (typeof index !== 'number' || !Number.isInteger(index) || index < 0)
      failArgument(this.at('moveTo'), 'a whole number of 0 or more');
    this.command('moveTo', (ref) => ({ op: 'moveSection', section: ref(this.path), index }));
  }

  /**
   * Delete the section. Its slides stay and join the section before it, or the section after it
   * for the first section. With `deleteSlides: true`, the slides are deleted too. A section that
   * holds every slide then refuses with `NotSupported`, because a presentation keeps one slide.
   * Deleting the last section removes the section list.
   */
  delete(options?: SectionDeleteOptions): void {
    const target = this.at('delete');
    if (options !== undefined && (typeof options !== 'object' || options === null))
      failArgument(target, 'a SectionDeleteOptions object');
    for (const key of Object.keys(options ?? {}))
      if (key !== 'deleteSlides')
        failArgument(`${target}.${key}`, "'deleteSlides'", nearest(key, ['deleteSlides']));
    const deleteSlides = options?.deleteSlides;
    if (deleteSlides !== undefined && typeof deleteSlides !== 'boolean')
      failArgument(`${target}.deleteSlides`, 'true or false');
    this.command('delete', (ref) => ({
      op: 'deleteSection',
      section: ref(this.path),
      ...(deleteSlides ? { deleteSlides: true } : {}),
    }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const names = this.selection(request, ['id', 'name', 'index', 'slideIds']);
    this.loadScalars(
      names,
      'section',
      (ref) => ({ op: 'readSection', section: ref(this.path) }),
      (v) => ({
        id: v.section.id,
        name: v.section.name,
        index: v.section.index,
        slideIds: [...v.section.slideIds],
      })
    );
  }
}

/**
 * The sections of the presentation, in order. An SDK extension: the pinned Office.js reference
 * has no sections. A presentation without sections has an empty collection.
 *
 * Slide calls keep the sections in step. `SlideCollection.add` puts the new slide at the end of
 * the last section. `Slide.moveTo` puts the slide into the section of the slide before its new
 * place, or into the first section at position 0. `Slide.duplicate` puts the copy into the
 * section of its source. `Slide.delete` takes the slide out of its section, and the section
 * stays, also when it is empty. `Presentation.insertSlidesFromBase64` puts the new slides into
 * the section of `targetSlideId`, or into the first section without one.
 *
 * @example
 * ```ts
 * const sections = context.presentation.sections;
 * sections.add('Appendix', '260');
 * sections.load('items/name,items/slideIds');
 * await context.sync();
 * ```
 *
 * @public
 */
export class SectionCollection extends ItemCollection<Section> {
  readonly #slides: SlideCollection;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, slides: SlideCollection) {
    super(context, path);
    this.#slides = slides;
  }

  /**
   * Add a section that starts at a slide, by id or as a `Slide`, and answer it. The new section
   * takes that slide and the slides after it in its section. In a presentation without
   * sections, the slides before it go into a first section named "Default Section". A slide
   * that already starts a section leaves that section empty. The new section works in the same
   * batch, with no load first.
   */
  add(name: string, firstSlide: string | Slide): Section {
    const target = this.at('add');
    const label = text(name, `${target}.name`);
    const slide =
      typeof firstSlide === 'string'
        ? this.#slides.getItem(firstSlide)
        : firstSlide instanceof Slide
          ? firstSlide
          : failArgument(`${target}.firstSlide`, 'a slide id or a Slide');
    const slidePath = pathOf(slide);
    const path = ObjectPath.pending(`${this.path.label}.add()`);
    const section = new Section(this.context, path);
    this.enqueue(
      'write',
      'add',
      (ref) => ({ op: 'addSection', name: label, slide: ref(slidePath) }),
      (value) => path.resolveTo(hydrated(value, 'handle', path.label).handle),
      path
    );
    return section;
  }

  /** Delete every section. The slides stay. With no sections, nothing changes. */
  deleteAll(): void {
    this.command('deleteAll', () => ({ op: 'deleteAllSections' }));
  }

  /** A section by id; `ItemNotFound` at the sync when there is none. */
  getItem(key: string): Section {
    return this.find({ id: this.keyArg(key, 'getItem') }, 'getItem()', false);
  }

  /** A section by zero-based position; `ItemNotFound` at the sync when there is none. */
  getItemAt(index: number): Section {
    const at = this.indexArg(index);
    return this.find({ index: at }, `getItemAt(${at})`, false);
  }

  /** A section by id, or a null object. */
  getItemOrNullObject(id: string): Section {
    return this.find({ id: this.keyArg(id, 'getItemOrNullObject') }, 'getItemOrNullObject()', true);
  }

  /** @internal */
  protected listing(): AutomationOperation {
    return { op: 'listSections' };
  }

  /** @internal */
  protected itemOf(path: ObjectPath, nullable: boolean): Section {
    return new Section(this.context, path, nullable);
  }

  /** @internal */
  protected lookup(_ref: RefOf, key: AutomationItemKey, orNull: boolean): AutomationOperation {
    return { op: 'getSection', key, ...(orNull ? { orNull: true } : {}) };
  }
}
