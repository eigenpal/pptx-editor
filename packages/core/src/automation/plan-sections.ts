// Sections of the presentation, for `Presentation.sections` of the object model. The section
// list is a `p14:sectionLst` in a `p:ext` (`CT_Extension`, ECMA-376 Part 1, 19.2.1.11) of the
// presentation's extension list (`p:extLst`, `CT_ExtensionList`, 19.2.1.12; `CT_Presentation`,
// 19.2.1.26). Each `p14:section` names its slides by `p:sldId/@id` (`ST_SlideId`, 19.7.13). The
// section elements come from the `p14` extension schema, not from ECMA-376. A consumer passes an
// extension it does not know through unchanged (ECMA-376 Part 3, clause 8), so each edit here
// changes only the section list and leaves every other `p:ext` as it is.
//
// The edits are the engine's section commands (`commands/slide-sections.ts`), which the editor's
// section commands run too. Each runs on the repaired presentation part, as a slide-list edit does.

import { presentationInfo } from '../store/index.ts';
import { CommandRefusal } from './commands/op-batch.ts';
import {
  addSection as addSectionEdit,
  deckSections,
  moveSection as moveSectionEdit,
  normalizePresentation,
  removeAllSections,
  removeSection,
  renameSection as renameSectionEdit,
} from './commands/slide-sections.ts';
import type { AutomationItemKey } from './operations.ts';
import type { AutomationRef, AutomationSectionRead, AutomationValue } from './protocol.ts';
import {
  AutomationFailure,
  badArgument,
  edit,
  indexIn,
  slideOf,
  targetOf,
  type PlanContext,
} from './targets.ts';

/**
 * The name of the first section that `addSection` makes in a presentation without sections, for
 * the slides before the new one. The editor's English text (`sections.default`); an automation
 * host has no locale.
 */
const DEFAULT_SECTION_NAME = 'Default Section';

/** The sections as the writer saves them, in order, each with its slide ids in slide order. */
function sectionsOf(ctx: PlanContext): AutomationSectionRead[] {
  const slides = presentationInfo(ctx.batch.pkg).slides;
  // `deckSections` reads the repaired part: the sections hold runs of slides in slide order.
  return deckSections(ctx.batch.pkg).map((s, index) => ({
    id: s.id,
    name: s.name,
    index,
    slideIds: slides.slice(s.firstSlide, s.firstSlide + s.slideCount).map((slide) => slide.sldId),
  }));
}

/** The section a handle names, or a stale-handle failure when an edit removed it. */
function sectionAt(ctx: PlanContext, ref: AutomationRef): AutomationSectionRead {
  const { id } = targetOf(ctx, ref, 'section');
  const hit = sectionsOf(ctx).find((s) => s.id === id);
  if (!hit) throw new AutomationFailure('invalid-handle', 'that section is gone', 'gone');
  return hit;
}

/** Repair the presentation part before a section edit, as a slide-list edit does. */
function normalized(ctx: PlanContext): void {
  edit(() => normalizePresentation(ctx.batch));
}

/** Run a section edit. A name the engine does not accept is the caller's argument. */
function sectionEdit<T>(run: () => T): T {
  return edit(() => {
    try {
      return run();
    } catch (e) {
      if (e instanceof CommandRefusal && e.reason === 'errors.sectionName')
        throw new AutomationFailure(
          'invalid-argument',
          'a section name has 1 to 255 characters',
          'errors.sectionName'
        );
      throw e;
    }
  });
}

/** A section name argument: a string. The engine checks its length and characters. */
function nameArg(name: unknown): string {
  if (typeof name !== 'string') badArgument('name');
  return name;
}

const handleOf = (ctx: PlanContext, id: string): AutomationValue => ({
  kind: 'handle',
  handle: ctx.handles.mint({ kind: 'section', id }),
});

export function listSections(ctx: PlanContext): AutomationValue {
  return {
    kind: 'handles',
    handles: sectionsOf(ctx).map((s) => ctx.handles.mint({ kind: 'section', id: s.id })),
  };
}

/** A section by position or by id. An id matches in either case: a GUID is case-insensitive. */
export function getSection(
  ctx: PlanContext,
  key: AutomationItemKey,
  orNull: boolean | undefined
): AutomationValue {
  const all = sectionsOf(ctx);
  const hit =
    'index' in key
      ? Number.isInteger(key.index) && key.index >= 0
        ? all[key.index]
        : undefined
      : typeof key.id === 'string'
        ? all.find((s) => s.id.toUpperCase() === key.id.toUpperCase())
        : undefined;
  if (hit) return handleOf(ctx, hit.id);
  if (orNull) return { kind: 'null' };
  throw new AutomationFailure('item-not-found', 'no such section', 'index' in key ? 'index' : 'id');
}

export function readSection(ctx: PlanContext, ref: AutomationRef): AutomationValue {
  return { kind: 'section', section: sectionAt(ctx, ref) };
}

/**
 * Add a section that starts at a slide (`addSection`). In a presentation without sections, this
 * adds the section list to the extension list, and the slides before the slide go into a first
 * section. A new section id is a fresh GUID (`ST_Guid`, ECMA-376 Part 1, 22.9.2.4).
 */
export function addSection(
  ctx: PlanContext,
  name: unknown,
  slideRef: AutomationRef
): AutomationValue {
  const label = nameArg(name);
  const { sldId } = targetOf(ctx, slideRef, 'slide');
  normalized(ctx);
  const { index } = slideOf(ctx.batch.pkg, sldId);
  const id = sectionEdit(() => addSectionEdit(ctx.batch, index, label, DEFAULT_SECTION_NAME));
  return handleOf(ctx, id);
}

/** Rename a section (`p14:section/@name`). The current name writes nothing. */
export function renameSection(
  ctx: PlanContext,
  ref: AutomationRef,
  name: unknown
): AutomationValue {
  const label = nameArg(name);
  const section = sectionAt(ctx, ref);
  if (label === section.name) return { kind: 'applied' };
  normalized(ctx);
  sectionEdit(() => renameSectionEdit(ctx.batch, section.id, label));
  return { kind: 'applied' };
}

/**
 * Move a section and its slides to position `to` in the section list: one step past one
 * section at a time (`moveSection`), so the slides keep their order in each section. The
 * current position writes nothing.
 */
export function moveSection(ctx: PlanContext, ref: AutomationRef, to: unknown): AutomationValue {
  const section = sectionAt(ctx, ref);
  const target = indexIn(to, 'index', sectionsOf(ctx).length);
  if (target === section.index) return { kind: 'applied' };
  normalized(ctx);
  const step = target > section.index ? 1 : -1;
  for (let at = section.index; at !== target; at += step)
    sectionEdit(() => moveSectionEdit(ctx.batch, section.id, step));
  return { kind: 'applied' };
}

/**
 * Remove a section (`removeSection`). Its slides join the section before it, or the section
 * after it for the first section. With `deleteSlides`, its slides are deleted with their parts;
 * a section that holds every slide refuses then, because a presentation keeps one slide. The
 * last section takes the section list with it.
 */
export function deleteSection(
  ctx: PlanContext,
  ref: AutomationRef,
  deleteSlides: unknown
): AutomationValue {
  if (deleteSlides !== undefined && typeof deleteSlides !== 'boolean') badArgument('deleteSlides');
  const section = sectionAt(ctx, ref);
  normalized(ctx);
  sectionEdit(() => removeSection(ctx.batch, section.id, deleteSlides === true));
  return { kind: 'applied' };
}

/** Remove every section (`removeAllSections`). The slides stay. No section writes nothing. */
export function deleteAllSections(ctx: PlanContext): AutomationValue {
  if (sectionsOf(ctx).length === 0) return { kind: 'applied' };
  normalized(ctx);
  sectionEdit(() => removeAllSections(ctx.batch));
  return { kind: 'applied' };
}
