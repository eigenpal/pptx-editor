// Slide commands of the Home and Slide Show tabs (change layout, reset, hide), speaker notes
// formatting, sections, and the slide size. `commands.ts` sends them here; each compiles to one
// transaction, so one undo step.

import { createTranslator } from '@pptx-editor.dev/i18n';
import { presentationInfo, xmlPartOf } from '../store/index.ts';
import type { EditorCommand } from '../contracts/index.ts';
import { refit } from '../automation/commands/autofit.ts';
import type { CommandContext, CommandEffect } from './commands.ts';
import { CommandRefusal, type OpBatch } from '../automation/commands/op-batch.ts';
import { formatNotes } from '../automation/commands/notes-commands.ts';
import { applyLayout, deckLayout } from '../automation/commands/slide-layouts.ts';
import { slideIndices } from '../automation/commands/slide-multi.ts';
import { setSlideSize } from '../automation/commands/slide-size.ts';
import {
  addSection,
  moveSection,
  removeAllSections,
  removeSection,
  renameSection,
  sectionAt,
} from '../automation/commands/slide-sections.ts';

/** The commands this table compiles. */
export type SlideCommand = Extract<
  EditorCommand,
  {
    readonly type:
      | 'applyLayout'
      | 'resetSlides'
      | 'setSlidesHidden'
      | 'formatNotes'
      | 'addSection'
      | 'renameSection'
      | 'removeSection'
      | 'removeAllSections'
      | 'moveSection'
      | 'setSlideSize';
  }
>;

const TYPES: ReadonlySet<string> = new Set([
  'applyLayout',
  'resetSlides',
  'setSlidesHidden',
  'formatNotes',
  'addSection',
  'renameSection',
  'removeSection',
  'removeAllSections',
  'moveSection',
  'setSlideSize',
]);

const english = createTranslator();

/** The host's text for an i18n key, or the English text when the host has none. */
function textOf(ctx: CommandContext, key: 'sections.untitled' | 'sections.default'): string {
  const text = ctx.translate?.(key);
  return text && text !== key ? text : english(key);
}

export function isSlideCommand(command: EditorCommand): command is SlideCommand {
  return TYPES.has(command.type);
}

/**
 * Hide or show slide `index`: `p:sld/@show` (ECMA-376 Part 1, §19.3.1.38) is `0` for a hidden
 * slide; a shown slide drops the attribute, whose default is true.
 */
function setHidden(batch: OpBatch, index: number, hidden: boolean): void {
  const ref = presentationInfo(batch.pkg).slides[index];
  const root = xmlPartOf(batch.pkg, ref?.partName)?.root;
  if (!ref || !root) throw new CommandRefusal('errors.malformed');
  if (ref.hidden === hidden) return;
  batch.push({
    op: 'setAttribute',
    part: ref.partName,
    node: root.id,
    name: { ns: '', local: 'show' },
    value: hidden ? '0' : null,
  });
}

/** Compile a slide command into `batch`. Throws CommandRefusal. */
export function compileSlideCommand(
  ctx: CommandContext,
  batch: OpBatch,
  command: SlideCommand
): CommandEffect {
  switch (command.type) {
    case 'setSlideSize': {
      const content = command.content ?? 'scale';
      if (content !== 'scale' && content !== 'keep')
        throw new CommandRefusal('errors.unsupported', 'content');
      if (!setSlideSize(batch, command.width, command.height, content))
        throw new CommandRefusal('disabled.noSizeChange');
      return { label: 'slideSize' };
    }
    case 'addSection': {
      const index = command.index ?? ctx.slide;
      const name = command.name ?? textOf(ctx, 'sections.untitled');
      addSection(batch, index, name, command.defaultSectionName ?? textOf(ctx, 'sections.default'));
      return { label: 'addSection', slide: index };
    }
    case 'removeAllSections':
      removeAllSections(batch);
      return { label: 'removeAllSections' };
    case 'renameSection':
      renameSection(batch, command.sectionId, command.name);
      return { label: 'renameSection' };
    case 'removeSection': {
      // With its slides: the current slide stays current when it is not deleted; otherwise the
      // slide at the place of the first deleted one becomes current, as after Delete Slide.
      const keep = presentationInfo(batch.pkg).slides[ctx.slide]?.sldId;
      const sectionId = command.sectionId ?? sectionAt(batch.pkg, ctx.slide);
      const at = removeSection(batch, sectionId, command.withSlides ?? false);
      if (at === undefined) return { label: 'removeSection' };
      const kept = presentationInfo(batch.pkg).slides.findIndex((s) => s.sldId === keep);
      return { label: 'removeSection', slide: kept >= 0 ? kept : at, selection: { kind: 'none' } };
    }
    case 'moveSection': {
      const slide = moveSection(batch, command.sectionId, command.direction);
      return {
        label: 'moveSection',
        selection: { kind: 'none' },
        ...(slide !== undefined ? { slide } : {}),
      };
    }
  }
  if (command.type === 'formatNotes') {
    const { index, start, end, format } = command;
    formatNotes(batch, index, start, end, format, command.value);
    return { label: 'notesFormat' };
  }
  const indices = slideIndices(batch, command.indices ?? ctx.selectedSlides ?? [ctx.slide]);
  // Placeholders that moved re-fit their autofit text, as after a resize.
  const refitAll = (index: number, ids: readonly string[]): void => {
    if (ctx.probe) return;
    const part = presentationInfo(batch.pkg).slides[index]!.partName;
    for (const id of ids) refit(batch, index, part, id, ctx.session.measurer);
  };
  switch (command.type) {
    case 'applyLayout': {
      const layout = deckLayout(batch.pkg, command.layoutId);
      for (const i of indices) refitAll(i, applyLayout(batch, i, layout));
      return { label: 'slideLayout', selection: { kind: 'none' } };
    }
    case 'resetSlides': {
      for (const i of indices) {
        const layout = presentationInfo(batch.pkg).slides[i]?.layoutPart;
        if (!layout) throw new CommandRefusal('errors.malformed');
        refitAll(i, applyLayout(batch, i, layout, true));
      }
      return { label: 'resetSlides', selection: { kind: 'none' } };
    }
    case 'setSlidesHidden': {
      // For a selection, the Hide Slide toggle hides the slides unless every one of them is
      // hidden already.
      const slides = presentationInfo(batch.pkg).slides;
      const hidden = command.hidden ?? !indices.every((i) => slides[i]?.hidden);
      for (const i of indices) setHidden(batch, i, hidden);
      return { label: hidden ? 'hideSlides' : 'showSlides' };
    }
  }
}
