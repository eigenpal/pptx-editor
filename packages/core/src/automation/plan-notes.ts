// The speaker notes of a slide for the object model (`Slide.speakerNotes`, an SDK extension).
//
// The notes live in the body placeholder (`p:ph` with `type="body"`, ECMA-376 Part 1, 19.3.1.36
// and 19.7.10 `ST_PlaceholderType`) of the slide's notes slide: the `p:notes` root
// (`CT_NotesSlide`, Part 1, 19.3.1.26) of a Notes Slide part (Part 1, 13.3.5). The read and the
// write are the editor's own (`commands/notes-commands.ts`): paragraphs (`a:p`,
// `CT_TextParagraph`, Part 1, 21.1.2.2.6) and line breaks (`a:br`, `CT_TextLineBreak`, Part 1,
// 21.1.2.2.1) read as `\n`, and each line of a write becomes a paragraph. A slide with no notes
// slide gets one, with a relationship to the Notes Master part (Part 1, 13.3.4).

import { isValidXmlText } from '../store/index.ts';
import { CommandRefusal } from './commands/op-batch.ts';
import { notesText, setNotes } from './commands/notes-commands.ts';
import type { AutomationRef, AutomationValue } from './protocol.ts';
import {
  AutomationFailure,
  badArgument,
  edit,
  slideOf,
  targetOf,
  type PlanContext,
} from './targets.ts';

/** The longest notes text one write may carry, as for a shape's text. */
const MAX_NOTES_TEXT = 1_000_000;

/** The speaker notes of a slide as plain text: `""` when the slide has no notes slide. */
export function readSlideNotes(ctx: PlanContext, ref: AutomationRef): AutomationValue {
  const { index } = slideOf(ctx.batch.pkg, targetOf(ctx, ref, 'slide').sldId);
  return { kind: 'text', text: notesText(ctx.batch.pkg, index) };
}

/**
 * Set the speaker notes of a slide to `text`, one paragraph for each line (`\n`, `\r\n`, `\r`,
 * or `\v` ends a line). Only the text that differs is rewritten, so unchanged runs keep their
 * formatting. Text with a character XML cannot hold is refused as a whole.
 */
export function setSlideNotes(ctx: PlanContext, ref: AutomationRef, text: string): AutomationValue {
  const { sldId } = targetOf(ctx, ref, 'slide');
  if (typeof text !== 'string' || text.length > MAX_NOTES_TEXT) badArgument('text');
  if (!isValidXmlText(text.replace(/\v/g, ' '))) badArgument('text');
  const { index } = slideOf(ctx.batch.pkg, sldId);
  try {
    setNotes(ctx.batch, index, text);
  } catch (e) {
    // Content the presentation lacks is unsupported content, not a refused transaction.
    const detail = e instanceof CommandRefusal && e.reason === 'errors.unsupported' && e.detail;
    if (detail === 'no notes master')
      throw new AutomationFailure(
        'unsupported-content',
        'the presentation has no notes master',
        'notes-master'
      );
    if (detail === 'notes without a body placeholder')
      throw new AutomationFailure(
        'unsupported-content',
        'the notes slide has no body placeholder with a text body',
        'notes-body'
      );
    edit(() => {
      throw e;
    });
  }
  return { kind: 'applied' };
}
