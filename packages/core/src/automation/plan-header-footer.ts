// The slide number, footer, and date of slides (`readHeaderFooter`, `setHeaderFooter`), the
// object model's `Slide.headerFooter` and `Presentation.applyHeaderFooter`. The writes are the
// editor's Insert > Slide numbers, Footer, and Date and time commands
// (`commands/header-footer.ts`, `commands/slide-numbers.ts`): each target slide gets or loses
// its own `sldNum`, `ftr`, or `dt` placeholder (`p:ph`, `CT_Placeholder`, ECMA-376 Part 1,
// 19.3.1.36; `ST_PlaceholderType`, 19.7.10), whose box and format come from its layout or
// master. The slide number is a `slidenum` field and an automatic date a `datetime1` field
// (`a:fld`, `CT_TextField`, 21.1.2.2.4).
//
// The writes leave `p:hf` (`CT_HeaderFooter`, 19.3.1.25) of the masters and layouts as it is: it
// records which placeholders a master enables, and a slide shows what its own placeholders hold.
// A slide whose layout and master have no such placeholder is skipped; a write that no target
// slide can show is refused (`disabled.no*Placeholder`), and the batch writes nothing.

import { presentationInfo } from '../store/index.ts';
import {
  MAX_HEADER_FOOTER_TEXT,
  applyHeaderFooter,
  setDateTime,
  setFooter,
  slideHeaderFooter,
} from './commands/header-footer.ts';
import { CommandRefusal } from './commands/op-batch.ts';
import { xmlText } from './commands/package-edits.ts';
import type { AutomationHeaderFooterWrite } from './operations.ts';
import type { AutomationRef, AutomationValue } from './protocol.ts';
import {
  AutomationFailure,
  badArgument,
  edit,
  slideOf,
  slidePartOf,
  targetOf,
  type PlanContext,
} from './targets.ts';

/** The names a `setHeaderFooter` write may carry. */
const NAMES: readonly (keyof AutomationHeaderFooterWrite)[] = [
  'slideNumberVisible',
  'footerVisible',
  'footerText',
  'dateVisible',
  'dateText',
  'dateFixed',
  'hideOnTitleSlides',
];

export function readHeaderFooter(ctx: PlanContext, ref: AutomationRef): AutomationValue {
  const { sldId } = targetOf(ctx, ref, 'slide');
  // Parses a slide that still waits in a progressive open, and refuses a damaged one.
  slidePartOf(ctx.batch.pkg, sldId);
  const read = slideHeaderFooter(ctx.batch.pkg, slideOf(ctx.batch.pkg, sldId).index);
  if (!read) throw new AutomationFailure('invalid-handle', 'that slide is gone', 'gone');
  return {
    kind: 'headerFooter',
    headerFooter: {
      slideNumberVisible: read.slideNumber,
      footerVisible: read.footer !== null,
      footerText: read.footer ?? '',
      dateVisible: read.date !== null,
      dateText: read.date?.text ?? '',
      dateFixed: read.date?.fixed ?? false,
    },
  };
}

/** One line of 1 to 1024 characters that XML can hold, or `invalid-argument` naming it. */
function lineOf(value: unknown, name: string): string {
  if (
    typeof value !== 'string' ||
    value.trim() === '' ||
    value.length > MAX_HEADER_FOOTER_TEXT ||
    /[\r\n\u000b\u0085\u2028\u2029]/.test(value)
  )
    badArgument(name);
  try {
    xmlText(value, name);
  } catch (e) {
    if (e instanceof CommandRefusal) badArgument(name);
    throw e;
  }
  return value;
}

/** A boolean or absent, or `invalid-argument` naming it. */
function flagOf(value: unknown, name: string): boolean | undefined {
  if (value !== undefined && typeof value !== 'boolean') badArgument(name);
  return value as boolean | undefined;
}

/**
 * The checked settings. Every check runs before anything is written. A read of a slide is a
 * valid write: a hidden footer or date may come with empty text, and a hidden date with
 * `dateFixed`, which the write then ignores.
 */
function settingsOf(value: unknown): AutomationHeaderFooterWrite {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) badArgument('settings');
  const raw = value as Record<string, unknown>;
  for (const key of Object.keys(raw))
    if (!(NAMES as readonly string[]).includes(key)) badArgument('settings');
  const footerVisible = flagOf(raw.footerVisible, 'footerVisible');
  const dateVisible = flagOf(raw.dateVisible, 'dateVisible');
  const dateFixed = flagOf(raw.dateFixed, 'dateFixed');
  // Text shows what it names, so a write that hides it takes no text but the empty text.
  const textOf = (v: unknown, name: string, hidden: boolean): string | undefined => {
    if (v === undefined || (hidden && v === '')) return undefined;
    if (hidden) badArgument(name);
    return lineOf(v, name);
  };
  const s = {
    slideNumberVisible: flagOf(raw.slideNumberVisible, 'slideNumberVisible'),
    footerVisible,
    footerText: textOf(raw.footerText, 'footerText', footerVisible === false),
    dateVisible,
    dateText: textOf(raw.dateText, 'dateText', dateVisible === false),
    dateFixed: dateVisible === false ? undefined : dateFixed,
    hideOnTitleSlides: flagOf(raw.hideOnTitleSlides, 'hideOnTitleSlides'),
  };
  const writes =
    s.slideNumberVisible ??
    s.footerVisible ??
    s.footerText ??
    s.dateVisible ??
    s.dateText ??
    s.dateFixed;
  if (writes === undefined) badArgument('settings');
  return s;
}

/**
 * Hide an item. When no target slide can have it, there is nothing to remove, so the hide
 * changes nothing instead of refusing: the engine refuses before it plans any op.
 */
function hide(run: () => unknown): void {
  try {
    run();
  } catch (e) {
    if (e instanceof CommandRefusal && /^disabled\.no\w+Placeholder$/.test(e.reason)) return;
    throw e;
  }
}

export function setHeaderFooter(
  ctx: PlanContext,
  refs: readonly AutomationRef[] | undefined,
  value: unknown
): AutomationValue {
  const pkg = ctx.batch.pkg;
  if (refs !== undefined && (!Array.isArray(refs) || refs.length === 0)) badArgument('slides');
  const slides = presentationInfo(pkg).slides;
  const ids = refs?.map((ref) => targetOf(ctx, ref, 'slide').sldId) ?? slides.map((s) => s.sldId);
  // Each target parses now, and a slide that did not open refuses the write.
  const indices = ids.map((sldId) => {
    slidePartOf(pkg, sldId);
    return slideOf(pkg, sldId).index;
  });
  const s = settingsOf(value);
  const options = { indices, hideOnTitleSlides: s.hideOnTitleSlides ?? false };
  const batch = ctx.batch;
  edit(() => {
    if (s.slideNumberVisible === true)
      applyHeaderFooter(batch, { type: 'setSlideNumbers', show: true, ...options });
    else if (s.slideNumberVisible === false)
      hide(() => applyHeaderFooter(batch, { type: 'setSlideNumbers', show: false, ...options }));
    if (s.footerText !== undefined || s.footerVisible === true)
      setFooter(batch, true, s.footerText, options);
    else if (s.footerVisible === false) hide(() => setFooter(batch, false, undefined, options));
    if (s.dateVisible === false) hide(() => setDateTime(batch, false, undefined, false, options));
    // With `dateFixed: false`, the text is the current text of an automatic date field.
    else if (s.dateText !== undefined)
      setDateTime(batch, true, s.dateText, s.dateFixed !== false, options);
    else if (s.dateVisible !== undefined || s.dateFixed !== undefined)
      setDateTime(batch, true, undefined, s.dateFixed === true, options, s.dateFixed === false);
  });
  return { kind: 'applied' };
}
