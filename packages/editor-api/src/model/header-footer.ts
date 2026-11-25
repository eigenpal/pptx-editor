/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The slide number, footer, and date of a slide: `Slide.headerFooter` and
// `Presentation.applyHeaderFooter`, SDK extensions with no Office.js counterpart. A slide shows
// them through its own placeholders of type `sldNum`, `ftr`, and `dt` (`p:ph`,
// `CT_Placeholder`, ECMA-376 Part 1, 19.3.1.36; `ST_PlaceholderType`, 19.7.10). A write gives the
// slide the placeholder of its layout or master, or removes the slide's own. The slide number is
// a `slidenum` field and an automatic date a `datetime1` field (`a:fld`, `CT_TextField`,
// 21.1.2.2.4). `p:hf` of the masters and layouts (`CT_HeaderFooter`, 19.3.1.25) stays as it is.

import type { AutomationHeaderFooterWrite } from '@pptx-editor.dev/core/automation';
import { failArgument, nearest, oneOf } from '../runtime/errors.ts';
import type { ResolvedLoadOptions } from '../runtime/load-options.ts';
import type { ObjectPath } from '../runtime/object-path.ts';
import type { RequestContext } from '../runtime/request-context.ts';
import { flag, ModelObject, text } from './model-object.ts';

/**
 * The settings that `Presentation.applyHeaderFooter` applies. A setting you leave out stays as
 * each slide has it. Pass at least one of the slide number, footer, or date settings. The
 * `toJSON()` result of a loaded `SlideHeaderFooter` is valid settings.
 *
 * @public
 */
export interface HeaderFooterSettings {
  /** Show or hide the slide number. */
  slideNumberVisible?: boolean;
  /**
   * Show or hide the footer. `true` without `footerText` keeps each slide's footer text. A slide
   * with no footer takes the footer text of another slide.
   */
  footerVisible?: boolean;
  /**
   * Show the footer with this text: one line of 1 to 1024 characters. With `footerVisible:
   * false`, only an empty string is accepted, and it changes nothing.
   */
  footerText?: string;
  /**
   * Show or hide the date. `true` alone keeps each slide's date. A slide with no date takes the
   * date of another slide, else today's date as an automatic date.
   */
  dateVisible?: boolean;
  /**
   * Show the date with this text: one line of 1 to 1024 characters. The date is fixed text,
   * unless `dateFixed` is `false`: then it is an automatic date field that holds this text. With
   * `dateVisible: false`, only an empty string is accepted, and it changes nothing.
   */
  dateText?: string;
  /**
   * Show the date, as fixed text (`true`) or as an automatic date (`false`). `true` keeps the
   * text that each date shows now. `false` turns fixed text into today's date as month/day/year,
   * and keeps an automatic date in its own format. With `false`, a slide with no date takes the
   * automatic date of another slide, else today's date. With `dateVisible: false`, it is ignored.
   */
  dateFixed?: boolean;
  /**
   * Hide the slide number, footer, and date that this call shows on the slides whose layout is a
   * title slide layout. Default `false`.
   */
  hideOnTitleSlides?: boolean;
}

const NAMES: readonly (keyof HeaderFooterSettings)[] = [
  'slideNumberVisible',
  'footerVisible',
  'footerText',
  'dateVisible',
  'dateText',
  'dateFixed',
  'hideOnTitleSlides',
];

/** The checked settings as the host takes them, or `InvalidArgument` naming the first fault. */
export function headerFooterWrite(value: unknown, target: string): AutomationHeaderFooterWrite {
  if (typeof value !== 'object' || value === null || Array.isArray(value))
    failArgument(target, 'a HeaderFooterSettings object');
  const raw = value as Record<string, unknown>;
  for (const key of Object.keys(raw))
    if (!(NAMES as readonly string[]).includes(key))
      failArgument(`${target}.${key}`, `a setting name: ${oneOf(NAMES)}`, nearest(key, NAMES));
  const out: Record<string, boolean | string> = {};
  for (const name of NAMES) {
    const v = raw[name];
    if (v === undefined) continue;
    out[name] =
      name === 'footerText' || name === 'dateText'
        ? text(v, `${target}.${name}`)
        : flag(v, `${target}.${name}`);
  }
  return out as AutomationHeaderFooterWrite;
}

/**
 * The slide number, footer, and date of one slide. Load the properties to read what the slide
 * shows. Set a property to change this slide only. To change all slides at once, call
 * `Presentation.applyHeaderFooter`.
 *
 * A slide shows each item through its own placeholder. A write that shows an item gives the
 * slide the placeholder that its layout defines, else the one that its master defines, so the
 * item takes the position and format of the layout. A slide that has the placeholder already
 * keeps its box and format. A write that hides an item removes the slide's placeholders of that
 * type; with no such placeholder, it changes nothing. When the layout and the master of the
 * slide have no such placeholder, a write that shows the item refuses with `NotSupported` and
 * the diagnostic `edit.no-slide-number-placeholder`, `edit.no-footer-placeholder`, or
 * `edit.no-date-placeholder`. A slide that did not open refuses with `DocumentUnavailable`.
 *
 * In a runtime over a live editor, the writes of one sync are one undo step. This object is an
 * SDK extension. The pinned Office.js reference has no such member.
 *
 * @example
 * ```ts
 * const hf = context.presentation.slides.getItemAt(1).headerFooter;
 * hf.slideNumberVisible = true;
 * hf.footerText = 'Quarterly review';
 * hf.load('slideNumberVisible,footerText');
 * await context.sync();
 * ```
 *
 * @public
 */
export class SlideHeaderFooter extends ModelObject {
  /** @internal The slide. */
  readonly #slide: ObjectPath;

  /** @internal */
  constructor(context: RequestContext, path: ObjectPath, slide: ObjectPath) {
    super(context, path);
    this.#slide = slide;
  }

  /**
   * Whether the slide has a slide number placeholder. Unlike the footer and the date, a
   * placeholder with no text reads as `true`. A new number is a `slidenum` field, and the engine
   * shows the position of the slide in the presentation.
   */
  get slideNumberVisible(): boolean {
    return this.loadedProperty('slideNumberVisible');
  }
  set slideNumberVisible(value: boolean) {
    this.write('slideNumberVisible', {
      slideNumberVisible: flag(value, this.at('slideNumberVisible')),
    });
  }

  /**
   * Whether the slide shows a footer. A footer placeholder with no text reads as `false`. A write
   * of `true` keeps the slide's footer text. A slide with no footer text takes the footer text of
   * another slide. When no slide has footer text, the write refuses with `NotSupported` and the
   * diagnostic `edit.no-footer-text`. Set `footerText` instead to give the text.
   */
  get footerVisible(): boolean {
    return this.loadedProperty('footerVisible');
  }
  set footerVisible(value: boolean) {
    this.write('footerVisible', { footerVisible: flag(value, this.at('footerVisible')) });
  }

  /**
   * The footer text, or an empty string when the slide shows no footer. A write shows the footer
   * with this text: one line of 1 to 1024 characters. Other text refuses with `InvalidArgument`.
   */
  get footerText(): string {
    return this.loadedProperty('footerText');
  }
  set footerText(value: string) {
    this.write('footerText', { footerText: text(value, this.at('footerText')) });
  }

  /**
   * Whether the slide shows a date. A write of `true` keeps the slide's date. A slide with no
   * date takes the date of another slide, else today's date as an automatic date.
   */
  get dateVisible(): boolean {
    return this.loadedProperty('dateVisible');
  }
  set dateVisible(value: boolean) {
    this.write('dateVisible', { dateVisible: flag(value, this.at('dateVisible')) });
  }

  /**
   * The text that the date shows, or an empty string when the slide shows no date. For an
   * automatic date, this is the date text that the file holds. The engine does not update it
   * when the presentation opens. A write shows the date as this fixed text: one line of 1 to 1024
   * characters.
   */
  get dateText(): string {
    return this.loadedProperty('dateText');
  }
  set dateText(value: string) {
    this.write('dateText', { dateText: text(value, this.at('dateText')) });
  }

  /**
   * Whether the date is fixed text (`true`) or an automatic date field (`false`). `false` when
   * the slide shows no date. A write shows the date. `true` keeps the text that the date shows
   * now. `false` turns fixed text into today's date as month/day/year, and keeps an automatic
   * date in its own format. A slide with no date takes the automatic date of another slide, else
   * today's date.
   */
  get dateFixed(): boolean {
    return this.loadedProperty('dateFixed');
  }
  set dateFixed(value: boolean) {
    this.write('dateFixed', { dateFixed: flag(value, this.at('dateFixed')) });
  }

  /** @internal Queue one write to this slide. */
  private write(name: string, settings: AutomationHeaderFooterWrite): void {
    const slide = this.#slide;
    this.command(name, (ref) => ({ op: 'setHeaderFooter', slides: [ref(slide)], settings }));
  }

  /** @internal */
  protected override onLoad(request: ResolvedLoadOptions): void {
    const slide = this.#slide;
    this.loadScalars(
      this.selection(request, [
        'slideNumberVisible',
        'footerVisible',
        'footerText',
        'dateVisible',
        'dateText',
        'dateFixed',
      ]),
      'headerFooter',
      (ref) => ({ op: 'readHeaderFooter', slide: ref(slide) }),
      (v) => ({ ...v.headerFooter })
    );
  }
}
