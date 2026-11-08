// Shared glue of the slide master view's rail (React and Vue `MasterRail`): the names, the
// usage text, and the keys that move between templates, and what the view turns off.
// `@internal`.

import type { TranslationKey, Translate } from '@pptx-editor.dev/i18n';
import type { EditorSnapshot, SlideSummary } from '../contracts/index.ts';

/**
 * The `slides` entry of the slide or template on the canvas. In the slide master view,
 * `currentSlide` is a template index, not a position in `slides`, so `slides[currentSlide]` is
 * undefined there: read the current entry with this function. @public
 */
export function currentSlideSummary(
  snapshot: Pick<EditorSnapshot, 'currentSlide' | 'slides'>
): SlideSummary | undefined {
  const at = snapshot.slides[snapshot.currentSlide];
  return at?.index === snapshot.currentSlide
    ? at
    : snapshot.slides.find((s) => s.index === snapshot.currentSlide);
}

/** The name of a rail template: "Master" for a master, else the layout name. @internal */
export function masterRailLabel(t: Translate, s: SlideSummary): string {
  if (s.template === 'master')
    return s.name ? `${t('masterView.master')}: ${s.name}` : t('masterView.master');
  return s.name || t('masterView.untitledLayout');
}

/**
 * The title of the bar over the canvas: "Layout: Title Slide" or "Master: Office Theme",
 * the name of what the canvas shows. @internal
 */
export function masterBarTitle(t: Translate, s: SlideSummary): string {
  if (s.template === 'master')
    return t('canvas.masterAriaLabel', { name: s.name || t('masterView.master') });
  return t('canvas.layoutAriaLabel', { name: s.name || t('masterView.untitledLayout') });
}

/** What an edit of the template on the canvas changes, in one sentence. @internal */
export function masterBarHint(t: Translate, s: SlideSummary): string {
  const count = s.usedBy ?? 0;
  if (s.template === 'master') {
    if (count === 0) return t('masterView.masterHintNone');
    return count === 1 ? t('masterView.masterHintOne') : t('masterView.masterHint', { count });
  }
  if (count === 0) return t('masterView.layoutHintNone');
  return count === 1 ? t('masterView.layoutHintOne') : t('masterView.layoutHint', { count });
}

/** "Used by 3 slides", "Used by 1 slide", or "Not used by any slide". @internal */
export function masterRailUsage(t: Translate, s: SlideSummary): string {
  const count = s.usedBy ?? 0;
  if (count === 0) return t('masterView.unused');
  return count === 1 ? t('masterView.usedByOne') : t('masterView.usedBy', { count });
}

/**
 * The template index that a key moves to in the master rail: ArrowUp / ArrowLeft and
 * ArrowDown / ArrowRight step, Home and End go to the ends. Null for other keys and at the
 * ends. @internal
 */
export function masterStepTarget(
  key: string,
  templates: readonly SlideSummary[],
  current: number
): number | null {
  if (templates.length === 0) return null;
  const at = templates.findIndex((s) => s.index === current);
  let to: number;
  if (key === 'Home') to = 0;
  else if (key === 'End') to = templates.length - 1;
  else if (key === 'ArrowUp' || key === 'ArrowLeft') to = Math.max(0, at - 1);
  else if (key === 'ArrowDown' || key === 'ArrowRight') to = Math.min(templates.length - 1, at + 1);
  else return null;
  const next = templates[to]?.index;
  return next === undefined || next === current ? null : next;
}

/**
 * Why Find cannot open in `view`, or null: it shows its matches on the slide canvas, which the
 * slide sorter does not show, and the slide master view shows layouts, not slides. @internal
 */
export function findViewRefusal(view: string | undefined): string | null {
  if (view === 'slideSorter') return 'disabled.findSorter';
  return view === 'slideMaster' ? 'disabled.slideMaster' : null;
}

/**
 * The heading and the empty-state text of the side panel's Format tab for the canvas: a slide,
 * or a layout or master in the slide master view. @internal
 */
export function formatPanelKeys(snapshot: Pick<EditorSnapshot, 'currentSlide' | 'slides'>): {
  readonly title: TranslationKey;
  readonly body: TranslationKey;
} {
  const template = currentSlideSummary(snapshot)?.template;
  if (!template) return { title: 'panel.format.slide', body: 'panel.format.emptyBody' };
  const title = template === 'master' ? 'panel.format.master' : 'panel.format.layout';
  return { title, body: 'panel.format.emptyBodyTemplate' };
}

/** The editor as `deckSlideCount` reads it. */
interface SessionHolder {
  readonly session?: { readonly info: { readonly slides: readonly unknown[] } } | null;
}

/**
 * The number of slides of the open deck, in every view. In the slide master view, the
 * snapshot's `slides` lists layouts and masters, so a slide link picker reads this. @internal
 */
export function deckSlideCount(editor: SessionHolder | null | undefined): number {
  return editor?.session?.info.slides.length ?? 0;
}
