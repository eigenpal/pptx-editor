// A localized summary of a document refresh change, for a host's change list. A processor
// change keeps its own `description`; a computed change has none, so this describes it from its
// `computed` fields with the `refresh.change.*` strings of `@pptx-editor.dev/i18n`.

import type { Translate, TranslationKey } from '@pptx-editor.dev/i18n';
import type { RefreshChange } from './document-refresh-types.ts';

/** English, for a host without a `translate` function. */
const ENGLISH: Readonly<Record<string, string>> = {
  'added.slide': 'Added slide {slide}',
  'added.shape': 'Added a shape to slide {slide}',
  'removed.slide': 'Removed slide {slide}',
  'removed.shape': 'Removed a shape from slide {slide}',
  'moved.slide': 'Moved slide {slide}',
  'changed.slide': 'Changed slide {slide}',
  'changed.shape': 'Changed a shape on slide {slide}',
  'changed.text': 'Changed text on slide {slide}',
  'changed.notes': 'Changed the notes of slide {slide}',
  'changed.layout': 'Changed a layout. Slides that use it: {count}',
  'changed.master': 'Changed a slide master. Slides that use it: {count}',
  'changed.theme': 'Changed the theme. Slides that use it: {count}',
  'changed.package': 'Changed the presentation settings',
};

/**
 * The text of a change for a change list: its `description`, or a summary of its `computed`
 * fields ("Changed text on slide 3"). Pass the editor's `translate` function to localize the
 * summary; without it, the summary is English. Render the result as text.
 *
 * @example
 * ```ts
 * for (const change of refresh.snapshot().changes)
 *   item.textContent = describeRefreshChange(change, t);
 * ```
 *
 * @public
 */
export function describeRefreshChange(
  change: RefreshChange,
  translate?: Translate | ((key: string, params?: Record<string, string | number>) => string)
): string {
  if (change.description !== undefined) return change.description;
  const c = change.computed;
  if (!c) return '';
  const scoped = c.scope === 'slide' || c.scope === 'shape';
  const action = scoped && (c.action !== 'moved' || c.scope === 'slide') ? c.action : 'changed';
  const key = `${action}.${c.scope}`;
  const index = change.location?.slideIndex ?? c.previousSlideIndex;
  const params = {
    slide: index === undefined ? '' : index + 1,
    count: c.affectedSlides?.length ?? 0,
  };
  const template = translate
    ? translate(`refresh.change.${key}` as TranslationKey, params)
    : undefined;
  if (template !== undefined && template !== `refresh.change.${key}`) return template;
  return (ENGLISH[key] ?? '').replace(/\{(\w+)\}/g, (_, name: string) =>
    String(params[name as keyof typeof params] ?? '')
  );
}
