/**
 * `@pptx-editor.dev/i18n/ja`: the Japanese (`ja`) catalog, as DOCX Editor's
 * `@docx-editor.dev/i18n/ja` subpath. A partial catalog: a `null` entry falls back to English.
 *
 * ```ts
 * import ja from '@pptx-editor.dev/i18n/ja';
 * // <PptxEditor locale={ja} />
 * ```
 *
 * @packageDocumentation
 * @public
 */

import data from '../locales/ja.json' with { type: 'json' };
import type { LocaleCatalog } from './index.ts';

/** Japanese (`ja`) strings. A `null` entry falls back to English. @public */
export const ja: LocaleCatalog = data;

export default ja;
