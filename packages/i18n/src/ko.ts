/**
 * `@pptx-editor.dev/i18n/ko`: the Korean (`ko`) catalog, in the shape of the `ja` and `zh-CN`
 * subpaths. A `null` entry falls back to English.
 *
 * ```ts
 * import ko from '@pptx-editor.dev/i18n/ko';
 * // <PptxEditor locale={ko} />
 * ```
 *
 * @packageDocumentation
 * @public
 */

import data from '../locales/ko.json' with { type: 'json' };
import type { LocaleCatalog } from './index.ts';

/** Korean (`ko`) strings. A `null` entry falls back to English. @public */
export const ko: LocaleCatalog = data;

export default ko;
