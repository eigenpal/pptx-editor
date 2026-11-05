/**
 * `@pptx-editor.dev/i18n/zh-CN`: the Simplified Chinese (`zh-CN`) catalog, as DOCX Editor's
 * `@docx-editor.dev/i18n/zh-CN` subpath. A partial catalog: a `null` entry falls back to
 * English.
 *
 * ```ts
 * import zhCN from '@pptx-editor.dev/i18n/zh-CN';
 * // <PptxEditor locale={zhCN} />
 * ```
 *
 * @packageDocumentation
 * @public
 */

import data from '../locales/zh-CN.json' with { type: 'json' };
import type { LocaleCatalog } from './index.ts';

/** Simplified Chinese (`zh-CN`) strings. A `null` entry falls back to English. @public */
export const zhCN: LocaleCatalog = data;

export default zhCN;
