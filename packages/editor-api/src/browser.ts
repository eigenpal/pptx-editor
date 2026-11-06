/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/**
 * Presentation automation over an editor a reader already has open, plus everything the root
 * entry offers.
 *
 * `createBrowser` lives one subpath along because reaching a live editor means reaching the
 * painted engine; a consumer that holds bytes does not pay for that.
 *
 * @example Format the reader's selected text
 * ```ts
 * import { PptxEditor } from '@pptx-editor.dev/editor-api/browser';
 *
 * const runtime = PptxEditor.createBrowser(editor);
 * await runtime.run(async (context) => {
 *   const range = context.presentation.getSelectedTextRangeOrNullObject();
 *   range.load('text');
 *   await context.sync();
 *   if (!range.isNullObject) range.font.color = '#C00000';
 *   await context.sync();
 * });
 * ```
 *
 * @packageDocumentation
 */

import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { createBrowser } from './runtime/browser.ts';
import type { PptxEditorRuntime, PptxEditorServerRuntime } from './runtime/runtime.ts';
import { createServer, type CreateServerOptions } from './runtime/server.ts';

export * from './runtime/public.ts';

/**
 * The entry point with `createBrowser`.
 *
 * @public
 */
export interface PptxEditorBrowserNamespace {
  /**
   * A runtime over an editor that is already open. The editor keeps its own lifetime: dispose
   * the runtime before the editor. Each `sync()` that writes is one undo step.
   */
  createBrowser(editor: PptxEditorInstance): PptxEditorRuntime;
  /** A runtime over PPTX bytes. It also offers `save()`. */
  createServer(bytes: Uint8Array, options?: CreateServerOptions): Promise<PptxEditorServerRuntime>;
}

/**
 * The entry point: borrow a live editor, or open bytes.
 *
 * @public
 */
export const PptxEditor: PptxEditorBrowserNamespace = Object.freeze({
  createBrowser,
  createServer,
});
