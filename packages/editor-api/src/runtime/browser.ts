/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Adapted from DOCX Editor's `packages/editor-api/src/runtime/browser.ts` (commit 5993bd223945).
//
// A runtime that borrows an editor a reader already has open. Each `sync()` that writes is one
// undo step in the editor, and the selection members read and move the reader's selection.

import { createBrowserAutomationHost, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { authorOption, failArgument } from './errors.ts';
import { createRuntime, type PptxEditorRuntime } from './runtime.ts';

/**
 * How `PptxEditor.createBrowser` borrows a live editor.
 *
 * @public
 */
export interface CreateBrowserOptions {
  /**
   * The author of the comments and replies this runtime adds when a call names none
   * (`authorName`). Default: the name that signs the editor's own comments, which is the
   * editor's `author` option, else the collaboration identity. With neither, a comment write
   * that names no author refuses with `NotSupported`. A name has 1 to 255 characters. This is an
   * SDK extension.
   */
  readonly author?: string;
}

export function createBrowser(
  editor: PptxEditorInstance,
  options: CreateBrowserOptions = {}
): PptxEditorRuntime {
  if (options === null || typeof options !== 'object')
    failArgument('createBrowser', 'an options object');
  const author = authorOption(options.author, 'createBrowser');
  return createRuntime({
    host: createBrowserAutomationHost(editor, author === undefined ? {} : { author }),
    save: false,
  });
}
