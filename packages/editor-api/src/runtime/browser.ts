/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After DOCX Editor's `packages/editor-api/src/runtime/browser.ts` (commit 5993bd223945).
//
// A runtime that borrows an editor a reader already has open. Each `sync()` that writes is one
// undo step in the editor, and the selection members read and move the reader's selection.

import { createBrowserAutomationHost, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { createRuntime, type PptxEditorRuntime } from './runtime.ts';

export function createBrowser(editor: PptxEditorInstance): PptxEditorRuntime {
  return createRuntime({ host: createBrowserAutomationHost(editor), save: false });
}
