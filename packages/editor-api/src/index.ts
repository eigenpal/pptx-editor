/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/**
 * Presentation automation for PPTX: a batching object model over the PPTX Editor engine.
 *
 * Work is described against proxy objects, and nothing reaches the presentation until
 * `context.sync()`, which sends the queued calls as one atomic batch: all of it applies, or none
 * of it. Reading a property that no `load` asked for is an error, not a silent `undefined`.
 *
 * This entry needs no browser. It opens PPTX bytes, drives them, and saves them back, so a
 * server, a worker, or a build script can use the whole API.
 *
 * @example Edit a presentation on a server
 * ```ts
 * import { PptxEditor } from '@pptx-editor.dev/editor-api';
 *
 * const runtime = await PptxEditor.createServer(bytes);
 * try {
 *   await runtime.run(async (context) => {
 *     const slide = context.presentation.slides.getItemAt(0);
 *     const box = slide.shapes.addTextBox('Quarterly results', { left: 40, top: 40, width: 400 });
 *     box.textFrame.textRange.font.bold = true;
 *     await context.sync();
 *   });
 *   const edited = await runtime.save();
 * } finally {
 *   runtime.dispose();
 * }
 * ```
 *
 * The API shape follows a documented subset of the Office.js presentation object model. It is
 * not Office.js, it does not run in an add-in host, and every type in it is authored here.
 *
 * @packageDocumentation
 */

import { createServer, type CreateServerOptions } from './runtime/server.ts';
import type { PptxEditorServerRuntime } from './runtime/runtime.ts';

export * from './runtime/public.ts';

/**
 * The entry point, as much of it as works without an editor. Import from
 * `@pptx-editor.dev/editor-api/browser` for the same object plus `createBrowser`.
 *
 * @public
 */
export interface PptxEditorNamespace {
  /** A runtime over PPTX bytes. It also offers `save()`. */
  createServer(bytes: Uint8Array, options?: CreateServerOptions): Promise<PptxEditorServerRuntime>;
}

/**
 * The entry point: open a presentation and get a runtime to work against.
 *
 * @example
 * ```ts
 * const runtime = await PptxEditor.createServer(bytes);
 * const titles = await runtime.run(async (context) => {
 *   const slides = context.presentation.slides;
 *   slides.load('items');
 *   await context.sync();
 *   const shapes = slides.items.map((slide) => slide.shapes.getItemAt(0).load('name'));
 *   await context.sync();
 *   return shapes.map((shape) => shape.name);
 * });
 * ```
 *
 * @public
 */
export const PptxEditor: PptxEditorNamespace = Object.freeze({ createServer });
