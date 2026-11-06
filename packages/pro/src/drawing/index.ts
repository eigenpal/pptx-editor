/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/**
 * `@pptx-editor.dev/pro/drawing`: advanced drawing rendering.
 *
 * The free engine paints pictures, shapes with fills, gradients, patterns, lines and outer
 * shadows, text, tables, and groups. This module adds the shape and picture effects glow, inner
 * shadow, reflection, and soft edge (ECMA-376 Part 1, 20.1.8).
 *
 * Without it, a shape or picture paints without those effects, and the editor chrome does not
 * offer them (Shape Effects shows only the outer shadows, Picture Styles only the frame styles).
 * The effect commands still write them to the file, and save keeps them. PDF export
 * (`@pptx-editor.dev/pptx-to-pdf`) always includes this module.
 *
 * The renderers serve every editor and export in the process. Install them before any document
 * opens; a later install logs a console warning once. The chrome controls refresh after an
 * install.
 *
 * @example Enable advanced drawing rendering in an editor
 * ```ts
 * import { createPptxEditor } from '@pptx-editor.dev/core';
 * import { drawingModule } from '@pptx-editor.dev/pro/drawing';
 *
 * // Next.js; with Vite, read import.meta.env.VITE_PPTX_LICENSE.
 * const modules = [drawingModule({ licenseKey: process.env.NEXT_PUBLIC_PPTX_LICENSE })];
 * const editor = createPptxEditor({ document: bytes, modules });
 * ```
 *
 * @packageDocumentation
 * @public
 */

import { installDrawingRenderers } from '@pptx-editor.dev/core/layout';
import type { EditorModule } from '@pptx-editor.dev/core/editor';
import { rememberLicenseKey, type ProLicenseOptions } from '../license.ts';
import { drawingEffectsRenderer } from './effects-renderer.ts';

/**
 * Install the drawing renderers (glow, inner shadow, reflection, and soft edge) for every editor
 * and export in this process, without an editor module: for headless rendering. Installing
 * again is harmless. A host without the EigenPal Pro License leaves this call out; its shapes
 * and pictures then paint without those effects.
 *
 * @public
 */
export function installDrawing(options: ProLicenseOptions = {}): void {
  rememberLicenseKey(options.licenseKey);
  installDrawingRenderers({ effects: drawingEffectsRenderer });
}

/**
 * The drawing module for `createPptxEditor({ modules })` and the adapters' `modules` prop.
 * Creating it installs the drawing renderers process-wide, as {@link installDrawing} does.
 *
 * @public
 */
export function drawingModule(options: ProLicenseOptions = {}): EditorModule {
  installDrawing(options);
  return Object.freeze({ id: 'drawing' });
}

// Internal: the effect rules that the PDF writer (`@pptx-editor.dev/pptx-to-pdf`) shares with
// the SVG renderer.
export {
  glowSpread,
  innerShadowShift,
  reflectionBlurSigma,
  reflectionFade,
  type ReflectionFade,
  type SpreadBlur,
} from './effects-math.ts';
export type { ProLicenseOptions } from '../license.ts';
