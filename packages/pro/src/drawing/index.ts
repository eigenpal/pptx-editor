/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/**
 * `@pptx-editor.dev/pro/drawing`: advanced drawing rendering.
 *
 * The free engine paints pictures, shapes with fills, gradients, patterns, lines and outer
 * shadows, text, tables, and groups. This module adds:
 *
 * - the shape and picture effects glow, inner shadow, reflection, soft edge, and blur
 *   (ECMA-376 Part 1, 20.1.8);
 * - 3-D shapes (ECMA-376 Part 1, 20.1.5): bevels, contours, and, under a turned camera, the
 *   turned face and its extrusion;
 * - charts, drawn from the values cached in their chart parts (ECMA-376 Part 1, 21.2): bar and
 *   column charts (clustered, stacked, and percent stacked), line, area, scatter, pie, doughnut,
 *   and radar charts, with titles, legends, axes, gridlines, data labels, and number formats.
 *   3-D, of-pie, stock, bubble, and surface charts, and charts with trendlines or error bars,
 *   keep the placeholder. A chart becomes shapes and labels that the engine lays out like any
 *   other shape, so the editor and the PDF export draw one layout. Its part saves unchanged.
 *
 * Without it, a chart shows a labeled placeholder, and a shape or picture paints without those
 * effects, and the editor chrome does not offer them (Shape Effects shows only the outer
 * shadows, Picture Styles only the frame styles). The effect commands still write them to the
 * file, and save keeps them. PDF export
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
import { renderChartFrame } from './chart-projection.ts';
import { drawingEffectsRenderer } from './effects-renderer.ts';

/**
 * Install the drawing renderers (charts, glow, inner shadow, reflection, soft edge, blur, and
 * 3-D) for
 * every editor and export in this process, without an editor module: for headless rendering.
 * Installing again is harmless. A host without the EigenPal Pro License leaves this call out;
 * its charts then show placeholders, and its shapes and pictures paint without those effects.
 *
 * @public
 */
export function installDrawing(options: ProLicenseOptions = {}): void {
  rememberLicenseKey(options.licenseKey);
  installDrawingRenderers({ effects: drawingEffectsRenderer, chart: renderChartFrame });
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
