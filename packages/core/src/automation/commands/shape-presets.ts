// Which preset geometries a command may insert. DOM-free, so the automation lane and the
// insert-shape gallery (`editor/shape-gallery.ts`) answer the same question the same way.

import { presetNames } from '../../layout/index.ts';

/** Whether a preset can be inserted: any ECMA-376 preset geometry (`ST_ShapeType`). */
export function isInsertablePreset(preset: string): boolean {
  return preset === 'line' || presetNames().includes(preset);
}
