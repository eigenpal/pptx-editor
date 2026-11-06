---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/pro': minor
'@pptx-editor.dev/pptx-to-pdf': minor
---

Glow, inner shadow, reflection, and soft edge now paint through the new `@pptx-editor.dev/pro/drawing` entry, as in DOCX Editor. Register `drawingModule({ licenseKey })` in `modules`, or call `installDrawing()` for headless rendering. Without it, a shape or picture paints without these four effects; outer shadows still paint, and the effect commands and save keep every effect. `@pptx-editor.dev/pptx-to-pdf` always installs the module, and its PDFs now draw reflections with their fade (`effect.reflection-blur` reports the blur that is left out). `@pptx-editor.dev/core/output` no longer exports the internal `glowSpread` and `innerShadowShift`.
