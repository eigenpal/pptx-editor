---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
---

Without the drawing module of `@pptx-editor.dev/pro`, the chrome no longer offers glow, inner shadow, reflection, or soft edge. Shape Effects shows only the outer shadows, and Picture Styles shows only the frame styles. The effect commands still write every effect, and save keeps it. New `drawingCapabilities()` in `@pptx-editor.dev/core/layout` says what the installed renderers paint, and `chromeOfferedControl(slot, control)` in `@pptx-editor.dev/core/editor` gives the choices a slot offers in this build. The controls and the `useEditorValueCommand` options of React and Vue refresh after an install; `drawingRenderersVersion()` and `subscribeDrawingRenderers(listener)` in `@pptx-editor.dev/core/layout` give the signal.
