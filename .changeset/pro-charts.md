---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/pro': minor
'@pptx-editor.dev/pptx-to-pdf': minor
---

Draw charts from the values cached in their chart parts with the new `@pptx-editor.dev/pro/drawing` entry: `drawingModule({ licenseKey })` for an editor and `installDrawing()` for headless use. Bar, column, line, area, scatter, pie, doughnut, and radar charts draw with titles, legends, axes, gridlines, data labels, and number formats; other chart types keep the placeholder. The PDF export always installs the module. The core adds `installDrawingRenderers` and `drawingRenderers` to `@pptx-editor.dev/core/layout`, the `chart` part of a chart frame, and the `chart` items of a laid-out frame.
