---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/pro': minor
'@pptx-editor.dev/pptx-to-pdf': minor
---

Charts now draw from the values cached in their chart parts when `@pptx-editor.dev/pro/drawing` is installed (`drawingModule({ licenseKey })` or `installDrawing()`). Bar, column, line, area, scatter, pie, doughnut, and radar charts draw with titles, legends, axes, gridlines, data labels, and number formats. Other chart types, and every chart without the module, keep the placeholder. `@pptx-editor.dev/pptx-to-pdf` draws charts in its PDFs, and reports `frame.chart` only for the chart types the module leaves out. The core adds the `chart` renderer key to `installDrawingRenderers`, the resolved `chart` part of a chart frame, the laid-out `chart` items of a frame, and `drawingCapabilities().chart`.
