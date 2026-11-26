---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add the SDK extension `ShapeCollection.addChart(type, data, options?)`. It adds a chart with an embedded workbook and returns the new shape. The automation `addShape` operation takes a `chart` kind. A chart edit now refuses a chart part over 16 MiB before it writes. `Chart.setData` now refuses a store rejection with `NotSupported`, not `InvalidArgument`.
