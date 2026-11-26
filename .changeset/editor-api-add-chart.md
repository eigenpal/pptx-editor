---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add `ShapeCollection.addChart(type, data, options?)` to the editing API, an SDK extension. It adds a chart of a `ChartType` and returns the new shape. An embedded workbook holds the same data. The automation `addShape` operation takes a `chart` kind, which uses the editor's Insert Chart edit.

A chart edit now refuses when its chart part is larger than 16 MiB. This applies to new data, to a new chart, and to format edits such as a type change. The edit refuses before it writes anything. Before, the edit wrote a part that then read back as no chart.

`Chart.setData` now refuses a store rejection with `NotSupported`, not with `InvalidArgument`.
