---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add `Shape.chart` to the editing API, an SDK extension. `Chart` reads the type, title, categories, series, and workbook state of a chart, and `setData` replaces its data. The new automation operations `readChart` and `setChartData` use the editor's chart edit, so the caches and an embedded workbook change together.
