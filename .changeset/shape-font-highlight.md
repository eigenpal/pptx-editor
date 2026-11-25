---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add `ShapeFont.highlightColor` to the editing API, an SDK extension. It reads the text highlight as `#RRGGBB`, or `null` for no highlight and for a mixed range. A color write sets `a:highlight`, and `null` removes it. It works on text ranges, text frames, and table cells. The automation `setFont` and `readFont` operations add the `highlightColor` field.
