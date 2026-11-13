---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/i18n': patch
---

Set the transparency of a shape's solid fill from the Format pane, with a slider and a percent field under Fill color. `ShapeFormatDetail.solid` reads the fill color and opacity, and the `{ kind: 'solid', opacity }` patch writes it. A new fill color now keeps the current transparency.
