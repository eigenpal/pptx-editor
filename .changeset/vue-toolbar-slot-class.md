---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/react': patch
'@pptx-editor.dev/vue': patch
---

A plain `class` on the Vue `PptxEditorToolbar.Slot` now goes on its control, as `className` on the React `Toolbar.Slot` does, and a Vue named part with both `class` and `control-class` now adds both names instead of only `control-class`. Both adapters print a development warning when a slot's own content replaces the control, so its class or icon has nothing to apply to, and when a slot has an attribute that never reaches its control, such as `style`, `id`, or `data-*`.
