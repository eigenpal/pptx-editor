---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
---

Paste shapes copied in other presentation applications as editable shapes in a desktop host. Pass `readClipboardDrawing` to read the DrawingML clipboard package from the system clipboard; the editor pastes its shapes, pictures, connectors, groups, and tables at their source positions. In a browser, the paste stays a picture. `drawingClipboardPayload` converts a package for the `paste` command, and `DRAWING_CLIPBOARD_TYPES` names the clipboard types.
