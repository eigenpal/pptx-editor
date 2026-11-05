---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/i18n': minor
---

`PptxEditor.Menu` and `PptxEditor.ContextMenu` render the menu bar and the canvas menus from a core registry (`CHROME_MENUS`, `CANVAS_CONTEXT_MENUS`) with `slotId` rows, `Group`, `Separator`, and `Submenu`, the `popups` registry gains `hyperlink`, `find`, `contextMenu`, and `railContextMenu`, `PptxEditor.Root` now hosts the link popup, and `useEditorCommand` runs dialog, clipboard, find, link, and slide show slots (in Vue, the clipboard and pictures; the others show as disabled with the reason).
