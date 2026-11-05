---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
---

Match the editor chrome to your app with public API only. `--doc-font-ui` now sets the font of the whole chrome, `--doc-radius-surface`, `--doc-radius-control`, and `--doc-radius-pill` (the zoom pill and chips) set its corners, the dialog panel falls back to `--doc-radius-surface`, dialog controls read `--doc-dialog-control-radius`, then `--doc-radius-control`, and the primary dialog button reads `--doc-primary-hover`. The slide rail's action row has named parts (`SlideRail.AddSlide`, `ChooseLayout`, `Duplicate`, `MoveEarlier`, `MoveLater`, `Delete`, `Collapse`, and `Action`) with `hidden`, `icon`, `className`, `asChild`, and `preset={false}`. A horizontal rail sets its action columns to half the buttons that show, rounded up, so a rail with `toggle={false}` shows its six buttons as three over three. The notes field fills the height that a host gives the notes pane. A focused rail slide rings its thumbnail, and the custom color field is as tall as Apply beside it. A context menu that is too tall for the room above or below the pointer now stays at the pointer and scrolls, so a slide menu opened in a rail at the bottom of a short window no longer covers the toolbar.
