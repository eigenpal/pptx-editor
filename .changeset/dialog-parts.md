---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/i18n': minor
---

Dialogs expose DOCX Editor's named parts (`Header`, `Title`, `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`, `Reset`) over a draft, with `usePageSetupDialog` and `useBackgroundDialog`: Format Background (the menu row is now "Format Background…") applies its choice on OK as one undo step and Cancel discards it, Page setup reports the appearance through `onThemeChange` on OK instead of on each click and has OK and Cancel instead of Close, and the Vue adapter gains the `popups` registry and the Page setup dialog.
