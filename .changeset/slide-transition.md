---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add `Slide.transition` to the editing API, an SDK extension. `SlideTransition` reads and writes the effect, `durationMs`, the advance settings, `direction`, and `throughBlack`. `set()` writes several in one step, and `remove()` clears the transition. The new automation operations `readTransition` and `setTransition` use the engine's transition edit.
