---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add `Slide.transition` to the editing API, an SDK extension. `SlideTransition` reads and writes the effect (`SlideTransitionEffect`), `durationMs`, `advanceOnClick`, and `advanceAfterMs`, and `set()` writes several in one step. The new automation operations `readTransition` and `setTransition` use the engine's transition edit.
