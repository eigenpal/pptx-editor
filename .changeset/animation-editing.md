---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/editor-api': patch
'@pptx-editor.dev/i18n': minor
---

Add animation and transition editing. `addAnimation`, `removeAnimation`, `moveAnimation`, and `setAnimationTiming` add, remove, reorder, and time the entrance, emphasis, and exit effects of a slide's main sequence, from 13 presets. `setTransition` sets or clears a fade, push, wipe, split, or cut transition with its duration and advance timing, and keeps a transition kind that it does not write when only the timing changes. `EditorSnapshot.animations` lists the effects and the transition of the current slide, and `showAnimationOrder` tags animated shapes on the canvas with their click steps. The side panel of the React and Vue adapters has an **Animations** tab with these edits. A save keeps one transition on each slide, before its timing. The editor does not play effects or transitions.
