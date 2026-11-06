---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/react': patch
'@pptx-editor.dev/vue': patch
---

A long deck shows its first slide and takes input before the rest is ready: the other slides open in the background, and `snapshot().pendingSlides` counts them. `load(bytes, { startSlide })` opens another slide first. Edits, undo, Find, and save work before the deck is ready. Decks with up to about 10,000 slides now open, and the slide rail re-renders only the slides that change.
