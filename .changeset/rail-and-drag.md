---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': patch
'@pptx-editor.dev/vue': patch
---

Long decks open faster: the slide rail and the Slide Sorter paint the thumbnails in view at once, and the others when they scroll into view or when the browser is idle. `renderThumbnail` still paints at once by default; pass `{ defer: true }` to paint as the rail does, and call the function it returns when the element goes. Shape drags stay smooth because the deck changes once, on release, so collaborators see a drag only when it ends. The preview shows the size and text scale that autofit gives on release, and autofit text in a scaled group now fits at the size the group draws it. Escape, a lost window focus, an undo, or a change to a dragged shape or its group cancels the drag; a change to another shape keeps it. An `Editor` implementation must now return a disposer from `renderThumbnail`.
