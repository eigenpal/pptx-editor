---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/react': patch
'@pptx-editor.dev/vue': patch
---

Large slides open without freezing the page. No main-thread task of the open takes longer than about 500 ms, where some took 1.5 seconds.

- A text that overflows its slide paints only the lines that can show. A thumbnail paints the lines on its slide. The editing surface paints as far past the slide as the window reaches at the zoom, and paints further when the window grows. A table paints the cells that can show. A shape with a shadow, a reflection, or a glow paints every line.
- `PaintOptions.view` sets the part of slide space that can show. Without it, everything paints, as before.
- A group of thousands of shapes paints over several frames, as thousands of shapes do. A thumbnail counts the shapes in groups and the cells of tables when it decides that a slide is large.
- The open parses the deck, resolves the first slide, lays it out, and paints it in separate tasks when each step is long. A small deck still opens in one task.
- The toolbar no longer walks the whole slide to decide whether **Insert Text Box** and the slide-wide commands can run.
- A slide with many `mc:AlternateContent` elements resolves each one once.
- Text paints with fewer SVG elements: one group for each text body, with the text style on it.
