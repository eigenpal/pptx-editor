---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/react': patch
'@pptx-editor.dev/vue': patch
'@pptx-editor.dev/i18n': patch
---

Large decks no longer freeze the editor. A slide of thousands of shapes opens and edits without a long pause, and clicks, keys, and toolbar updates on large slides answer in a few frames.

- A slide with more than 2,000 shapes paints over several frames. The shapes painted first stay on screen through an edit.
- The thumbnail of a large current slide paints when the editor is idle. A thumbnail repaints only the shapes that changed, and a large slide fills its thumbnail over several idle periods.
- `can()` checks only the editing mode for a selection of more than 20,000 XML nodes, as it does for more than 250 shapes. The `applyLayout` and `resetSlides` commands also count a slide of more than 20,000 XML nodes as large. `exec()` makes the full check.
- A text format check reads at most 64 runs or paragraphs, with the same answer as a full check. On a text box of 50,000 runs, Select All and Escape answer in one frame.
- A transaction writes in place into the elements it made. Ops on the runs of one long paragraph no longer copy its run list once per run.
- Clicks, double-clicks, and Select All on a slide of 10,000 shapes no longer lay out the slide again. The pointer's cursor sits on a layer over the slide, so a cursor change no longer restyles every shape.
- Shrink-on-overflow fit lays out a long paragraph once per step only up to the box, and finds its font slots and line breaks once.
- Placeholder matching, Duplicate, Delete, Group, and Bring Forward stay linear on slides of thousands of shapes.
- A space after a long run of spaces or a long word no longer slows AutoCorrect.
- The missing-fonts notice names at most four families, then "and N more".
