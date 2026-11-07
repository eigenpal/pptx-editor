---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/react': patch
'@pptx-editor.dev/vue': patch
'@pptx-editor.dev/i18n': patch
---

Large decks no longer freeze the editor. A slide of thousands of shapes opens and edits without a long pause, and clicks, keys, and toolbar updates on large slides answer in a few frames.

- A slide with more than 2,000 shapes paints over several frames. Until a shape paints again, its earlier picture stays on screen.
- The thumbnail of a large current slide paints when the editor is idle. A thumbnail repaints only the shapes that changed, and a large slide fills its thumbnail over several idle periods.
- On a selection of more than 250 shapes, `can()` checks only the editing mode for the commands that arrange shapes, and checks other commands on the first 250 shapes first. It checks only the mode for `duplicate` on a selection of more than 20,000 XML nodes. The `applyLayout` and `resetSlides` commands also count a named slide of more than 20,000 XML nodes as large. `exec()` makes the full check.
- A text format check reads at most 64 runs or paragraphs, with the same answer as a full check. List buttons are exact again on the key that starts a list. On a text box of 50,000 runs, Select All and Escape answer in one frame.
- A transaction writes in place into the elements it made. Ops on the runs of one long paragraph no longer copy its run list once per run.
- Clicks, double-clicks, and Select All on a slide of 10,000 shapes no longer lay out the slide again. The pointer's cursor sits on a layer over the slide, so a cursor change no longer restyles every shape.
- Shrink-on-overflow fit lays out a long paragraph once per step only up to the box, and finds its font slots and line breaks once.
- Placeholder matching, Duplicate, Delete, and Group stay linear on slides of thousands of shapes.
- A space after a long run of spaces or a long word no longer slows AutoCorrect.
- The missing-fonts notice names up to four families. With five or more, it names three, then "and N more".
- Bring to Front, Send to Back, Bring Forward, and Send Backward move thousands of shapes in one pass. Bullets, Numbering, Remove Bullet, and Bullet Format resolve the slide once for all selected shapes. Auto-fit, Columns, Text Direction, Paste Format, and a resize of many shapes re-fit them in one pass.
