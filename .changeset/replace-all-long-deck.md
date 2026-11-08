---
'@pptx-editor.dev/core': patch
---

**Replace All** in a long deck takes about half the time: the slide number of each changed slide is read from one list of the slides, not found again in the deck for each slide. An automation batch that adds or moves slides between two text edits reads the new slide numbers, so caret offsets after a `slidenum` field stay correct.
