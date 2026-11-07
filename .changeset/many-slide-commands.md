---
'@pptx-editor.dev/core': patch
---

Delete and duplicate many slides at once without freezing the page. With every slide of a 3,000-slide deck selected, **Delete** no longer takes most of a minute, and **Duplicate** no longer stops the page. Each command reads the deck once and writes the slide list once, whatever the number of selected slides.
