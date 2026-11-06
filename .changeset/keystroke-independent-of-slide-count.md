---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/react': patch
'@pptx-editor.dev/vue': patch
---

A keystroke in a deck of 5,000 slides now costs about the same as in a short deck. An edit records only the parts it changes and copies no list of every part. The slide summaries, the font list, and the font notice update from the changed parts only. The slide rail renders only the items whose slide changed, a thumbnail out of view takes no layout or paint, and the canvas reads its fit zoom only when the slide or canvas size changes.
