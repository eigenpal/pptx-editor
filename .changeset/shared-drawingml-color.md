---
'@pptx-editor.dev/core': patch
---

Color transforms now read a percentage written with a `%` sign (`lumMod val="50%"`) and an integer with a leading `+`, so the transform draws the value the file states instead of 0.
