---
'@pptx-editor.dev/core': patch
---

Copy and cut slides in the slide list again. **Copy** in the slide list threw an error and put nothing on the clipboard. A slide copy now holds only the copied slides and the parts they use, so copying and pasting one slide of a 3,000-slide deck no longer writes and reads the whole deck. The editor keeps only the file of the last slide copy, so copying many different slides holds one file in memory.
