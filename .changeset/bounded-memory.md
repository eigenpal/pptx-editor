---
'@pptx-editor.dev/core': patch
---

Keep memory level in long sessions. The rail and the slide sorter keep at most 200 thumbnails painted, and paint the others again when they scroll into view. The text measurements that the editor keeps are limited in size, so typing in a long paragraph no longer keeps a copy of each version.
