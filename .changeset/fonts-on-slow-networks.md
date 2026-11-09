---
'@pptx-editor.dev/core': patch
---

Show real fonts sooner on a slow network. The packaged substitute fonts now download next to the text engine, not after it. The publish that reports a deck as ready now reports late fonts as `resolving`, so a host can keep its loading screen until the text is measured with the real fonts.
