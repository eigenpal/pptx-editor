---
'@pptx-editor.dev/core': patch
---

Bound deferred font composition to prevent cyclic resolver answers from blocking browser input. Release stalled browser font registration waits and clean up partially registered faces after failures.
