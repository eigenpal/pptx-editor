---
'@pptx-editor.dev/core': patch
---

A font that never loads no longer holds the page.

- A font resolver that answers in a cycle no longer blocks input.
- A browser font load that never settles holds the open, and each PNG export of a slide, for at most 2 seconds.
- Faces that were registered before a failure are removed.
