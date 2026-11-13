---
'@pptx-editor.dev/core': patch
---

Improve the horizontal spacing of equations: scripts use the script-style advances of a math font, and radicals widen with their height. `TextMeasurer.advances` takes a new `mathScript` option, and `ResolvedFace` has a new `mathTable` flag for faces with a `MATH` table.
