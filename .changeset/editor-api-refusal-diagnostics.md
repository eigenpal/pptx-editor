---
'@pptx-editor.dev/editor-api': patch
---

Every refusal of the engine now gives a diagnostic with a fix hint in `debugInfo.innerError`, including edit refusals such as an uneven split, a table that is too large, or a shape with no fill.
