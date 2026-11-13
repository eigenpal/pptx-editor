---
'@pptx-editor.dev/core': minor
---

Draw group effects (`p:grpSpPr/a:effectLst`). A group's outer shadow paints as one shadow of all its shapes together, placed by the group's box. With the Pro drawing renderers, glow, soft edge, reflection, and blur also apply to the whole group. The text of a soft-edged group paints again over the faded picture, so it stays sharp and stays text in a printed PDF. A group's inner shadow draws nothing. `ResolvedGroup` has a new `effects` field.
