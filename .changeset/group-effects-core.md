---
'@pptx-editor.dev/core': minor
---

Draw group effects (`p:grpSpPr/a:effectLst`) on the whole group: one outer shadow for all its shapes, and with the Pro drawing renderers, glow, soft edge, reflection, and blur. The text of shapes under a group soft edge paints once over the faded picture, sharp and at its own opacity; a table's text fades with it. A group's inner shadow draws nothing. Effects of groups nested inside two groups that draw effects, and of groups of more than 1,000 elements, do not draw and are reported. `ResolvedGroup` has a new `effects` field.
