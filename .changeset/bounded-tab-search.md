---
'@pptx-editor.dev/core': patch
---

Prevent long pauses when a paragraph contains many tab characters and explicit tab stops. Search an indexed stop list without changing tab positions or saved content.

Keep tab advances in the text layout and omit nonpainting tab glyphs from SVG output. This prevents browser stalls on long tab-only runs.
