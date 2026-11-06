---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pptx-to-pdf': patch
---

Load the packaged substitute fonts under Node, so headless layout and PDF export measure and draw text with real font bytes. `fontReport.substituteFailures` and `fontResolution.substituteFailures` name each face that fails to load. The PDF export now draws single and double underline and strikethrough, and its `font.missing` diagnostic names the font.
