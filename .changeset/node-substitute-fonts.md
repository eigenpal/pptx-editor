---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pptx-to-pdf': patch
---

Load the packaged substitute fonts under Node, so headless layout and PDF export measure and draw text with real font bytes. `fontReport.substituteFailures` and `fontResolution.substituteFailures` list each face that fails to load as `{ family, file, diagnostic }`, and the PDF export reports each one as a document-wide `font.substitute-failed` diagnostic. The PDF export now draws single and double underline and strikethrough, and its `font.missing` diagnostic names the font. The editor also opens a deck in a DOM without the CSS Font Loading API.
