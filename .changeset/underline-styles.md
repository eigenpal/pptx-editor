---
'@pptx-editor.dev/core': patch
---

Draw every underline style: heavy, dotted, dashed, long dashed, dot-dash, dot-dot-dash, and wavy underlines, with their heavy and double forms. An underline outline (`a:uLn`, `a:uLnTx`) strokes the underline with its width, dash, caps, and fill, and the new `RunProps.underlineLine` field holds it. Underlined runs on one line share one underline position and thickness, and spaces between runs are underlined. Small caps in a superscript or subscript keep the script size.
