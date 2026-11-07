---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pro': patch
---

Decks with hostile pictures, huge effects, and very long text no longer freeze the editor.

- An SVG picture draws only when it is safe to draw. Its `<use>` links, and its markers drawn at each vertex of the paths they reach, can expand to at most 200,000 drawn elements, with no link cycle and no chain of more than 16 links. A marker inside marker content is refused. So is a link or a style that loads a `data:` URI that can hold a document (SVG, XML, HTML, or an unknown type), such as a nested SVG document in an `<image>` or an `feImage`. Raster pictures and fonts in `data:` URIs draw. The XML limits apply, and an internal document type declaration is refused. Otherwise the picture draws its raster alternative, or the placeholder. Before, an SVG picture whose `<use>` links expanded to billions of shapes stopped the page.
- Pictures decode off the frame (`decoding="async"`). A slide of many large pictures no longer holds a frame for seconds, and each picture shows when it is decoded.
- A glow larger than 200 pt draws as a 200 pt glow, and a soft edge shrinks the outline by at most 100 pt. The outline grows and shrinks only near the shape, not over the whole blur region. A soft edge that removes the whole shape paints nothing at once. Before, a glow or soft edge of a few inches stopped the page for seconds on each paint.
- Stacked text (`wordArtVert`) and upright East Asian characters in vertical text paint one text element for each run, not one for each character. The pixels stay the same.
- A text body sets its kerning and ligature style once on its group, not on each text element. The browser restyles long text faster.
