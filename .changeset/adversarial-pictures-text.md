---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pro': patch
---

Decks with hostile pictures, huge effects, and very long text no longer freeze the editor.

- An SVG picture draws only when it is safe to draw. The editor counts, as an upper bound, the elements the picture draws with every link expanded: `<use>` copies, masks, patterns, clip paths, and filters for each element that links to them, and markers at each vertex of the paths they can reach. A repeated id counts as its largest element. The count can be at most 200,000, with no `<use>` cycle and no chain of more than 16 links. An `xml-stylesheet` instruction, a `<foreignObject>`, an internal document type declaration, and a `data:` URI that can hold a document (anything but a raster picture or a font) are refused. Otherwise the picture draws its raster alternative, or the placeholder. Before, an SVG picture whose `<use>` links expanded to billions of shapes stopped the page.
- Pictures decode off the frame (`decoding="async"`). A slide of many large pictures no longer holds a frame for seconds, and each picture shows when it is decoded.
- A glow larger than 200 pt draws as a 200 pt glow, and a soft edge shrinks the outline by at most 100 pt. The outline grows and shrinks only near the shape, not over the whole blur region. A soft edge that removes the whole shape paints nothing at once. Before, a glow or soft edge of a few inches stopped the page for seconds on each paint.
- Stacked text (`wordArtVert`) and upright East Asian characters in vertical text paint one text element for each run, not one for each character. The pixels stay the same.
- A text body sets its kerning and ligature style once on its group, not on each text element. The browser restyles long text faster.
