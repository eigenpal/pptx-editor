---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/pro': patch
---

Decks with hostile pictures, huge effects, and very long text no longer freeze the editor.

- An SVG picture draws as vector art only inside a subset whose drawing cost the editor can bound, and draws its raster alternative otherwise. The subset holds shapes, paths, text, groups, gradients, clip paths, `<use>` and `<symbol>`, style sheets, and raster `data:` pictures of up to 40 megapixels. Filters, masks, patterns, markers, animation, `<foreignObject>`, `<switch>`, links, dash arrays, external URLs, `data:` URIs other than raster pictures and fonts, and style sheets with escapes, custom properties, or at-rules other than `@font-face` and `@media` draw the raster alternative. The cost counts elements, path vertices, text characters, layers, picture pixels, and the share of the picture each shape covers, with each `<use>` expanded where it sits. A part of more than 2 MB draws the raster alternative. The check takes time linear in the part. Before, an SVG picture whose `<use>` links expanded to billions of shapes stopped the page.
- Pictures decode off the frame (`decoding="async"`). A slide of many large pictures no longer holds a frame for seconds, and each picture shows when it is decoded.
- A glow larger than 200 pt draws as a 200 pt glow, and a soft edge shrinks the outline by at most 100 pt. The outline grows and shrinks only near the shape, not over the whole blur region. A soft edge that removes the whole shape paints nothing at once. Before, a glow or soft edge of a few inches stopped the page for seconds on each paint.
- Stacked text (`wordArtVert`) and upright East Asian characters in vertical text paint one text element for each run, not one for each character. The pixels stay the same.
- A text body sets its kerning and ligature style once on its group, not on each text element. The browser restyles long text faster.
