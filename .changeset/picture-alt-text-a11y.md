---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/i18n': patch
---

A screen reader now hears the alt text of a picture on the slide. A picture with a description is an image named by it, and a decorative picture or a picture with no description is hidden. When Tab selects an object, the screen reader hears its kind, its name, and its description. The selection overlay is hidden from assistive technology. The slide show is no longer one unnamed image: a screen reader reads the slide text and the alt text of its pictures. `PaintOptions` takes an optional `altText` port; without it, the paint has no ARIA.
