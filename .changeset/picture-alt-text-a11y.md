---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/i18n': patch
---

A screen reader now hears the alt text of a picture on the slide. A described picture is an image named by its description. A screen reader skips a decorative picture and a picture with no description. When Tab selects an object, the screen reader hears its kind, its name, and its description. The selection overlay is hidden from assistive technology. In the slide show, a screen reader reads the slide text and the alt text of its pictures. `PaintOptions` takes an optional `altText` port. Without it, the paint has no ARIA.
