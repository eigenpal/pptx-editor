---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/i18n': patch
'@pptx-editor.dev/react': patch
'@pptx-editor.dev/vue': patch
---

The editor chrome is easier to use with assistive technology and high contrast themes. `PptxEditor.Viewport` is a landmark region named "Presentation editor"; its new `label` prop gives another name, and `label={false}` removes the region. Slides in the rail and the slide sorter take their name ("Slide 3", "Slide 3 (hidden)") from hidden text, and the thumbnail art is hidden from screen readers. With a forced-colors theme, selected slides, pressed and checked toggles, focus rings, disabled controls, and card edges stay visible in system colors. Color swatches are 24 by 24 pixels.
