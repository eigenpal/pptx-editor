---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
---

`moveSlides` and `moveSlide` take an optional `section`, so a drop below a section header or a Ctrl/Cmd+Arrow step across one puts the slides into that section, and the saved file keeps one slide list with a valid id for every slide after two collaborators add the first slide at the same time. Behavior change: `moveSlide` now follows the rule of `moveSlides`, so a slide moved up to the first place of a section joins the section before it instead of that section; pass the section's id as `section` to keep the earlier result.
