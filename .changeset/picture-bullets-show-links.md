---
'@pptx-editor.dev/core': minor
'@pptx-editor.dev/react': minor
'@pptx-editor.dev/vue': minor
'@pptx-editor.dev/i18n': minor
'@pptx-editor.dev/pptx-to-pdf': minor
'@pptx-editor.dev/editor-api': patch
---

Draw picture bullets (`a:buBlip`) in the editor and the PDF export: the picture stands on the baseline, 0.7 times the bullet size high, with its own aspect. `setBulletFormat` takes `{ kind: 'picture', picture }` with PNG, JPEG, or GIF bytes, and the Bullets and Numbering dialog has a **Choose picture…** field. `setLink` links to a custom show (`#custom-show-<id>`, with `?return=true` to come back after the show) and to another file by a checked relative path (`./report.pdf`). `EditorSnapshot.customShows` lists the deck's custom shows, and the link popup offers them. The editor and the slide show never open a file link or a program link, and `setLink` never writes a program link.
