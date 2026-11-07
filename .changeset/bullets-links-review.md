---
'@pptx-editor.dev/core': patch
'@pptx-editor.dev/react': patch
'@pptx-editor.dev/vue': patch
'@pptx-editor.dev/i18n': patch
'@pptx-editor.dev/editor-api': patch
---

Check file links more strictly: a link names only a document, picture, sound, or video file, and a name that ends with a dot or a space, or names a device such as `CON`, is refused. **Other file** adds `./` to a bare name, so the link stays a file link. Copy and paste keeps a safe file link, and automation reads its address. A table cell draws a picture bullet from the default text style. The Bullets and Numbering dialog says why it refuses a picture, and `setBulletFormat` refuses a picture over 25 million pixels.
