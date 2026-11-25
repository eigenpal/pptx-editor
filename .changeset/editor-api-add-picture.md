---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add `ShapeCollection.addPicture` to the editing API, an SDK extension. It adds a picture shape from PNG, JPEG, or GIF bytes of at most 50 MiB and answers the new `Shape`. The core automation operation `addShape` takes the new `picture` shape kind, which uses the Insert Picture edit.
