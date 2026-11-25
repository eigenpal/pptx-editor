---
'@pptx-editor.dev/editor-api': minor
'@pptx-editor.dev/core': minor
---

Add `ShapeCollection.addPicture` to the editing API, an SDK extension. It adds a picture shape from PNG, JPEG, or GIF bytes of at most 50 MiB and returns the new `Shape`. The automation operation `addShape` takes a new `picture` kind. Picture bytes over 50 MiB now refuse with `InvalidArgument` in every picture write.
