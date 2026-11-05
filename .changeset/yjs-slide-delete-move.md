---
'@pptx-editor.dev/pro': patch
---

A slide that one collaborator deletes while another moves it stays deleted on every replica, an undo of the delete brings it back in one section, and the saved file no longer names a slide that is gone or a slide twice. A copied slide jump to the deleted slide loses its link on every replica that gets the delete; a copy that arrives before the undo of its target slide loses its link too.
