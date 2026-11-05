# @pptx-editor.dev/pro

## 0.1.0

### Minor Changes

- 412d750: The Vue adapter, the pro modules for review and collaboration, and PDF export are now published with the other packages.
- 412d750: In a collaboration session, two collaborators who change the same fill, outline, effect list, background, or text property at once now get one valid element, and the saved file opens without a repair prompt.

### Patch Changes

- 412d750: Two peers who type at once into a table cell with no text body both keep their text, a blank second text body in a cell no longer adds an empty line, and the automation API reads and edits the text of every text body of such a cell.
- 412d750: Content that two collaborators move at the same time, such as a paragraph in a table cell, now shows the same on every replica after a later edit deletes the place where it shows.
- 412d750: A character that a peer deleted no longer comes back after another peer presses Enter in the same text.
- 412d750: Two collaborators who format one run at once no longer duplicate its text, a format no longer loses its word when a collaborator joins the paragraph at the same time, and content that one collaborator moves into content that another collaborator deletes at the same time is kept in the file.
- 412d750: In collaboration, text that a peer edits while another peer's format or Enter moves it stays in place, two Enters in one paragraph keep each paragraph's text once, and a deleted run keeps the text a peer typed into it at the same time.
- 412d750: An undo in a collaboration session keeps the text and properties that a collaborator wrote into the undone element, and an undo or redo while collaborators format, split, delete, or press Enter in the same text keeps every character and leaves all replicas alike.
- 412d750: In a collaboration session, two peers who press Backspace at the start of two neighboring paragraphs at the same time now keep the text of all three paragraphs on every replica.
- 412d750: In a collaboration session, a replica that could not apply a peer's change no longer writes older merged table cell text over that change when it edits again.
- 412d750: Clarify package documentation and host responsibilities for collaboration access.
- 412d750: In a collaboration session, two peers that add or copy slides, pictures, links, shapes, tables, or comments at the same time no longer save one id twice.
- 412d750: In a collaboration session, a table whose rows or merged cells two people changed at once, such as a merge while a peer deletes one of its rows, now shows the row below the merge on every replica and saves with valid merge markers; the room format is now 3, so a replica of an earlier build refuses the room.
- 412d750: In a collaboration session, text in a table cell that a peer merges, or that a concurrent row or column delete, insert, split, or merge moves, now shows once on every replica and is no longer lost; the room format is now 4, so a replica of an earlier build refuses the room.
- 412d750: A slide that one collaborator deletes while another moves it stays deleted on every replica, an undo of the delete brings it back in one section, and the saved file no longer names a slide that is gone or a slide twice. A copied slide jump to the deleted slide loses its link on every replica that gets the delete; a copy that arrives before the undo of its target slide loses its link too.
