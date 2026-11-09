# @pptx-editor.dev/pro

## 0.3.0

### Minor Changes

- 9d7d60b: Keep text that peers type at the same time as deletes, joins, and formats. The room format is now 6, so update every peer of a room together. Older replicas refuse the room with `schema-version-mismatch`.
- 9d7d60b: Add **New comment** to the canvas context menus, with `registerCommentDraft` and `requestCommentDraft`. Comment commands take `expect` and can fail with `errors.commentChanged`, `errors.commentAmbiguous`, or `errors.noComment`. Threads report `identity`; send it in `expect`. Ids of repeated comments, comments with `~`, and some legacy comments change, so read ids from the current snapshot.
- d3545c2: Highlight commented text and elements in each author's color. Add text comments (`range` on `addComment`, the `text` anchor), `getReviewAuthors()`, and `useReviewAuthors()`. Remove `commentAuthorSlot` and `COMMENT_AUTHOR_SLOTS`.
- 9d7d60b: Add slide comments with the comment commands, `readCommentThreads`, the `pane` option of `reviewModule`, and `PptxEditorReview`, `useReview`, `useReviewOf`, `useReviewItem`, and `useReviewAuthor` from `@pptx-editor.dev/pro/react` and `@pptx-editor.dev/pro/vue`. The room format is now 7, so update every peer of a room together.
- 9d7d60b: Draw 3-D shapes with `drawingModule` from `@pptx-editor.dev/pro/drawing`, and preset shadows in every edition. Set bevels, 3-D rotation, and blur with `setShapeEffect` or the Shape Effects menu.

### Patch Changes

- f448820: Decks with hostile SVG pictures, very large glow or soft-edge effects, and very long text no longer freeze the editor. Complex SVG pictures draw their raster alternative.
- 9d7d60b: Bound the Yjs session's reads of removed text, and refuse a room format that arrives after attach with `schema-version-mismatch`.
- 9d7d60b: Read removed runs and paragraphs the same way on every Yjs peer, with a step limit.
- 9d7d60b: Fix 3-D shapes with zero height, turned faces, and inherited cameras.
- Updated dependencies: @pptx-editor.dev/core@0.3.0, @pptx-editor.dev/react@0.3.0, @pptx-editor.dev/vue@0.3.0

## 0.2.0

### Minor Changes

- a3afcdd: Charts now draw from the values cached in their chart parts when `@pptx-editor.dev/pro/drawing` is installed (`drawingModule({ licenseKey })` or `installDrawing()`). Bar, column, line, area, scatter, pie, doughnut, and radar charts draw with titles, legends, axes, gridlines, data labels, and number formats. Other chart types, and every chart without the module, keep the placeholder. `@pptx-editor.dev/pptx-to-pdf` draws charts in its PDFs, and reports `frame.chart` only for the chart types the module leaves out. The core adds the `chart` renderer key to `installDrawingRenderers`, the resolved `chart` part of a chart frame, the laid-out `chart` items of a frame, and `drawingCapabilities().chart`.
- c292536: Glow, inner shadow, reflection, and soft edge now paint through the new `@pptx-editor.dev/pro/drawing` entry, as in DOCX Editor. Register `drawingModule({ licenseKey })` in `modules`, or call `installDrawing()` for headless rendering. Without it, a shape or picture paints without these four effects; outer shadows still paint, and the effect commands and save keep every effect. `@pptx-editor.dev/pptx-to-pdf` always installs the module, and its PDFs now draw reflections with their fade (`effect.reflection-blur` reports the blur that is left out). `@pptx-editor.dev/core/output` no longer exports the internal `glowSpread` and `innerShadowShift`.

### Patch Changes

- 98b85e2: Right chart legends, top and bottom legends of one row, and radar labels keep the reference spacing at text sizes from 9 to 20 pt in frames from 5.5 to 10 inches wide, not only at 12 pt in a 10-inch frame.
- 5a03952: Published packages no longer ship source maps, so a stack trace points into `dist/`. Every JavaScript file of `pro`, `pptx-to-pdf`, and `editor-api` starts with its license header, and these packages are minified. The preset shape geometry in `core` carries the Ecma copyright notice in a legal comment, so bundlers that keep legal comments carry the notice into your app bundle.
- cd341f2: Each package now accepts only the same minor version of `@pptx-editor.dev/core` as a peer, for example `~0.1.0`. The links in the package READMEs now open from npm.
- Updated dependencies [4594311]
- Updated dependencies [0057839]
- Updated dependencies [db9407d]
- Updated dependencies [7d4b0be]
- Updated dependencies [5a11387]
- Updated dependencies [bfca3d6]
- Updated dependencies [159eb87]
- Updated dependencies [0653436]
- Updated dependencies [5a03952]
- Updated dependencies [136b3b5]
- Updated dependencies [6bed74e]
- Updated dependencies [0593988]
- Updated dependencies [b906153]
- Updated dependencies [a3afcdd]
- Updated dependencies [c292536]
- Updated dependencies [c292536]
- Updated dependencies [60ed234]
- Updated dependencies [60ed234]
- Updated dependencies [84256e4]
- Updated dependencies [0653436]
- Updated dependencies [60c621b]
- Updated dependencies [6bed74e]
- Updated dependencies [db11b83]
- Updated dependencies [14a9bc9]
- Updated dependencies [1358a9d]
- Updated dependencies [aa99dad]
- Updated dependencies [cd341f2]
  - @pptx-editor.dev/core@0.2.0

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
