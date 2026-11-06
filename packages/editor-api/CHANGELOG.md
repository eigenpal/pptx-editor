# @pptx-editor.dev/editor-api

## 0.2.0

### Patch Changes

- 5a03952: Published packages no longer ship source maps, so a stack trace points into `dist/`. Every JavaScript file of `pro`, `pptx-to-pdf`, and `editor-api` starts with its license header, and these packages are minified. The preset shape geometry in `core` carries the Ecma copyright notice in a legal comment, so bundlers that keep legal comments carry the notice into your app bundle.
- 60ed234: A long deck with a damaged slide or notes slide is refused at once, wherever the slide is read. After edits, the deck stays open read-only with `snapshot().readOnlyReason`, a status bar message says why, its other slides go on opening, and commands, undo, save, and automation writes and exports are refused. `snapshot().isOpening` is true until the first slide shows, `load` accepts any `startSlide` number, and the slide rail mounts in steps of about 8 ms.
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

- 412d750: Add `@pptx-editor.dev/editor-api`, an Office.js-shaped API that edits slides, shapes, text, and tables on a server or in an open editor, over the new DOM-free automation host in `@pptx-editor.dev/core/automation`.
- 412d750: Bind shapes by id with `Presentation.bindings`, `BindingCollection`, and `Binding`; the runtime keeps bindings in memory and never saves them in the file.
- 412d750: Use the rest of the Office.js request context: `sync` answers a pass-through value, and `loadRecursive`, `trace` with `error.traceMessages`, `debugInfo`, and `application` are available.
- 412d750: Draw `getImageAsBase64` pictures on a server with `createServer({ rasterize })`: the engine paints the SVG with the editor's own renderer and your function turns it into a PNG.
- 412d750: Read and delete the hyperlinks of a text range through `TextRange.hyperlinks`, a `HyperlinkScopedCollection` of the text links that overlap the range.

### Patch Changes

- 412d750: A binding whose shape or slide is deleted no longer binds a new shape that takes the same id; an undo that restores the shape makes it work again.
- 412d750: A sync that splits before a read is now one undo step on a live editor: automation batches that carry the same `historyGroup` share one history entry.
- 412d750: A server picture now sends the rasterizer only the faces that may leave the engine, as the SVG export does.
- 412d750: Font face ids are now unique for each open presentation, so a rasterizer's font registry never mixes the faces of two presentations.
- 412d750: A call with an argument that the API does not accept now refuses at the next sync, as in Office.js, so no call of that batch applies.
- 412d750: Every argument refusal now says what to pass, for example "Use '#RRGGBB' or a named color" or "Use a font size from 1 to 4000 points".
- 412d750: A refused enum value or option name now suggests the accepted value that holds its words, such as `ArabicNumeralPeriod` for `ArabicPeriod`, and not only a near typo.
- 412d750: A refused group of a placeholder now gives the diagnostic `edit.group-placeholder` in `debugInfo.innerError`.
- 412d750: `getSelectedShapes().group()` now groups the selected shapes without a load of the selection first, as in Office.js.
- 412d750: The reference for `ParagraphFormat.indentLevel` now says that levels count from 1, as the property reads and writes them.
- 412d750: A wrong name in a collection load path, such as `load('itmes/name')`, now throws at the call with the nearest name.
- 412d750: Every read and lookup now sees the writes queued before it in the same `sync()`, such as `getItemOrNullObject` after an `add` or `getSelectedTextRange()` after `setSelected()`.
- 412d750: Every refusal of the engine now gives a diagnostic with a fix hint in `debugInfo.innerError`, including edit refusals such as an uneven split, a table that is too large, or a shape with no fill.
- 412d750: A call on an object that a refused sync would have created now refuses with the diagnostic `sync.object-from-refused-sync`, which says to create it again.
- 412d750: Proxies from a refused sync stay usable, as in Office.js: only the refused sync's calls are dropped, and a retry with the same objects works while they exist.
- 412d750: The development warning for a run that ends with unsent writes now names each argument error that the missing sync would have given.
