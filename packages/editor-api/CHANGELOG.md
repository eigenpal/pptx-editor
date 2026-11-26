# @pptx-editor.dev/editor-api

## 0.8.0

### Minor Changes

- eb5a5b9: Add the SDK extension `ShapeCollection.addChart(type, data, options?)`. It adds a chart with an embedded workbook and returns the new shape. The automation `addShape` operation takes a `chart` kind. A chart edit now refuses a chart part over 16 MiB before it writes. `Chart.setData` now refuses a store rejection with `NotSupported`, not `InvalidArgument`.
- 74e6cf7: Add `ShapeCollection.addPicture` to the editing API, an SDK extension. It adds a picture shape from PNG, JPEG, or GIF bytes of at most 50 MiB. It returns the new `Shape`. The automation operation `addShape` takes a new `picture` kind. Picture bytes over 50 MiB now refuse with `InvalidArgument` in every picture write.
- 894ae1f: Add `Shape.chart` to the editing API, an SDK extension. `Chart` reads the type, title, categories, series, and workbook state of a chart, and `setData` replaces its data. The new automation operations `readChart` and `setChartData` use the editor's chart edit, so the caches and an embedded workbook change together.
- face20e: Add `Slide.comments` to the editing API, an SDK extension. It reads, adds, replies to, edits, resolves, and deletes slide comments in both formats, through `Comment`, `CommentReply`, and their collections. The new `author` option of `createServer` and `createBrowser` signs new comments. The automation host adds the comment operations, `BrowserAutomationOptions`, and `commentAuthorName`.
- c5ed475: Add `Slide.headerFooter` and `Presentation.applyHeaderFooter()` to the editing API, as SDK extensions. They show or hide the slide number, the date, and the footer on one slide or on all slides, and can skip title slides. The automation protocol adds `readHeaderFooter` and `setHeaderFooter`.
- d06ae7b: Add `Presentation.replaceAll(searchText, replacementText, options)` and `Slide.replaceAll` to the editing API, with `SearchOptions` (`matchCase`, `matchWholeWord`). It replaces text in shapes, table cells, and notes. It keeps the font and returns the count. The automation host adds the `replaceText` operation. Find no longer matches a field through U+FFFC.
- 477012a: Add the SDK extension `Presentation.sections`. Read each section's id, name, position, and slide ids. Add a section at a slide, rename it, move it with its slides, and delete it with or without its slides. Each sync is one undo step. The core automation protocol adds section operations and the `section` handle kind.
- bdfeefb: Add the SDK extension `Slide.duplicate()`. It copies a slide directly after its source in the same presentation and returns the new `Slide`, which works in the same batch. The copy keeps the layout, master, and pictures, and gets its own notes, charts, and tags. The core automation protocol adds the `duplicateSlide` operation.
- 8a1753a: Add the SDK extension `Slide.speakerNotes`. Load it to read the notes as plain text, and set it to replace them, one paragraph for each line. A write to a slide with no notes creates the notes slide from the notes master. The core automation protocol adds the `readSlideNotes` and `setSlideNotes` operations.
- 07fefab: Add the SDK extension `TextRange.search(searchText, options?)`, in the shape of `Range.search` of the DOCX Editor API. It returns a `TextRangeCollection` of the matches inside the range, with `items`, `getCount`, `getItemAt`, `getFirst`, `getFirstOrNullObject`, and `getLast`. It uses the rules of `replaceAll`. The automation protocol adds the `findText` operation.
- 658f45e: Add SDK extensions on `ParagraphFormat` for line spacing, space before and after, and indents. Add the `ParagraphLineSpacingRule` enum and `TableCell.paragraphFormat`. A `150%` spacing value now reads as 1.5 lines. Reads return the value in force, or `null` for a mixed range. The automation `setParagraph` and `readParagraph` operations add the matching fields.
- 15de8f2: Add `ShapeFont.highlightColor` to the editing API, an SDK extension. It reads the text highlight as `#RRGGBB`, or `null` for no highlight and for a mixed range. A color write sets `a:highlight`, and `null` removes it. It works on text ranges, text frames, and table cells. The automation `setFont` and `readFont` operations add the `highlightColor` field.
- 26fff35: Add `Slide.hidden` to the editing API, an SDK extension. Load it to read whether a slide is hidden, and set it to hide or show the slide. The write changes `p:sld/@show` through the new automation operation `setSlideHidden`, which the `setSlidesHidden` editor command also uses.
- c03c16c: Add `Slide.transition` to the editing API, an SDK extension. `SlideTransition` reads and writes the effect, `durationMs`, the advance settings, `direction`, and `throughBlack`. `set()` writes several in one step, and `remove()` clears the transition. The new automation operations `readTransition` and `setTransition` use the engine's transition edit.
- ae05398: Add the `TextFrame.getTextFit()` SDK extension. It reports whether text overflows its shape, from the layout that `getImageAsBase64()` draws, with `overflows`, `textHeight`, `availableHeight`, and `substitutedFonts`. It works on a server runtime with no canvas. The automation protocol adds the `readTextFit` operation.

### Patch Changes

- 9e9c76f: `editor-api` declares `@napi-rs/canvas` `^1.0.12` as an optional peer dependency. Earlier releases draw the SVG of a `getImageAsBase64()` picture without its `<image>` elements. The slide pictures are then missing with no error. The runtime guide states this minimum. If your package manager reports a peer conflict, update `@napi-rs/canvas`.
- Updated dependencies: @pptx-editor.dev/core@0.8.0

## 0.7.0

### Patch Changes

- Updated dependencies: @pptx-editor.dev/core@0.7.0

## 0.6.0

### Patch Changes

- Updated dependencies: @pptx-editor.dev/core@0.6.0

## 0.5.0

### Minor Changes

- ef380c0: Set the transparency of a shape's solid fill from the Format pane, with a slider and a percent field under Fill color. `ShapeFormatDetail.solid` reads the fill color and opacity, and the `{ kind: 'solid', opacity }` patch writes it. A new fill color now keeps the current transparency.

### Patch Changes

- Updated dependencies: @pptx-editor.dev/core@0.5.0

## 0.4.0

### Minor Changes

- 91482da: Add customizable selection-aware inspectors for React and Vue, with safe drafts and matching automation APIs. Edit line widths, endpoints, gradients, shadows, and reflections while preserving unedited XML. Keep picture pickers bound to their original document and selection. Line endpoint commands now refuse mixed line and non-line selections without changing any object.

### Patch Changes

- Updated dependencies: @pptx-editor.dev/core@0.4.0

## 0.3.1

No changes in this release.

## 0.3.0

### Minor Changes

- 9d7d60b: Type in diagram shapes, and add or delete diagram nodes with `addDiagramShape` and `deleteDiagramShape`. Add `EditorSnapshot.diagram`.
- 9d7d60b: Add slide comments with the comment commands, `readCommentThreads`, the `pane` option of `reviewModule`, and `PptxEditorReview`, `useReview`, `useReviewOf`, `useReviewItem`, and `useReviewAuthor` from `@pptx-editor.dev/pro/react` and `@pptx-editor.dev/pro/vue`. The room format is now 7, so update every peer of a room together.

### Patch Changes

- 9d7d60b: Add animation and transition editing with `addAnimation`, `removeAnimation`, `moveAnimation`, `setAnimationTiming`, `setTransition`, `EditorSnapshot.animations`, `showAnimationOrder`, and the **Animations** side panel tab. The editor does not play effects or transitions.
- 9d7d60b: Check file links more strictly, keep safe file links through copy and paste, and refuse picture bullets over 25 million pixels.
- 7046dbf: Edit diagram text in the **Diagram** side panel tab, opened with **Edit text…**. Add `setDiagramText`, and fix context menu rows that open a side panel tab.
- 30dc48e: Add the `setFooter` and `setDateTime` commands, and `EditorSnapshot.footer` and `EditorSnapshot.dateTime`. `setSlideNumbers` takes `indices` and `hideOnTitleSlides`, and refuses `show: false` when no slide can have a slide number.
- 9d7d60b: Draw picture bullets, and set them with `setBulletFormat`. `setLink` links to custom shows and to relative file paths. Add `EditorSnapshot.customShows`.
- Updated dependencies: @pptx-editor.dev/core@0.3.0

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
