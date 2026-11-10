# @pptx-editor.dev/react

## 0.4.0

### Minor Changes

- 91482da: Add customizable selection-aware inspectors for React and Vue, with safe drafts and matching automation APIs. Edit line widths, endpoints, gradients, shadows, and reflections while preserving unedited XML. Keep picture pickers bound to their original document and selection. Line endpoint commands now refuse mixed line and non-line selections without changing any object.

### Patch Changes

- Updated dependencies: @pptx-editor.dev/core@0.4.0, @pptx-editor.dev/i18n@0.4.0

## 0.3.1

### Patch Changes

- 8fc9345: Draw each dash style in the outline dash and table border style menus, and show the current dash on the button.
- Updated dependencies: @pptx-editor.dev/i18n@0.3.1

## 0.3.0

### Minor Changes

- 9d7d60b: Add animation and transition editing with `addAnimation`, `removeAnimation`, `moveAnimation`, `setAnimationTiming`, `setTransition`, `EditorSnapshot.animations`, `showAnimationOrder`, and the **Animations** side panel tab. The editor does not play effects or transitions.
- 9d7d60b: Edit the data of bar, column, line, area, pie, doughnut, radar, and scatter charts with `editor.chartData()`, `setChartData`, and the **Chart Data** side panel tab.
- 9d7d60b: Add **New comment** to the canvas context menus, with `registerCommentDraft` and `requestCommentDraft`. Comment commands take `expect` and can fail with `errors.commentChanged`, `errors.commentAmbiguous`, or `errors.noComment`. Threads report `identity`; send it in `expect`. Ids of repeated comments, comments with `~`, and some legacy comments change, so read ids from the current snapshot.
- d3545c2: Highlight commented text and elements in each author's color. Add text comments (`range` on `addComment`, the `text` anchor), `getReviewAuthors()`, and `useReviewAuthors()`. Remove `commentAuthorSlot` and `COMMENT_AUTHOR_SLOTS`.
- 9d7d60b: Type in diagram shapes, and add or delete diagram nodes with `addDiagramShape` and `deleteDiagramShape`. Add `EditorSnapshot.diagram`.
- 7046dbf: Edit diagram text in the **Diagram** side panel tab, opened with **Edit text…**. Add `setDiagramText`, and fix context menu rows that open a side panel tab.
- 9d7d60b: Draw equations on the canvas and in the PDF export. Insert and edit them with `insertEquation`, `setEquation`, and the Equation dialog.
- c47f5ce: Decks no longer stay on the loading spinner: the open waits at most 2 seconds for fonts. Add `editor.fontMeasurement()`, `openFailureReport()`, and the `showCopy` prop of `PptxEditorOpenFailure`. `DocumentOpenFailureReason` adds `blocked`, so update exhaustive maps of it. Large and damaged decks open faster and more safely.
- 8253ddb: Add **Lock aspect ratio** to the Format tab.
- fef01c9: Insert a chart with **Insert** > **Chart** or `insertChart`. Change its type, title, legend, labels, axes, gridlines, and colors in the **Chart** tab or with `setChartType`, `setChartFormat`, and `setChartColor`, and read them with `editor.chartFormat()`. `requestPane` takes `{ focus: false }`.
- 09ddfcc: Show only frequent tasks on the default toolbar. Equation, Layout, picture color and crop options, text direction, and similar controls move to the menus and context menus, and `Toolbar.Slot` still places them. The view switch shows Normal and Slide Sorter; open the slide master view from View > Slide Master.
- 58e17f1: Paste shapes from other presentation applications as editable shapes in desktop hosts with `readClipboardDrawing`. Add `drawingClipboardPayload` and `DRAWING_CLIPBOARD_TYPES`.
- 8253ddb: Fill shapes with preset patterns or pictures with `setShapePattern` and `setShapePicture`, and set pattern or tiled picture slide backgrounds. Add `PatternDialog`.
- 9d7d60b: Draw picture bullets, and set them with `setBulletFormat`. `setLink` links to custom shows and to relative file paths. Add `EditorSnapshot.customShows`.
- 8253ddb: Add the selection pane: the **Selection** side panel tab and the `SelectionPane` component. `SidePanel` takes `tab` and `onTabChange` (`v-model:tab` in Vue). Add `requestPane` and `listenPaneRequests`. The Vue adapter requires Vue 3.5 or later.
- 9d7d60b: Add slide comments with the comment commands, `readCommentThreads`, the `pane` option of `reviewModule`, and `PptxEditorReview`, `useReview`, `useReviewOf`, `useReviewItem`, and `useReviewAuthor` from `@pptx-editor.dev/pro/react` and `@pptx-editor.dev/pro/vue`. The room format is now 7, so update every peer of a room together.
- 8253ddb: Link text to a slide or an email address. `setLink` takes `#slide-<n>`, `#first-slide`, `#last-slide`, `#next-slide`, and `#previous-slide`. `HyperLink.Fields` renders a wrapper, so its `className` goes on that wrapper.
- 3d5112e: Add the slide master view (`setPresentationView` with `slideMaster`): edit the shapes, placeholders, and background of layouts and masters, and every slide on them updates. Open it from **View**, the status bar, or **Edit layout** on the slide. `SlideSummary` gains `template`, `name`, `usedBy`, and `masterId`.
- 8253ddb: Change the slide size with `setSlideSize` or **Page setup**.
- 9d7d60b: Draw warped text in the editor and the PDF export. Set it with `setTextWarp` or `Toolbar.TextEffects`.

### Patch Changes

- 9d7d60b: Keep more of the file and converge in collaboration when you edit animations and transitions. `setTransition` refuses unknown kinds. Add `animationStartLabel`.
- a90b74e: Open the Bullets and Numbering dialog on the tab of the selected list. Add `SelectionFormat.bulletFormatMixed`.
- 9d7d60b: Check file links more strictly, keep safe file links through copy and paste, and refuse picture bullets over 25 million pixels.
- 5c0bcca: Clear stale chart drafts without moving focus, and preserve drafts when categories or series are appended. Hide placeholder insert buttons while another object is selected or edited. Show keyboard tooltips immediately and skip entry animation when moving between nearby controls in React and Vue.
- 8253ddb: Fix focus in the pattern and gradient dialogs. `setShapePicture` takes `mime`, and `slideBackground.tile` reports tiled backgrounds.
- f448820: `usePptxSource` stops stalled downloads with a `TimeoutError` and oversized downloads with a `RangeError`. Fonts hold the open for at most 2 seconds, and `registerFontFaces` takes `waitMs`.
- f448820: Open large slides without long main-thread tasks. Add `PaintOptions.view`.
- 9d3e18e: Keep the editor responsive on slides with thousands of shapes. On selections of more than 250 shapes, `can()` checks only the editing mode for some commands, and `exec()` makes the full check.
- cbee759: Show only the dragged slides in the slide drag image.
- 9c48ff4: End the slide show when the browser leaves full screen, so one Escape closes the show.
- Updated dependencies: @pptx-editor.dev/core@0.3.0, @pptx-editor.dev/i18n@0.3.0

## 0.2.0

### Minor Changes

- c292536: Without the drawing module of `@pptx-editor.dev/pro`, the chrome no longer offers glow, inner shadow, reflection, or soft edge. Shape Effects shows only the outer shadows, and Picture Styles shows only the frame styles. The effect commands still write every effect, and save keeps it. New `drawingCapabilities()` in `@pptx-editor.dev/core/layout` says what the installed renderers paint, and `chromeOfferedControl(slot, control)` in `@pptx-editor.dev/core/editor` gives the choices a slot offers in this build. The controls and the `useEditorValueCommand` options of React and Vue refresh after an install; `drawingRenderersVersion()` and `subscribeDrawingRenderers(listener)` in `@pptx-editor.dev/core/layout` give the signal.
- aa99dad: Lists and text are faster to edit from the keyboard.

  - Enter starts a new list item at the same level, also on an empty item.
  - Tab and Shift+Tab change the level only at the start of a list item or across paragraphs. Elsewhere, both keys type a tab.
  - Backspace at the start of a list item removes its bullet or number and keeps its level. The text moves to where the bullet was, and the next Backspace joins the item to the paragraph above. The new `removeBullet` command does this step.
  - Ctrl/Cmd+L, E, R, and J align paragraphs. Ctrl/Cmd+Shift+> and < step the font size. Ctrl/Cmd+] and [ step the list level. Shift+F3 steps the case.
  - Alt/Ctrl+Up and Down move to a paragraph start. A word step forward stops at the start of the next word. Cmd+Backspace deletes the word before the caret.
  - Enter on a selected shape selects all its text.
  - AutoCorrect works while you type: curly quotes, symbols, fractions, en dashes, sentence capitals, and lists started by typing `1.`, `1)`, `*`, or `-` and a space. Backspace right after a new list keeps the typed text. The new `autoCorrect` option and prop turn AutoCorrect off.

### Patch Changes

- 0057839: The editor chrome is easier to use with assistive technology and high contrast themes. `PptxEditor.Viewport` is a landmark region named "Presentation editor"; its new `label` prop gives another name, and `label={false}` removes the region. If your own `header` or `main` is inside the viewport, set `label={false}`, so your landmarks stay at the top level. Slides in the rail and the slide sorter take their name ("Slide 3", "Slide 3 (hidden)") from hidden text, and the thumbnail art is hidden from screen readers. The items have no `aria-label` now, so a test that finds a slide by its label (`getByLabelText('Slide 3')`, `[aria-label="Slide 3"]`) must find it by role and name (`getByRole('option', { name: 'Slide 3' })`). With a forced-colors theme, selected slides, pressed and checked toggles, focus rings, disabled controls, and card edges stay visible in system colors. Color swatches are 24 by 24 pixels, and 20 by 20 pixels in a dialog that is too narrow for them.
- bfca3d6: The Find and Replace buttons are off in the slide sorter. They turn on again in the Normal view.
- 159eb87: Without the drawing module of `@pptx-editor.dev/pro`, Shape Effects can now remove a reflection, glow, or soft edge from the file. The menu offers No reflection, No glow, or No soft edges when the selected shapes or pictures have that effect, and No shadow also removes an inner shadow. The menu follows the selection. `chromeOfferedControl(slot, control, editor)` takes the editor to give these rows to a custom control.
- 0653436: A keystroke in a deck of 5,000 slides now costs about the same as in a short deck. An edit records only the parts it changes and copies no list of every part. The slide summaries, the font list, and the font notice update from the changed parts only. The slide rail renders only the items whose slide changed, a thumbnail out of view takes no layout or paint, and the canvas reads its fit zoom only when the slide or canvas size changes.
- 5a03952: Published packages no longer ship source maps, so a stack trace points into `dist/`. Every JavaScript file of `pro`, `pptx-to-pdf`, and `editor-api` starts with its license header, and these packages are minified. The preset shape geometry in `core` carries the Ecma copyright notice in a legal comment, so bundlers that keep legal comments carry the notice into your app bundle.
- 6bed74e: Let a tall menu scroll on a short screen also after it is placed a second time. Before, its last rows could be out of reach.
- 60ed234: A long deck with a damaged slide or notes slide is refused at once, wherever the slide is read. After edits, the deck stays open read-only with `snapshot().readOnlyReason`, a status bar message says why, its other slides go on opening, and commands, undo, save, and automation writes and exports are refused. `snapshot().isOpening` is true until the first slide shows, `load` accepts any `startSlide` number, and the slide rail mounts in steps of about 8 ms.
- 60ed234: A long deck shows its first slide and takes input before the rest is ready: the other slides open in the background, and `snapshot().pendingSlides` counts them. `load(bytes, { startSlide })` opens another slide first. Edits, undo, Find, and save work before the deck is ready. Decks with up to about 10,000 slides now open, and the slide rail re-renders only the slides that change.
- 84256e4: Long decks open faster: the slide rail and the Slide Sorter paint the thumbnails in view at once, and the others when they scroll into view or when the browser is idle. `renderThumbnail` still paints at once by default; pass `{ defer: true }` to paint as the rail does, and call the function it returns when the element goes. Shape drags stay smooth because the deck changes once, on release, so collaborators see a drag only when it ends. The preview shows the size and text scale that autofit gives on release, and autofit text in a scaled group now fits at the size the group draws it. Escape, a lost window focus, an undo, or a change to a dragged shape or its group cancels the drag; a change to another shape keeps it. An `Editor` implementation must now return a disposer from `renderThumbnail`.
- 6bed74e: Keep keyboard focus on the current slide when you leave the slide sorter. Before, a layout that swaps the slide rail and the slide sorter left focus on the page body after a switch back to the Normal view.
- db11b83: The packaged editor keeps the view switch, the zoom, and the font notice above the speaker notes. They no longer cover the notes text. The canvas leaves room under the slide for them, so they also do not cover the slide.
- 1358a9d: Tables now show editing chrome on the canvas. A handle at the top-left corner selects and moves the table, grid lines drag, and "+" buttons insert rows and columns. Cell ranges copy and paste as tab-separated text, Shift+Arrow extends into a cell range, and the cell menu groups Insert, Delete, and Select in submenus. A table selected as an object has its own context menu kind, `tableFrame`, with the Insert submenu. The Format panel sets row height, column width, and table style options, and the Borders menu draws with a pen color, style, and weight. `insertColumn` now keeps the table width, and the columns narrow in proportion.
- cd341f2: Each package now accepts only the same minor version of `@pptx-editor.dev/core` as a peer, for example `~0.1.0`. The links in the package READMEs now open from npm.
- Updated dependencies [4594311]
- Updated dependencies [0057839]
- Updated dependencies [ba27037]
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
  - @pptx-editor.dev/i18n@0.2.0

## 0.1.0

### Minor Changes

- 412d750: The React and Vue adapters add `useFontFamily()` for a host font picker, export the core types their props and hooks use (such as `PptxEditorInstance`, `EditorCommand`, and `ChromeSlotId`), warn in development when `modules` changes after mount, and no longer export the deprecated `WORD_DEFAULT_FONT`.
- 412d750: `PptxEditor` and its root take a `modules` prop for the pro review and collaboration modules, a new `fontFamilies` list updates the font picker in place (`setFontFamilies`), and `save()` returns `Uint8Array<ArrayBuffer>`, so `new Blob([bytes])` type-checks on TypeScript 5.7 and later.
- 412d750: A title shrinks to 90 % at most on overflow, the font size box shows the size the text is drawn at, and a new Text fitting toolbar control (`paragraph.autofit`) switches between no autofit, shrink text on overflow, and resize shape to fit text.
- 412d750: Add the Bullets and Numbering galleries (`setBulletFormat`, the `list.bulletStyle` and `list.numberStyle` dropdowns), `PptxEditor.BulletsAndNumberingDialog` for size, color, and start number (with the dialog named parts, `useBulletsAndNumberingDialog`, and Format > Bullets and numbering…), and `useBulletFormat` in React and Vue.
- 929ca86: Build your own chrome with DOCX Editor's hooks. `useZoom` reads and sets the slide zoom, steps it, and fits the slide. `useEditorEvent` subscribes to `change`, `selectionChange`, `rejection`, and `openFailure`, and `onEditorEvent` gives the same events without an adapter. `useEditorSnapshot` follows every editor change, and `useScopeClassName` returns the classes that make the `--doc-*` tokens apply to your own parts. Vue has the same composables.

  More of DOCX Editor's hooks have PPTX twins, in React and Vue: `useChromeTranslate`, `useToolbarContext`, `useToolbarLabel`, and `useToolbarLabelFor` for labels, `useHistoryGroup` to make one gesture on your own control one undo step, `useEditorCaret` for the caret while text is edited, and `useDocumentSearch` for your own find bar. `Toolbar`, `Menu`, `ContextMenu`, `SlideRail`, and `StatusBar` take DOCX Editor's `t` prop. `StatusBar.Action` adds your own button to the status bar. `SlideRail.Expand` changes the Show thumbnails button of the folded rail, `toggle={false}` now also removes a `SlideRail.Collapse` child under `preset={false}`, a Vue `v-if` child no longer counts as a rail action, and the rail keeps the current slide in view when its panel changes size.

  `<PptxEditor>` takes DOCX Editor's part props: `toolbar`, `menu`, `contextMenu`, `hyperlinkPopup`, `slideRail`, and `notesPane` each take `true`, `false`, or the part's props, and `contextMenu` and `hyperlinkPopup` set the same `popups` entries. `useHyperlinkPopup` and `useHyperlinkPopupInstance` give the link popup's state with no UI, and with `popups.hyperlink: false` a mounted host panel opens for the link controls. Named menu rows (`ContextMenu.Cut`, `ContextMenu.MergeCells`, `Menu.PageSetup`, `Menu.Cut`, and more) change or hide one row with `label`, `icon`, `className`, and `hidden`. `FindBar`, `OpenFailure`, `FontNotice`, and `EditNotice` take `t`, and `useOpenFailure` takes `{ t }`. `usePptxSource` fetches a deck and resolves its fonts as DOCX Editor's `useDocxSource`, and `useEditorHostDocument` returns the document that the slide is mounted in. Vue has the same props, parts, and composables.

  The `--doc-slide-padding` token sets the room around the slide in the canvas (default `24px`), and zoom to fit keeps it, so a host can compute the fitted slide height for a width. The slide menu of a horizontal rail says Move left and Move right, with their shortcuts. A link that the open link popup refuses shows its reason once, in the popup, and not again in the edit notice. `useHyperlinkPopup().open(link)` checks `link.url` with `sanitizeHref` again and ignores the given `href`. `useDocumentSearch` reads the matches again when a hook mounts after edits made while none was mounted. Vue `usePptxSource` keeps the deck when a getter returns a new `fonts` value, and keeps the old bytes while a new URL loads, as React does. Vue `useHistoryGroup` binds an element from an iframe.

  `PptxEditor.ContextMenu` placed in the root or the viewport is the canvas menu, as in DOCX Editor, through the same entry as `popups.contextMenu`, which wins over it. It takes `disabled` and `onOpenChange` (Vue: `openChange`). The link popup opens at the selected text and has Cancel. A horizontal rail narrower than 420px keeps New slide and folds the other slide actions into one More menu. A right-to-left horizontal rail names its moves, shortcuts, and icons by its direction, as its arrow keys step. A change of a compound's `t` between set and unset no longer remounts it.

- 929ca86: Match the editor chrome to your app with public API only. `--doc-font-ui` now sets the font of the whole chrome, `--doc-radius-surface`, `--doc-radius-control`, and `--doc-radius-pill` (the zoom pill and chips) set its corners, the dialog panel falls back to `--doc-radius-surface`, dialog controls read `--doc-dialog-control-radius`, then `--doc-radius-control`, and the primary dialog button reads `--doc-primary-hover`. The slide rail's action row has named parts (`SlideRail.AddSlide`, `ChooseLayout`, `Duplicate`, `MoveEarlier`, `MoveLater`, `Delete`, `Collapse`, and `Action`) with `hidden`, `icon`, `className`, `asChild`, and `preset={false}`. A horizontal rail sets its action columns to half the buttons that show, rounded up, so a rail with `toggle={false}` shows its six buttons as three over three. The notes field fills the height that a host gives the notes pane. A focused rail slide rings its thumbnail, and the custom color field is as tall as Apply beside it. A context menu that is too tall for the room above or below the pointer now stays at the pointer and scrolls, so a slide menu opened in a rail at the bottom of a short window no longer covers the toolbar.
- 412d750: Collaboration sessions can own undo and redo, report a status snapshot with typed failure codes, and exchange presence and selections. The canvas shows collaboration peers' selected shapes and text carets. Custom `EditorCollaborationSession` implementations must now provide `statusSnapshot()`, and `subscribeStatus` listeners receive the reason code and detail.
- 412d750: `useEditorCommand(slotId).execute()` returns `true` when the engine ran the command and `false` on a refusal, as in DOCX Editor.
- 412d750: Right-click a slide in the rail or the slide canvas to open a context menu with icons, shortcut hints, and the reason when a command is unavailable.
- 412d750: Add a font substitution notice with local font access, and upright East Asian characters in vertical text.
- 412d750: After you delete slides or shapes, the edit notice says what went and offers Undo for ten seconds; `snapshot.lastRemoval` reports the removal.
- 412d750: Dialogs expose DOCX Editor's named parts (`Header`, `Title`, `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`, `Reset`) over a draft, with `usePageSetupDialog` and `useBackgroundDialog`: Format Background (the menu row is now "Format Background…") applies its choice on OK as one undo step and Cancel discards it, Page setup reports the appearance through `onThemeChange` on OK instead of on each click and has OK and Cancel instead of Close, and the Vue adapter gains the `popups` registry and the Page setup dialog.
- 412d750: Toolbar slot ids now use the DOCX Editor names (`font.family`, `font.size`, `format.clear`, `list.bullet`, `list.numbered`, `list.indent`, `list.outdent`, `list.lineSpacing`, `list.bulletStyle`, `list.numberStyle`), and `Toolbar.FontFamily` has `Trigger`, `Content`, and `Item` parts in React and Vue.
- 412d750: Fix lost or misplaced edits: speaker notes stay on their slide, undo and redo return to the edited slide, deleted slides leave no broken links, emoji and fields edit as one character, and refused edits show a reason.
- 412d750: `createPptxEditor` accepts `modules` for review and collaboration, and the store can publish parts that a collaboration replica rebuilt.
- 412d750: `PptxEditor` and its root now take a ref to a `PptxEditorRef` with `load`, `save`, `exec`, `snapshot`, `focus`, and `getEditor`, the same members in React and Vue.
- 412d750: Draw scaled, skewed, aligned, and rotating shadows, reflections, and effects on text in shapes without fill, without clipping large blurs.
- 412d750: Draw rectangular path gradients, preset patterns, gradient text, tiled pictures at their natural size, and compound lines.
- 412d750: Find and replace in slide text, tables, and speaker notes (`editor.find`, the `replace` and `replaceAll` commands, and the `FindBar` with Ctrl/Cmd+F and Ctrl/Cmd+H), and export of slides as PNG, a ZIP of PNGs, or standalone SVG (`exportSlidePng`, `exportSlidesPngZip`, `exportSlideSvg`).
- 412d750: Add DOCX Editor's font configuration API: the `fonts` option and prop (a `FontConfiguration`, a fragment, or a `FontResolver` such as `packagedFonts()`), `editor.setFonts()`, `onFontError`, `useFonts`, `loadFonts`, `createFontSource`, `customFonts`, and `deferFonts`, and replace `fontProvider` and `addFonts` with them.
- 412d750: Add `focus()` on the editor, superscript and subscript (`script.super`, `script.sub`, Ctrl/Cmd+Shift+= and Ctrl/Cmd+=), the Character Spacing and Change Case menus, the Font dialog (`setFontFormat`, `PptxEditor.FontDialog` with the dialog named parts and `useFontDialog`), and `useFontFormat` in React and Vue; formatting at a caret now applies to the word around it, or to the text typed next at a word edge.
- 412d750: Add the format painter: the `format.painter` toolbar button, Ctrl/Cmd+Shift+C and Ctrl/Cmd+Shift+V, and the `toggleFormatPainter`, `copyFormat`, and `pasteFormat` commands copy text or shape formatting and paint it as one undo step.
- 412d750: An empty list item shows its bullet or number dimmed while the caret is in it, and the caret starts where the item's text will start.
- 412d750: Edit inside groups, select and transform several shapes together (Shift+click, marquee, Ctrl/Cmd+A), group and ungroup, and copy, cut, and paste shapes within and between decks.
- 412d750: The core package now bundles its text shaper, so Bun and esbuild apps build, and `setHarfBuzzWasmUrl` points the shaper at a self-hosted `@pptx-editor.dev/core/harfbuzz.wasm`.
- 412d750: Group undo steps with `Editor.beginHistoryGroup()` and `exec(command, { historyGroup })`, restore the selection after a step on redo, and commit `projection` or `awareness` transactions that add no undo step; `runToolbarCommand` and `useEditorValueCommand().execute` take `{ historyGroup }` last; `historyGroup` replaces the store's `mergeKey` and `setTableGrid.gesture`, `closeHistoryGroup()` replaces `sealHistory()`, and `redo()` returns the selection after the step.
- 412d750: The slide rail takes `orientation="horizontal"` to run as a row under the slide in a tall, narrow host, and it keeps the current slide in view by scrolling only the rail, so a page around the editor no longer moves.
- 412d750: A click into linked text opens a link popup with Copy, Edit, and Remove (a read-only editor shows only the target and Copy), and `PptxEditor.HyperLink` and `useHyperlinkPopup` let you restyle, replace, or rebuild its parts; a `popups.hyperlink` renderer shows this reading mode only when registered with `definePopup(Component, { reading: true })`.
- 412d750: Add a loading screen (`PptxEditor.Loading` in React, `PptxEditorLoading` in Vue) that shows while a deck opens, and paint it before the engine parses a large deck.
- 412d750: `PptxEditor.Menu` and `PptxEditor.ContextMenu` render the menu bar and the canvas menus from a core registry (`CHROME_MENUS`, `CANVAS_CONTEXT_MENUS`) with `slotId` rows, `Group`, `Separator`, and `Submenu`, the `popups` registry gains `hyperlink`, `find`, `contextMenu`, and `railContextMenu`, `PptxEditor.Root` now hosts the link popup, and `useEditorCommand` runs dialog, clipboard, find, link, and slide show slots (in Vue, the clipboard and pictures; the others show as disabled with the reason).
- 412d750: Both adapters export the same props and return types, such as `PptxEditorViewportProps`, `FindBarProps`, `LoadingSpinnerProps`, `UseEditorCommandReturn`, and `UseSlideshowReturn` (the React `*Handle` names stay as deprecated aliases), and `PptxEditor` takes `locale`, `fontFamilies`, and `onChange` as `PptxEditor.Root` does.
- 412d750: Speaker notes show their formatting and keep it through every edit, and the notes pane sets bold, italic, and underline (Ctrl/Cmd+B, I, U, its buttons, or the toolbar while the notes have the cursor) through the new `formatNotes` command.
- 412d750: `PptxEditor` now shows a notice that says why a file did not open, with `PptxEditor.OpenFailure`, `useOpenFailure`, the `openFailure` option, and `renderOpenFailureActions` (the Vue `openFailureActions` slot) to change it.
- 412d750: The slide rail and the notes pane take DOCX Editor's pane props (`open`, `defaultOpen`, `onOpenChange`, `className`, and `paneWidth` and `toggle` on the rail), the side panel takes `className` and `paneWidth`, `useNotesPane` and `useEditorColorMode` give the notes state and the resolved theme, and `PptxEditor.Root` takes `colorMode` (`'light'`, `'dark'`, or `'system'`) for the chrome, with `colorMode` and `useEditorColorMode` in Vue too.
- 412d750: Add the Paragraph dialog: `setParagraphFormat` sets alignment, indents, space before and after, and line spacing as one undo step, `PptxEditor.ParagraphDialog` (with the dialog named parts and `useParagraphDialog`) opens from Format > Paragraph…, Line spacing options…, and the text context menu, `useParagraphFormat` reads and writes the same fields in React and Vue, with its `ParagraphFormatRead` and `ParagraphFormatCommand` types, and `setLineSpacing` now refuses a multiple below 0.01 or above 9.99, as the dialog does.
- 412d750: Pictures can be cropped, changed, reset, adjusted, styled, sized, and given alt text through new commands, the toolbar's `image` group, and the `imageAltText` and `imageProperties` popups, and picture borders now draw outside the picture.
- 412d750: Add PPTX Editor v0.1: standards-based slide rendering (ECMA-376), a minimal editing slice, and the React adapter.
- 412d750: The Vue adapter, the pro modules for review and collaboration, and PDF export are now published with the other packages.
- 412d750: Add links, clear formatting, slide backgrounds, line end point editing, a searchable font picker with host-loaded families, zoom that scrolls to every edge, and slide rail shortcuts.
- 412d750: Add `PptxEditor.Ribbon` and `PptxEditor.StatusBar` (Vue: `PptxEditorRibbon` and `PptxEditorStatusBar`), so your editor can show its commands on tabs, with contextual Format and Table tabs, and slide x of y, notes, views, and zoom under the slide.
- 412d750: `moveSlides` and `moveSlide` take an optional `section`, so a drop below a section header or a Ctrl/Cmd+Arrow step across one puts the slides into that section, and the saved file keeps one slide list with a valid id for every slide after two collaborators add the first slide at the same time. Behavior change: `moveSlide` now follows the rule of `moveSlides`, so a slide moved up to the first place of a section joins the section before it instead of that section; pass the section's id as `section` to keep the earlier result.
- 412d750: Sections show in the rail and the slide sorter and can be added, renamed, removed, moved, and collapsed (`addSection`, `renameSection`, `removeSection`, `removeAllSections`, `moveSection`, `setSectionsCollapsed`, the `section.*` slots), and the new Slide Sorter view (`setPresentationView`, `PptxEditor.SlideSorter`, the `view.presentationView` slot) shows every slide by section. The `railContextMenu` popup entry covers the slide sorter, and the new `sectionContextMenu` entry (`SectionContextMenu`) replaces or turns off the section menu. A slide entry that repeats a slide id of another slide keeps its slide with a fresh id, so the slide list that the editor shows, the commands, and the saved file agree. In the Vue adapter, the `section.*` slots draw as buttons and `view.presentationView` works through `useEditorCommand`. The default editor's zoom pill moved into the status bar, next to the view switch.
- 412d750: The toolbar shows only the controls for the selection (set `contextual={false}` to show every group), and superscript, subscript, character spacing, and columns leave the default bar and `formattingBarChromeGroups()`.
- 412d750: Add shadow, reflection, glow, and soft edge presets, Quick Styles, and commands to distribute, flip, rotate, lock, hide, and rename objects, with toolbar parts that a host places and a typed `useEditorValueCommand` hook.
- 412d750: Add a visual shape gallery, gradient fills with a custom gradient dialog, a line tool with editable line ends, vertical text alignment, single-click text editing, and a toolbar overflow menu.
- 412d750: Add shape fill, border color, border weight, border dash, and text fitting commands and controls, an editable Format panel, and a Page setup dialog that replaces the Settings tab.
- 412d750: Add change layout, reset slide, hide slide, and New slide with a layout gallery, through the new `applyLayout`, `resetSlides`, and `setSlidesHidden` commands, the `slide.*` slots, the rail's context menu, `LayoutGallery`, and the `useSlideLayouts` hook in React and Vue.
- 412d750: The menu bar has **Insert > New slide** (Ctrl+M, ⌃M on a Mac, now also while you edit text or notes) and **Slide Show > Hide slide**, and `Menu.Item` takes `direct` to run a value slot at once or open its choices.
- 412d750: Select several slides in the rail with Shift+click (a range) and Ctrl/Cmd+click (one more or one less), then delete, duplicate, or move them together as one undo step.
- 412d750: The `setSlideNumbers` command shows or hides slide numbers on every slide through the layout's slide-number placeholder (`p:ph type="sldNum"`).
- 412d750: Present a deck full screen with `PptxEditor.Slideshow` and `useSlideshow()`, or F5 and Ctrl/Cmd+Enter, with the standard slide show navigation keys, hidden slides skipped, and the editor's own slide painting.
- 412d750: Shapes snap to smart guides while you move or resize them (edges, centers, the slide, and equal spacing), Alt turns snapping off, and View > Smart guides switches it.
- 412d750: Tables are editable: type in cells, insert tables, rows, and columns, merge and split cells, and format cells and table styles with new commands and the toolbar's `table` group.
- 412d750: Render tables with their styles, merged cells, borders, and rows that fit their text, and draw text in styled shapes in the style's font color.
- 412d750: Add font and underline color to the Font dialog (`setFontFormat` `underlineColor`, drawn from `a:uFill`), `PptxEditor.ColumnsDialog` (Format > Columns… and "More columns…") for the column count and spacing, and a custom bullet character and font in the Bullets and Numbering dialog.
- 412d750: Add Text Direction and Columns (`setTextDirection` with Office.js orientation words (`TextOrientation`), `setColumns` with `count` and `gapEmu`, and the `paragraph.textDirection` and `paragraph.columns` menus in the React toolbar), and resize-to-fit and shrink-on-overflow that follow rotated text, columns, and the text rectangle of shapes that are not rectangles; `mongolianVertical` and `wordArtVerticalRtl` are refused and listed as approximations.
- 412d750: Apply the Font, Paragraph, and Bullets and Numbering formats, superscript, subscript, Change Case, and character spacing to table cells, type the text slots in `ToolbarValueMap` (`font.size` steps with `{ step: 1 }` or `{ step: -1 }`), and export `TextFormatCommand`, `BulletGalleryKey`, and `NumberGalleryKey`.
- 412d750: Color menus show the deck's theme colors with their lighter and darker variants and write them as theme colors, and every toolbar menu works with the keyboard.
- 412d750: `PptxEditor.Toolbar` takes `preset`, `Toolbar.Slot`, and `Toolbar.Group` for customization, and `PptxEditor.Root` takes a `popups` registry with `definePopup` and `useEditorDialog`.
- 412d750: The toolbar gains DOCX Editor's named parts (`Toolbar.Bold`, `Toolbar.FontFamily`, …), `Toolbar.Action` and `Toolbar.Separator`, group `priority`, `pinned`, `after`, `className`, and `overflowContent`, and `overflow={false}` to wrap the bar instead of using a More menu; the Vue toolbar gets the named parts, `ToolbarAction`, `ToolbarSeparator`, `after`, and `className`, and its bar wraps until it has a More menu.
- 412d750: The React adapter exports its stylesheet as `@pptx-editor.dev/react/styles.css`, and the logic it shares with the Vue adapter now lives in core.

### Patch Changes

- 412d750: Install the React or Vue adapter with `@pptx-editor.dev/core` only: the adapter now installs `@pptx-editor.dev/i18n` for you.
- 412d750: Tab and Shift+Tab move the selection through the shapes on the slide, a table cell gets its own context menu (`CanvasMenuKind` `table`) with row and column commands, and the canvas keeps the caret in view, places IME input at the caret, shows shapes outside the slide dimmed, and zooms in 10 % wheel steps within one zoom range.
- 412d750: The canvas context menu's **New slide** row is the `slide.add` registry row, so `ContextMenu.Slot slotId="slide.add"` hides or replaces it.
- 412d750: Controls show a styled tooltip with the reason a control is off, the font picker groups the theme, deck, and other fonts and applies a theme font as a theme reference (`SelectionFormat.fontFamilyTheme`), color menus take a custom hex color, menus close on any press outside, a menu opens on its checked row, a selected section shows its first slide, Change case and Insert link act on the word at the caret, and undo, find, notes formatting, and keyboard focus work as expected.
- 412d750: Typing and dragging stay fast on large decks: an edit lays out and repaints only the slides it changes, thumbnails repaint when the editor is idle, and two editors on one page no longer share SVG ids or link shortcuts.
- 3bf89a8: Menus, popovers, and tooltips now open at their controls in a transformed, scaled, or clipping host container. A host CSS reset no longer changes the chrome, and Ctrl/Cmd+F and the link popup now stay with their own editor.
- 412d750: When a deck embeds no regular face of a font, regular text now uses an upright face instead of the italic one.
- 412d750: Find counts and highlights each match in the speaker notes, Redo shows ⇧⌘Z on macOS, the custom color field shows the current color, and the slide sorter shows no zoom percentage.
- 412d750: The font notice drops the result of an earlier local font request when another presentation opens, and dialog buttons get a destructive style and press feedback.
- 412d750: The core package now carries the full Ecma copyright notice for its preset geometry data, and every toolbar icon is an official Material Symbol.
- 412d750: Enter starts a new paragraph, so bullets apply to single paragraphs; Tab and Shift+Tab change the list level and move the text; Backspace or Enter on an empty item leaves a list; and bullets in placeholders keep the master's style.
- 412d750: The embedded font decoder's LZCOMP stage is a new implementation, written from its specification, that decodes every font exactly as before.
- 412d750: Every menu row kind, including the Insert > Shape gallery row, now shows the `icon` that a host gives `Menu.Item` or `ContextMenu.Slot`.
- 412d750: Escape in a menu bar menu now returns focus to its menu bar item, the open find bar no longer covers the slide, and a disabled control's reason is read once.
- 412d750: Describe every feature by its ECMA-376 behavior in docs and labels: the export menu item reads "Presentation (.pptx)", the diagram placeholder labels read "Diagram", and the diagram placeholder prompt now uses the i18n key `canvas.prompts.diagram`.
- 412d750: Bold, italic, and underline for the speaker notes say why they are off and follow every selection change, Tab and Escape leave the notes after typing, and a locale change on the root repaints the canvas placeholder prompts once, also with an inline `locale` object.
- 412d750: The package reader and writer now come from a shared internal OOXML module, with no change in behavior or public API.
- 412d750: Each package now ships its Apache-2.0 `LICENSE` file and a README, and the `core` third-party notices name the Material Symbols and ECMA-376 geometry data it contains.
- 412d750: The picture Corrections control shows its brightness and contrast presets as a grid of preview samples instead of a long list.
- 412d750: Ctrl/Cmd+Z and redo now work while the slide rail has focus.
- 412d750: Open decks whose custom geometry asks for huge arcs without hanging, and draw text in flipped groups, line ends, trailing line breaks, aligned bullets, numbering restarts, arcs in stretched shapes, and faded text decorations correctly.
- 412d750: Moving slides refuses an invalid target, the rail selection stays valid when a host deletes the current slide, and Ctrl+click on a Mac opens the context menu without collapsing the text selection.
- 412d750: Text typed over a selection is one undo step, a quick drag in the slide rail or sorter reorders slides, the Format menu names its picture row Format picture, and the find bar moves the slide only when it would cover it.
- 412d750: Keep sections and custom shows in step when you add, duplicate, move, or delete slides, give a duplicated slide its own charts, comments, and embedded objects, and remove the parts only a deleted slide used.
- 412d750: A slide show that starts right after another one ends no longer closes at once, and leaving full screen keeps the show running in the window.
- 412d750: Refuse zip bombs, oversized embedded font tables, deep XML, and unsafe names in edits, and open and save decks with SVG media, encoded part names, or large blank bitmaps without data loss.
- 412d750: The packages build with tsup 9 and esbuild 0.28, with no change to the public API.
- 412d750: The Vue adapter now has every part of the React adapter under the same names, including the menu bar, context menus, ribbon, status bar, slide rail, slide sorter, notes pane, side panel, find bar, slide show, form dialogs, `Icon`, `MenuRow`, `MenuSubmenu`, and `useNotesPane`, and its parts take `className` and `style` as the React parts do.
- 412d750: A plain `class` on the Vue `PptxEditorToolbar.Slot` now goes on its control, as `className` on the React `Toolbar.Slot` does, and a Vue named part with both `class` and `control-class` now adds both names instead of only `control-class`. Both adapters print a development warning when a slot's own content replaces the control, so its class or icon has nothing to apply to, and when a slot has an attribute that never reaches its control, such as `style`, `id`, or `data-*`.
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [457a754]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
- Updated dependencies [412d750]
  - @pptx-editor.dev/i18n@0.1.0
