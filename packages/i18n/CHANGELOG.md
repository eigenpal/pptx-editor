# @pptx-editor.dev/i18n

## 0.5.0

### Patch Changes

- ef380c0: Set the transparency of a shape's solid fill from the Format pane, with a slider and a percent field under Fill color. `ShapeFormatDetail.solid` reads the fill color and opacity, and the `{ kind: 'solid', opacity }` patch writes it. A new fill color now keeps the current transparency.

## 0.4.0

### Patch Changes

- 91482da: Add customizable selection-aware inspectors for React and Vue, with safe drafts and matching automation APIs. Edit line widths, endpoints, gradients, shadows, and reflections while preserving unedited XML. Keep picture pickers bound to their original document and selection. Line endpoint commands now refuse mixed line and non-line selections without changing any object.

## 0.3.1

No changes in this release.

## 0.3.0

### Minor Changes

- 8253ddb: Drag the adjust handles of a preset shape, or set adjust values with `setAdjustments`.
- 9d7d60b: Add animation and transition editing with `addAnimation`, `removeAnimation`, `moveAnimation`, `setAnimationTiming`, `setTransition`, `EditorSnapshot.animations`, `showAnimationOrder`, and the **Animations** side panel tab. The editor does not play effects or transitions.
- 9d7d60b: Edit the data of bar, column, line, area, pie, doughnut, radar, and scatter charts with `editor.chartData()`, `setChartData`, and the **Chart Data** side panel tab.
- 9d7d60b: Add **New comment** to the canvas context menus, with `registerCommentDraft` and `requestCommentDraft`. Comment commands take `expect` and can fail with `errors.commentChanged`, `errors.commentAmbiguous`, or `errors.noComment`. Threads report `identity`; send it in `expect`. Ids of repeated comments, comments with `~`, and some legacy comments change, so read ids from the current snapshot.
- 4ae229f: Drag a segment handle of an elbow or curved connector to move the segment. On an end segment, the drag adds a bend.
- 9d7d60b: Type in diagram shapes, and add or delete diagram nodes with `addDiagramShape` and `deleteDiagramShape`. Add `EditorSnapshot.diagram`.
- 7046dbf: Edit diagram text in the **Diagram** side panel tab, opened with **Edit text…**. Add `setDiagramText`, and fix context menu rows that open a side panel tab.
- 9d7d60b: Draw equations on the canvas and in the PDF export. Insert and edit them with `insertEquation`, `setEquation`, and the Equation dialog.
- c47f5ce: Decks no longer stay on the loading spinner: the open waits at most 2 seconds for fonts. Add `editor.fontMeasurement()`, `openFailureReport()`, and the `showCopy` prop of `PptxEditorOpenFailure`. `DocumentOpenFailureReason` adds `blocked`, so update exhaustive maps of it. Large and damaged decks open faster and more safely.
- fef01c9: Insert a chart with **Insert** > **Chart** or `insertChart`. Change its type, title, legend, labels, axes, gridlines, and colors in the **Chart** tab or with `setChartType`, `setChartFormat`, and `setChartColor`, and read them with `editor.chartFormat()`. `requestPane` takes `{ focus: false }`.
- 8253ddb: Fill shapes with preset patterns or pictures with `setShapePattern` and `setShapePicture`, and set pattern or tiled picture slide backgrounds. Add `PatternDialog`.
- 9d7d60b: Draw picture bullets, and set them with `setBulletFormat`. `setLink` links to custom shows and to relative file paths. Add `EditorSnapshot.customShows`.
- 8253ddb: Add the selection pane: the **Selection** side panel tab and the `SelectionPane` component. `SidePanel` takes `tab` and `onTabChange` (`v-model:tab` in Vue). Add `requestPane` and `listenPaneRequests`. The Vue adapter requires Vue 3.5 or later.
- 9d7d60b: Add slide comments with the comment commands, `readCommentThreads`, the `pane` option of `reviewModule`, and `PptxEditorReview`, `useReview`, `useReviewOf`, `useReviewItem`, and `useReviewAuthor` from `@pptx-editor.dev/pro/react` and `@pptx-editor.dev/pro/vue`. The room format is now 7, so update every peer of a room together.
- 8253ddb: Link text to a slide or an email address. `setLink` takes `#slide-<n>`, `#first-slide`, `#last-slide`, `#next-slide`, and `#previous-slide`. `HyperLink.Fields` renders a wrapper, so its `className` goes on that wrapper.
- 3d5112e: Add the slide master view (`setPresentationView` with `slideMaster`): edit the shapes, placeholders, and background of layouts and masters, and every slide on them updates. Open it from **View**, the status bar, or **Edit layout** on the slide. `SlideSummary` gains `template`, `name`, `usedBy`, and `masterId`.
- 8253ddb: Change the slide size with `setSlideSize` or **Page setup**.
- 9d7d60b: Draw warped text in the editor and the PDF export. Set it with `setTextWarp` or `Toolbar.TextEffects`.
- 9d7d60b: Draw 3-D shapes with `drawingModule` from `@pptx-editor.dev/pro/drawing`, and preset shadows in every edition. Set bevels, 3-D rotation, and blur with `setShapeEffect` or the Shape Effects menu.

### Patch Changes

- 9d7d60b: Keep more of the file and converge in collaboration when you edit animations and transitions. `setTransition` refuses unknown kinds. Add `animationStartLabel`.
- 9d7d60b: Check file links more strictly, keep safe file links through copy and paste, and refuse picture bullets over 25 million pixels.
- 30dc48e: Add the `setFooter` and `setDateTime` commands, and `EditorSnapshot.footer` and `EditorSnapshot.dateTime`. `setSlideNumbers` takes `indices` and `hideOnTitleSlides`, and refuses `show: false` when no slide can have a slide number.
- 9d3e18e: Keep the editor responsive on slides with thousands of shapes. On selections of more than 250 shapes, `can()` checks only the editing mode for some commands, and `exec()` makes the full check.
- 39dc46e: Add the `app.sheetGrip` label for the grip of a side panel shown as a sheet on a narrow screen.

## 0.2.0

### Patch Changes

- 0057839: The editor chrome is easier to use with assistive technology and high contrast themes. `PptxEditor.Viewport` is a landmark region named "Presentation editor"; its new `label` prop gives another name, and `label={false}` removes the region. If your own `header` or `main` is inside the viewport, set `label={false}`, so your landmarks stay at the top level. Slides in the rail and the slide sorter take their name ("Slide 3", "Slide 3 (hidden)") from hidden text, and the thumbnail art is hidden from screen readers. The items have no `aria-label` now, so a test that finds a slide by its label (`getByLabelText('Slide 3')`, `[aria-label="Slide 3"]`) must find it by role and name (`getByRole('option', { name: 'Slide 3' })`). With a forced-colors theme, selected slides, pressed and checked toggles, focus rings, disabled controls, and card edges stay visible in system colors. Color swatches are 24 by 24 pixels, and 20 by 20 pixels in a dialog that is too narrow for them.
- ba27037: Add the strings for the demos' PDF export: `menu.exportPdf`, `menu.exportPdfHint`, `menu.exportPdfBusy`, `app.pdfExportFailed`, `app.pdfTooLarge`, `app.pdfTooComplex`, and `app.pdfBusy`.
- 5a03952: Published packages no longer ship source maps, so a stack trace points into `dist/`. Every JavaScript file of `pro`, `pptx-to-pdf`, and `editor-api` starts with its license header, and these packages are minified. The preset shape geometry in `core` carries the Ecma copyright notice in a legal comment, so bundlers that keep legal comments carry the notice into your app bundle.
- b906153: A screen reader now hears the alt text of a picture on the slide. A described picture is an image named by its description. A screen reader skips a decorative picture and a picture with no description. When Tab selects an object, the screen reader hears its kind, its name, and its description. The selection overlay is hidden from assistive technology. In the slide show, a screen reader reads the slide text and the alt text of its pictures. `PaintOptions` takes an optional `altText` port. Without it, the paint has no ARIA.
- 60ed234: A long deck with a damaged slide or notes slide is refused at once, wherever the slide is read. After edits, the deck stays open read-only with `snapshot().readOnlyReason`, a status bar message says why, its other slides go on opening, and commands, undo, save, and automation writes and exports are refused. `snapshot().isOpening` is true until the first slide shows, `load` accepts any `startSlide` number, and the slide rail mounts in steps of about 8 ms.

## 0.1.0

### Minor Changes

- 412d750: Add the Bullets and Numbering galleries (`setBulletFormat`, the `list.bulletStyle` and `list.numberStyle` dropdowns), `PptxEditor.BulletsAndNumberingDialog` for size, color, and start number (with the dialog named parts, `useBulletsAndNumberingDialog`, and Format > Bullets and numbering…), and `useBulletFormat` in React and Vue.
- 412d750: Collaboration sessions can own undo and redo, report a status snapshot with typed failure codes, and exchange presence and selections. The canvas shows collaboration peers' selected shapes and text carets. Custom `EditorCollaborationSession` implementations must now provide `statusSnapshot()`, and `subscribeStatus` listeners receive the reason code and detail.
- 412d750: Right-click a slide in the rail or the slide canvas to open a context menu with icons, shortcut hints, and the reason when a command is unavailable.
- 412d750: Add a font substitution notice with local font access, and upright East Asian characters in vertical text.
- 412d750: Dialogs expose DOCX Editor's named parts (`Header`, `Title`, `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`, `Reset`) over a draft, with `usePageSetupDialog` and `useBackgroundDialog`: Format Background (the menu row is now "Format Background…") applies its choice on OK as one undo step and Cancel discards it, Page setup reports the appearance through `onThemeChange` on OK instead of on each click and has OK and Cancel instead of Close, and the Vue adapter gains the `popups` registry and the Page setup dialog.
- 457a754: East Asian text measures and draws with each run's `a:ea` font (complex-script text with `a:cs`, symbol characters with `a:sym`), and a theme with an empty `a:ea` uses its script font for the run's language. `eaLnBrk="0"` now breaks East Asian text between any two characters with no kinsoku rule, `latinLnBrk="1"` breaks inside Latin words, Korean text breaks at spaces, `hangingPunct` lets a comma or full stop hang past the line end, the presentation's `p:kinsoku` lists replace the built-in lists, and justified East Asian lines spread between characters, as the reference PDFs show. A missing East Asian font falls back to a host CJK face for the text's language, such as Noto Sans CJK JP, and a character that the run's face lacks uses another loaded face. The canvas shows the text that an input method composes at the caret, and the key that ends a composition runs no shortcut. `@pptx-editor.dev/i18n` adds Japanese (`/ja`), Korean (`/ko`), and Simplified Chinese (`/zh-CN`) catalogs, and `@pptx-editor.dev/pptx-to-pdf` draws East Asian text with Noto Sans CJK JP when the optional `@docx-editor.dev/fonts-cjk` is installed.
- 412d750: Fix lost or misplaced edits: speaker notes stay on their slide, undo and redo return to the edited slide, deleted slides leave no broken links, emoji and fields edit as one character, and refused edits show a reason.
- 412d750: `createPptxEditor` accepts `modules` for review and collaboration, and the store can publish parts that a collaboration replica rebuilt.
- 412d750: Draw scaled, skewed, aligned, and rotating shadows, reflections, and effects on text in shapes without fill, without clipping large blurs.
- 412d750: Draw rectangular path gradients, preset patterns, gradient text, tiled pictures at their natural size, and compound lines.
- 412d750: Find and replace in slide text, tables, and speaker notes (`editor.find`, the `replace` and `replaceAll` commands, and the `FindBar` with Ctrl/Cmd+F and Ctrl/Cmd+H), and export of slides as PNG, a ZIP of PNGs, or standalone SVG (`exportSlidePng`, `exportSlidesPngZip`, `exportSlideSvg`).
- 412d750: Add `focus()` on the editor, superscript and subscript (`script.super`, `script.sub`, Ctrl/Cmd+Shift+= and Ctrl/Cmd+=), the Character Spacing and Change Case menus, the Font dialog (`setFontFormat`, `PptxEditor.FontDialog` with the dialog named parts and `useFontDialog`), and `useFontFormat` in React and Vue; formatting at a caret now applies to the word around it, or to the text typed next at a word edge.
- 412d750: Add the format painter: the `format.painter` toolbar button, Ctrl/Cmd+Shift+C and Ctrl/Cmd+Shift+V, and the `toggleFormatPainter`, `copyFormat`, and `pasteFormat` commands copy text or shape formatting and paint it as one undo step.
- 412d750: An empty list item shows its bullet or number dimmed while the caret is in it, and the caret starts where the item's text will start.
- 412d750: Edit inside groups, select and transform several shapes together (Shift+click, marquee, Ctrl/Cmd+A), group and ungroup, and copy, cut, and paste shapes within and between decks.
- 412d750: The core package now bundles its text shaper, so Bun and esbuild apps build, and `setHarfBuzzWasmUrl` points the shaper at a self-hosted `@pptx-editor.dev/core/harfbuzz.wasm`.
- 412d750: The slide rail takes `orientation="horizontal"` to run as a row under the slide in a tall, narrow host, and it keeps the current slide in view by scrolling only the rail, so a page around the editor no longer moves.
- 412d750: A click into linked text opens a link popup with Copy, Edit, and Remove (a read-only editor shows only the target and Copy), and `PptxEditor.HyperLink` and `useHyperlinkPopup` let you restyle, replace, or rebuild its parts; a `popups.hyperlink` renderer shows this reading mode only when registered with `definePopup(Component, { reading: true })`.
- 412d750: `PptxEditor.Menu` and `PptxEditor.ContextMenu` render the menu bar and the canvas menus from a core registry (`CHROME_MENUS`, `CANVAS_CONTEXT_MENUS`) with `slotId` rows, `Group`, `Separator`, and `Submenu`, the `popups` registry gains `hyperlink`, `find`, `contextMenu`, and `railContextMenu`, `PptxEditor.Root` now hosts the link popup, and `useEditorCommand` runs dialog, clipboard, find, link, and slide show slots (in Vue, the clipboard and pictures; the others show as disabled with the reason).
- 412d750: Speaker notes show their formatting and keep it through every edit, and the notes pane sets bold, italic, and underline (Ctrl/Cmd+B, I, U, its buttons, or the toolbar while the notes have the cursor) through the new `formatNotes` command.
- 412d750: The slide rail and the notes pane take DOCX Editor's pane props (`open`, `defaultOpen`, `onOpenChange`, `className`, and `paneWidth` and `toggle` on the rail), the side panel takes `className` and `paneWidth`, `useNotesPane` and `useEditorColorMode` give the notes state and the resolved theme, and `PptxEditor.Root` takes `colorMode` (`'light'`, `'dark'`, or `'system'`) for the chrome, with `colorMode` and `useEditorColorMode` in Vue too.
- 412d750: Add the Paragraph dialog: `setParagraphFormat` sets alignment, indents, space before and after, and line spacing as one undo step, `PptxEditor.ParagraphDialog` (with the dialog named parts and `useParagraphDialog`) opens from Format > Paragraph…, Line spacing options…, and the text context menu, `useParagraphFormat` reads and writes the same fields in React and Vue, with its `ParagraphFormatRead` and `ParagraphFormatCommand` types, and `setLineSpacing` now refuses a multiple below 0.01 or above 9.99, as the dialog does.
- 412d750: Pictures can be cropped, changed, reset, adjusted, styled, sized, and given alt text through new commands, the toolbar's `image` group, and the `imageAltText` and `imageProperties` popups, and picture borders now draw outside the picture.
- 412d750: Empty placeholders show their prompt text ("Click to add title", or the layout's custom prompt) in the editor in the style typed text takes; thumbnails, exports, and saved files never hold it.
- 412d750: Add PPTX Editor v0.1: standards-based slide rendering (ECMA-376), a minimal editing slice, and the React adapter.
- 412d750: The Vue adapter, the pro modules for review and collaboration, and PDF export are now published with the other packages.
- 412d750: Add links, clear formatting, slide backgrounds, line end point editing, a searchable font picker with host-loaded families, zoom that scrolls to every edge, and slide rail shortcuts.
- 412d750: Add `PptxEditor.Ribbon` and `PptxEditor.StatusBar` (Vue: `PptxEditorRibbon` and `PptxEditorStatusBar`), so your editor can show its commands on tabs, with contextual Format and Table tabs, and slide x of y, notes, views, and zoom under the slide.
- 412d750: Sections show in the rail and the slide sorter and can be added, renamed, removed, moved, and collapsed (`addSection`, `renameSection`, `removeSection`, `removeAllSections`, `moveSection`, `setSectionsCollapsed`, the `section.*` slots), and the new Slide Sorter view (`setPresentationView`, `PptxEditor.SlideSorter`, the `view.presentationView` slot) shows every slide by section. The `railContextMenu` popup entry covers the slide sorter, and the new `sectionContextMenu` entry (`SectionContextMenu`) replaces or turns off the section menu. A slide entry that repeats a slide id of another slide keeps its slide with a fresh id, so the slide list that the editor shows, the commands, and the saved file agree. In the Vue adapter, the `section.*` slots draw as buttons and `view.presentationView` works through `useEditorCommand`. The default editor's zoom pill moved into the status bar, next to the view switch.
- 412d750: Add shadow, reflection, glow, and soft edge presets, Quick Styles, and commands to distribute, flip, rotate, lock, hide, and rename objects, with toolbar parts that a host places and a typed `useEditorValueCommand` hook.
- 412d750: Add a visual shape gallery, gradient fills with a custom gradient dialog, a line tool with editable line ends, vertical text alignment, single-click text editing, and a toolbar overflow menu.
- 412d750: Add shape fill, border color, border weight, border dash, and text fitting commands and controls, an editable Format panel, and a Page setup dialog that replaces the Settings tab.
- 412d750: Add change layout, reset slide, hide slide, and New slide with a layout gallery, through the new `applyLayout`, `resetSlides`, and `setSlidesHidden` commands, the `slide.*` slots, the rail's context menu, `LayoutGallery`, and the `useSlideLayouts` hook in React and Vue.
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

### Patch Changes

- 412d750: Tab and Shift+Tab move the selection through the shapes on the slide, a table cell gets its own context menu (`CanvasMenuKind` `table`) with row and column commands, and the canvas keeps the caret in view, places IME input at the caret, shows shapes outside the slide dimmed, and zooms in 10 % wheel steps within one zoom range.
- 412d750: Controls show a styled tooltip with the reason a control is off, the font picker groups the theme, deck, and other fonts and applies a theme font as a theme reference (`SelectionFormat.fontFamilyTheme`), color menus take a custom hex color, menus close on any press outside, a menu opens on its checked row, a selected section shows its first slide, Change case and Insert link act on the word at the caret, and undo, find, notes formatting, and keyboard focus work as expected.
- 412d750: Toolbar slot ids now use the DOCX Editor names (`font.family`, `font.size`, `format.clear`, `list.bullet`, `list.numbered`, `list.indent`, `list.outdent`, `list.lineSpacing`, `list.bulletStyle`, `list.numberStyle`), and `Toolbar.FontFamily` has `Trigger`, `Content`, and `Item` parts in React and Vue.
- 412d750: Typing and dragging stay fast on large decks: an edit lays out and repaints only the slides it changes, thumbnails repaint when the editor is idle, and two editors on one page no longer share SVG ids or link shortcuts.
- 412d750: When a deck embeds no regular face of a font, regular text now uses an upright face instead of the italic one.
- 412d750: The core package now carries the full Ecma copyright notice for its preset geometry data, and every toolbar icon is an official Material Symbol.
- 412d750: Enter starts a new paragraph, so bullets apply to single paragraphs; Tab and Shift+Tab change the list level and move the text; Backspace or Enter on an empty item leaves a list; and bullets in placeholders keep the master's style.
- 412d750: The embedded font decoder's LZCOMP stage is a new implementation, written from its specification, that decodes every font exactly as before.
- 412d750: Both adapters export the same props and return types, such as `PptxEditorViewportProps`, `FindBarProps`, `LoadingSpinnerProps`, `UseEditorCommandReturn`, and `UseSlideshowReturn` (the React `*Handle` names stay as deprecated aliases), and `PptxEditor` takes `locale`, `fontFamilies`, and `onChange` as `PptxEditor.Root` does.
- 412d750: Describe every feature by its ECMA-376 behavior in docs and labels: the export menu item reads "Presentation (.pptx)", the diagram placeholder labels read "Diagram", and the diagram placeholder prompt now uses the i18n key `canvas.prompts.diagram`.
- 412d750: The package reader and writer now come from a shared internal OOXML module, with no change in behavior or public API.
- 412d750: `PptxEditor` now shows a notice that says why a file did not open, with `PptxEditor.OpenFailure`, `useOpenFailure`, the `openFailure` option, and `renderOpenFailureActions` (the Vue `openFailureActions` slot) to change it.
- 412d750: Each package now ships its Apache-2.0 `LICENSE` file and a README, and the `core` third-party notices name the Material Symbols and ECMA-376 geometry data it contains.
- 412d750: Ctrl/Cmd+Z and redo now work while the slide rail has focus.
- 412d750: Open decks whose custom geometry asks for huge arcs without hanging, and draw text in flipped groups, line ends, trailing line breaks, aligned bullets, numbering restarts, arcs in stretched shapes, and faded text decorations correctly.
- 412d750: Moving slides refuses an invalid target, the rail selection stays valid when a host deletes the current slide, and Ctrl+click on a Mac opens the context menu without collapsing the text selection.
- 412d750: Add the `menu.openSample` string, and list picture brightness and contrast (`a:lum`) as a known approximation.
- 412d750: The menu bar has **Insert > New slide** (Ctrl+M, ⌃M on a Mac, now also while you edit text or notes) and **Slide Show > Hide slide**, and `Menu.Item` takes `direct` to run a value slot at once or open its choices.
- 412d750: Text typed over a selection is one undo step, a quick drag in the slide rail or sorter reorders slides, the Format menu names its picture row Format picture, and the find bar moves the slide only when it would cover it.
- 412d750: Keep sections and custom shows in step when you add, duplicate, move, or delete slides, give a duplicated slide its own charts, comments, and embedded objects, and remove the parts only a deleted slide used.
- 412d750: Refuse zip bombs, oversized embedded font tables, deep XML, and unsafe names in edits, and open and save decks with SVG media, encoded part names, or large blank bitmaps without data loss.
- 412d750: The packages build with tsup 9 and esbuild 0.28, with no change to the public API.
- 412d750: The Vue adapter now has every part of the React adapter under the same names, including the menu bar, context menus, ribbon, status bar, slide rail, slide sorter, notes pane, side panel, find bar, slide show, form dialogs, `Icon`, `MenuRow`, `MenuSubmenu`, and `useNotesPane`, and its parts take `className` and `style` as the React parts do.
- 412d750: Add the `app.adapter` string, the label of the React and Vue switcher in the demos.
