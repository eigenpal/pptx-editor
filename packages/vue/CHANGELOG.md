# @pptx-editor.dev/vue

## 0.1.0

### Minor Changes

- 412d750: The React and Vue adapters add `useFontFamily()` for a host font picker, export the core types their props and hooks use (such as `PptxEditorInstance`, `EditorCommand`, and `ChromeSlotId`), warn in development when `modules` changes after mount, and no longer export the deprecated `WORD_DEFAULT_FONT`.
- 412d750: `PptxEditor` and its root take a `modules` prop for the pro review and collaboration modules, a new `fontFamilies` list updates the font picker in place (`setFontFamilies`), and `save()` returns `Uint8Array<ArrayBuffer>`, so `new Blob([bytes])` type-checks on TypeScript 5.7 and later.
- 412d750: Add the Bullets and Numbering galleries (`setBulletFormat`, the `list.bulletStyle` and `list.numberStyle` dropdowns), `PptxEditor.BulletsAndNumberingDialog` for size, color, and start number (with the dialog named parts, `useBulletsAndNumberingDialog`, and Format > Bullets and numbering…), and `useBulletFormat` in React and Vue.
- 929ca86: Build your own chrome with DOCX Editor's hooks. `useZoom` reads and sets the slide zoom, steps it, and fits the slide. `useEditorEvent` subscribes to `change`, `selectionChange`, `rejection`, and `openFailure`, and `onEditorEvent` gives the same events without an adapter. `useEditorSnapshot` follows every editor change, and `useScopeClassName` returns the classes that make the `--doc-*` tokens apply to your own parts. Vue has the same composables.

  More of DOCX Editor's hooks have PPTX twins, in React and Vue: `useChromeTranslate`, `useToolbarContext`, `useToolbarLabel`, and `useToolbarLabelFor` for labels, `useHistoryGroup` to make one gesture on your own control one undo step, `useEditorCaret` for the caret while text is edited, and `useDocumentSearch` for your own find bar. `Toolbar`, `Menu`, `ContextMenu`, `SlideRail`, and `StatusBar` take DOCX Editor's `t` prop. `StatusBar.Action` adds your own button to the status bar. `SlideRail.Expand` changes the Show thumbnails button of the folded rail, `toggle={false}` now also removes a `SlideRail.Collapse` child under `preset={false}`, a Vue `v-if` child no longer counts as a rail action, and the rail keeps the current slide in view when its panel changes size.

  `<PptxEditor>` takes DOCX Editor's part props: `toolbar`, `menu`, `contextMenu`, `hyperlinkPopup`, `slideRail`, and `notesPane` each take `true`, `false`, or the part's props, and `contextMenu` and `hyperlinkPopup` set the same `popups` entries. `useHyperlinkPopup` and `useHyperlinkPopupInstance` give the link popup's state with no UI, and with `popups.hyperlink: false` a mounted host panel opens for the link controls. Named menu rows (`ContextMenu.Cut`, `ContextMenu.MergeCells`, `Menu.PageSetup`, `Menu.Cut`, and more) change or hide one row with `label`, `icon`, `className`, and `hidden`. `FindBar`, `OpenFailure`, `FontNotice`, and `EditNotice` take `t`, and `useOpenFailure` takes `{ t }`. `usePptxSource` fetches a deck and resolves its fonts as DOCX Editor's `useDocxSource`, and `useEditorHostDocument` returns the document that the slide is mounted in. Vue has the same props, parts, and composables.

  The `--doc-slide-padding` token sets the room around the slide in the canvas (default `24px`), and zoom to fit keeps it, so a host can compute the fitted slide height for a width. The slide menu of a horizontal rail says Move left and Move right, with their shortcuts. A link that the open link popup refuses shows its reason once, in the popup, and not again in the edit notice. `useHyperlinkPopup().open(link)` checks `link.url` with `sanitizeHref` again and ignores the given `href`. `useDocumentSearch` reads the matches again when a hook mounts after edits made while none was mounted. Vue `usePptxSource` keeps the deck when a getter returns a new `fonts` value, and keeps the old bytes while a new URL loads, as React does. Vue `useHistoryGroup` binds an element from an iframe.

  `PptxEditor.ContextMenu` placed in the root or the viewport is the canvas menu, as in DOCX Editor, through the same entry as `popups.contextMenu`, which wins over it. It takes `disabled` and `onOpenChange` (Vue: `openChange`). The link popup opens at the selected text and has Cancel. A horizontal rail narrower than 420px keeps New slide and folds the other slide actions into one More menu. A right-to-left horizontal rail names its moves, shortcuts, and icons by its direction, as its arrow keys step. A change of a compound's `t` between set and unset no longer remounts it.

- 929ca86: Match the editor chrome to your app with public API only. `--doc-font-ui` now sets the font of the whole chrome, `--doc-radius-surface`, `--doc-radius-control`, and `--doc-radius-pill` (the zoom pill and chips) set its corners, the dialog panel falls back to `--doc-radius-surface`, dialog controls read `--doc-dialog-control-radius`, then `--doc-radius-control`, and the primary dialog button reads `--doc-primary-hover`. The slide rail's action row has named parts (`SlideRail.AddSlide`, `ChooseLayout`, `Duplicate`, `MoveEarlier`, `MoveLater`, `Delete`, `Collapse`, and `Action`) with `hidden`, `icon`, `className`, `asChild`, and `preset={false}`. A horizontal rail sets its action columns to half the buttons that show, rounded up, so a rail with `toggle={false}` shows its six buttons as three over three. The notes field fills the height that a host gives the notes pane. A focused rail slide rings its thumbnail, and the custom color field is as tall as Apply beside it. A context menu that is too tall for the room above or below the pointer now stays at the pointer and scrolls, so a slide menu opened in a rail at the bottom of a short window no longer covers the toolbar.
- 412d750: `useEditorCommand(slotId).execute()` returns `true` when the engine ran the command and `false` on a refusal, as in DOCX Editor.
- 412d750: After you delete slides or shapes, the edit notice says what went and offers Undo for ten seconds; `snapshot.lastRemoval` reports the removal.
- 412d750: Dialogs expose DOCX Editor's named parts (`Header`, `Title`, `Body`, `Footer`, `Field`, `Error`, `Apply`, `Cancel`, `Reset`) over a draft, with `usePageSetupDialog` and `useBackgroundDialog`: Format Background (the menu row is now "Format Background…") applies its choice on OK as one undo step and Cancel discards it, Page setup reports the appearance through `onThemeChange` on OK instead of on each click and has OK and Cancel instead of Close, and the Vue adapter gains the `popups` registry and the Page setup dialog.
- 412d750: Toolbar slot ids now use the DOCX Editor names (`font.family`, `font.size`, `format.clear`, `list.bullet`, `list.numbered`, `list.indent`, `list.outdent`, `list.lineSpacing`, `list.bulletStyle`, `list.numberStyle`), and `Toolbar.FontFamily` has `Trigger`, `Content`, and `Item` parts in React and Vue.
- 412d750: `PptxEditor` and its root now take a ref to a `PptxEditorRef` with `load`, `save`, `exec`, `snapshot`, `focus`, and `getEditor`, the same members in React and Vue.
- 412d750: Add DOCX Editor's font configuration API: the `fonts` option and prop (a `FontConfiguration`, a fragment, or a `FontResolver` such as `packagedFonts()`), `editor.setFonts()`, `onFontError`, `useFonts`, `loadFonts`, `createFontSource`, `customFonts`, and `deferFonts`, and replace `fontProvider` and `addFonts` with them.
- 412d750: Add `focus()` on the editor, superscript and subscript (`script.super`, `script.sub`, Ctrl/Cmd+Shift+= and Ctrl/Cmd+=), the Character Spacing and Change Case menus, the Font dialog (`setFontFormat`, `PptxEditor.FontDialog` with the dialog named parts and `useFontDialog`), and `useFontFormat` in React and Vue; formatting at a caret now applies to the word around it, or to the text typed next at a word edge.
- 412d750: Group undo steps with `Editor.beginHistoryGroup()` and `exec(command, { historyGroup })`, restore the selection after a step on redo, and commit `projection` or `awareness` transactions that add no undo step; `runToolbarCommand` and `useEditorValueCommand().execute` take `{ historyGroup }` last; `historyGroup` replaces the store's `mergeKey` and `setTableGrid.gesture`, `closeHistoryGroup()` replaces `sealHistory()`, and `redo()` returns the selection after the step.
- 412d750: The slide rail takes `orientation="horizontal"` to run as a row under the slide in a tall, narrow host, and it keeps the current slide in view by scrolling only the rail, so a page around the editor no longer moves.
- 412d750: A click into linked text opens a link popup with Copy, Edit, and Remove (a read-only editor shows only the target and Copy), and `PptxEditor.HyperLink` and `useHyperlinkPopup` let you restyle, replace, or rebuild its parts; a `popups.hyperlink` renderer shows this reading mode only when registered with `definePopup(Component, { reading: true })`.
- 412d750: Add a loading screen (`PptxEditor.Loading` in React, `PptxEditorLoading` in Vue) that shows while a deck opens, and paint it before the engine parses a large deck.
- 412d750: `PptxEditor.Menu` and `PptxEditor.ContextMenu` render the menu bar and the canvas menus from a core registry (`CHROME_MENUS`, `CANVAS_CONTEXT_MENUS`) with `slotId` rows, `Group`, `Separator`, and `Submenu`, the `popups` registry gains `hyperlink`, `find`, `contextMenu`, and `railContextMenu`, `PptxEditor.Root` now hosts the link popup, and `useEditorCommand` runs dialog, clipboard, find, link, and slide show slots (in Vue, the clipboard and pictures; the others show as disabled with the reason).
- 412d750: Both adapters export the same props and return types, such as `PptxEditorViewportProps`, `FindBarProps`, `LoadingSpinnerProps`, `UseEditorCommandReturn`, and `UseSlideshowReturn` (the React `*Handle` names stay as deprecated aliases), and `PptxEditor` takes `locale`, `fontFamilies`, and `onChange` as `PptxEditor.Root` does.
- 412d750: `PptxEditor` now shows a notice that says why a file did not open, with `PptxEditor.OpenFailure`, `useOpenFailure`, the `openFailure` option, and `renderOpenFailureActions` (the Vue `openFailureActions` slot) to change it.
- 412d750: The slide rail and the notes pane take DOCX Editor's pane props (`open`, `defaultOpen`, `onOpenChange`, `className`, and `paneWidth` and `toggle` on the rail), the side panel takes `className` and `paneWidth`, `useNotesPane` and `useEditorColorMode` give the notes state and the resolved theme, and `PptxEditor.Root` takes `colorMode` (`'light'`, `'dark'`, or `'system'`) for the chrome, with `colorMode` and `useEditorColorMode` in Vue too.
- 412d750: Add the Paragraph dialog: `setParagraphFormat` sets alignment, indents, space before and after, and line spacing as one undo step, `PptxEditor.ParagraphDialog` (with the dialog named parts and `useParagraphDialog`) opens from Format > Paragraph…, Line spacing options…, and the text context menu, `useParagraphFormat` reads and writes the same fields in React and Vue, with its `ParagraphFormatRead` and `ParagraphFormatCommand` types, and `setLineSpacing` now refuses a multiple below 0.01 or above 9.99, as the dialog does.
- 412d750: Pictures can be cropped, changed, reset, adjusted, styled, sized, and given alt text through new commands, the toolbar's `image` group, and the `imageAltText` and `imageProperties` popups, and picture borders now draw outside the picture.
- 412d750: The Vue adapter, the pro modules for review and collaboration, and PDF export are now published with the other packages.
- 412d750: Add `PptxEditor.Ribbon` and `PptxEditor.StatusBar` (Vue: `PptxEditorRibbon` and `PptxEditorStatusBar`), so your editor can show its commands on tabs, with contextual Format and Table tabs, and slide x of y, notes, views, and zoom under the slide.
- 412d750: Sections show in the rail and the slide sorter and can be added, renamed, removed, moved, and collapsed (`addSection`, `renameSection`, `removeSection`, `removeAllSections`, `moveSection`, `setSectionsCollapsed`, the `section.*` slots), and the new Slide Sorter view (`setPresentationView`, `PptxEditor.SlideSorter`, the `view.presentationView` slot) shows every slide by section. The `railContextMenu` popup entry covers the slide sorter, and the new `sectionContextMenu` entry (`SectionContextMenu`) replaces or turns off the section menu. A slide entry that repeats a slide id of another slide keeps its slide with a fresh id, so the slide list that the editor shows, the commands, and the saved file agree. In the Vue adapter, the `section.*` slots draw as buttons and `view.presentationView` works through `useEditorCommand`. The default editor's zoom pill moved into the status bar, next to the view switch.
- 412d750: The toolbar shows only the controls for the selection (set `contextual={false}` to show every group), and superscript, subscript, character spacing, and columns leave the default bar and `formattingBarChromeGroups()`.
- 412d750: Add shadow, reflection, glow, and soft edge presets, Quick Styles, and commands to distribute, flip, rotate, lock, hide, and rename objects, with toolbar parts that a host places and a typed `useEditorValueCommand` hook.
- 412d750: Add change layout, reset slide, hide slide, and New slide with a layout gallery, through the new `applyLayout`, `resetSlides`, and `setSlidesHidden` commands, the `slide.*` slots, the rail's context menu, `LayoutGallery`, and the `useSlideLayouts` hook in React and Vue.
- 412d750: Tables are editable: type in cells, insert tables, rows, and columns, merge and split cells, and format cells and table styles with new commands and the toolbar's `table` group.
- 412d750: Add font and underline color to the Font dialog (`setFontFormat` `underlineColor`, drawn from `a:uFill`), `PptxEditor.ColumnsDialog` (Format > Columns… and "More columns…") for the column count and spacing, and a custom bullet character and font in the Bullets and Numbering dialog.
- 412d750: Add Text Direction and Columns (`setTextDirection` with Office.js orientation words (`TextOrientation`), `setColumns` with `count` and `gapEmu`, and the `paragraph.textDirection` and `paragraph.columns` menus in the React toolbar), and resize-to-fit and shrink-on-overflow that follow rotated text, columns, and the text rectangle of shapes that are not rectangles; `mongolianVertical` and `wordArtVerticalRtl` are refused and listed as approximations.
- 412d750: The toolbar gains DOCX Editor's named parts (`Toolbar.Bold`, `Toolbar.FontFamily`, …), `Toolbar.Action` and `Toolbar.Separator`, group `priority`, `pinned`, `after`, `className`, and `overflowContent`, and `overflow={false}` to wrap the bar instead of using a More menu; the Vue toolbar gets the named parts, `ToolbarAction`, `ToolbarSeparator`, `after`, and `className`, and its bar wraps until it has a More menu.
- 412d750: The Vue adapter now has every part of the React adapter under the same names, including the menu bar, context menus, ribbon, status bar, slide rail, slide sorter, notes pane, side panel, find bar, slide show, form dialogs, `Icon`, `MenuRow`, `MenuSubmenu`, and `useNotesPane`, and its parts take `className` and `style` as the React parts do.

### Patch Changes

- 412d750: Install the React or Vue adapter with `@pptx-editor.dev/core` only: the adapter now installs `@pptx-editor.dev/i18n` for you.
- 412d750: Tab and Shift+Tab move the selection through the shapes on the slide, a table cell gets its own context menu (`CanvasMenuKind` `table`) with row and column commands, and the canvas keeps the caret in view, places IME input at the caret, shows shapes outside the slide dimmed, and zooms in 10 % wheel steps within one zoom range.
- 412d750: Controls show a styled tooltip with the reason a control is off, the font picker groups the theme, deck, and other fonts and applies a theme font as a theme reference (`SelectionFormat.fontFamilyTheme`), color menus take a custom hex color, menus close on any press outside, a menu opens on its checked row, a selected section shows its first slide, Change case and Insert link act on the word at the caret, and undo, find, notes formatting, and keyboard focus work as expected.
- 3bf89a8: Menus, popovers, and tooltips now open at their controls in a transformed, scaled, or clipping host container. A host CSS reset no longer changes the chrome, and Ctrl/Cmd+F and the link popup now stay with their own editor.
- 412d750: Find counts and highlights each match in the speaker notes, Redo shows ⇧⌘Z on macOS, the custom color field shows the current color, and the slide sorter shows no zoom percentage.
- 412d750: The font notice drops the result of an earlier local font request when another presentation opens, and dialog buttons get a destructive style and press feedback.
- 412d750: Escape in a menu bar menu now returns focus to its menu bar item, the open find bar no longer covers the slide, and a disabled control's reason is read once.
- 412d750: Describe every feature by its ECMA-376 behavior in docs and labels: the export menu item reads "Presentation (.pptx)", the diagram placeholder labels read "Diagram", and the diagram placeholder prompt now uses the i18n key `canvas.prompts.diagram`.
- 412d750: Bold, italic, and underline for the speaker notes say why they are off and follow every selection change, Tab and Escape leave the notes after typing, and a locale change on the root repaints the canvas placeholder prompts once, also with an inline `locale` object.
- 412d750: The picture Corrections control shows its brightness and contrast presets as a grid of preview samples instead of a long list.
- 412d750: Clarify package documentation and host responsibilities for collaboration access.
- 412d750: Text typed over a selection is one undo step, a quick drag in the slide rail or sorter reorders slides, the Format menu names its picture row Format picture, and the find bar moves the slide only when it would cover it.
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
