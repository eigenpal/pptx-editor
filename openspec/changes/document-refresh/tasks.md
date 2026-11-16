# Tasks

## Planning

- [x] Read the DOCX Editor document refresh guide, types, and command contract.
- [x] Create the proposal, design, and specification delta.
- [x] Self-review against DOCX Editor's code. Apply the findings to the proposal and design.
- [x] Confirm the license boundary: change reporting in Pro, matcher and controller in `core`.
- [ ] Resolve review findings and the open questions in `design.md`.

## Stage 1: comparison, reload mode, highlights

- [ ] Copy the shared contracts from DOCX Editor with source headers: refresh types, `AnchorHighlightOptions`, `ScrollToAnchorOptions`, `DocumentHandle`, `DocumentChange`.
- [ ] Add `editor.getDocumentHandle()` from the store revision and a load id.
- [ ] Pass `DocumentChange` to `change` listeners, with `source` for load, refresh, and recovery.
- [ ] Implement the package matcher in a DOM-free `core` lane: parts, slides, and shapes, groups included. Record the lane in `core-lane-graph.ts`.
- [ ] Add the `refresh` hook to `EditorModule`.
- [ ] Implement change reporting in `@pptx-editor.dev/pro`: aspects, text ranges, notes, layouts, masters, themes, and stable ids.
- [ ] Register change reporting in `reviewModule()`.
- [ ] Export `comparePresentations` from `@pptx-editor.dev/pro` and test it on Node.js.
- [ ] Implement `createDocumentRefresh` with reload mode, sequence rules, refusals, cancel, and recovery.
- [ ] Port the DOCX Editor controller tests for the shared behavior.
- [ ] Validate `RefreshLocation` and report diagnostics.
- [ ] Reload in place: open the returned file in a second session off the page, keep the loading screen hidden, and swap in one frame.
- [ ] Paint the new current slide off the page and wait for its font faces before the swap.
- [ ] Keep thumbnails of unchanged slides. Repaint changed thumbnails off the page. Never empty a painted container during a refresh.
- [ ] Key rail and sorter items by slide match in React and Vue, and keep the rail scroll position.
- [ ] Split parse, comparison, layout, and paint into steps under 50 milliseconds.
- [ ] Keep the old paint and return `load-failed` with `recovered: true` when the new file fails.
- [ ] Add a Playwright frame test for both modes and both adapters: no loading screen, no empty canvas, no emptied thumbnail, and no long task.
- [ ] Flash changed thumbnails in the rail and sorter, with the ring, wash, and "Changed" marker.
- [ ] Flash shape rings and text bands on the canvas, and bands in the notes pane, from one start time.
- [ ] Add `RefreshHighlightOptions.surfaces` and the `pptx-refresh-*` classes. Apply reduced motion.
- [ ] Test that a thumbnail that scrolls into view during the hold shows the remaining flash only.
- [ ] Implement `navigateToChange` with `block`, `behavior`, `offsetPx`, and `focus`.
- [ ] Add `refresh.change.*` strings to `en.json` and the `ja`, `ko`, and `zh-CN` catalogs.
- [ ] Add a benchmark for comparison time on large fixtures. Set the time budget from it.

## Stage 2: edit mode

- [ ] Verify whether one transaction can remove and add the same part name. Add a `replacePart` op if it cannot.
- [ ] Implement `replaceDocument` from the comparison, with part replacement as fallback.
- [ ] Map the current slide, the shape selection, and the text caret after replacement.
- [ ] Connect `applyUpdate({ mode: 'edit' })` to `replaceDocument`.
- [ ] Test undo, redo, save equality with `load()`, unknown content, and refusals.

## Stage 3: collaboration

- [ ] Run `replaceDocument` with an attached `collaborationModule`.
- [ ] Refuse reload with `collaboration`, and map room refusals to `collaboration-refused`.
- [ ] Add convergence tests with concurrent edits, offline edits, and undo.
- [ ] Update the collaboration baseline if the op set changes.

## Documentation and release

- [ ] Write `docs/site/content/guides/document-refresh.mdx` with the DOCX Editor page structure, `FrameworkTabs` samples, and a highlight simulation.
- [ ] Add `comparePresentations` to the Pro documentation.
- [ ] Add `examples/document-refresh/`, as in DOCX Editor.
- [ ] Add the `replaceDocument` section to `docs/site/content/guides/loading-and-saving.mdx`.
- [ ] Register the page in both `meta.json` files.
- [ ] Add a feature matrix row with evidence fixtures.
- [ ] Apply a server update with highlights in `examples/chat/`.
- [ ] Update `docs/repo-map.md` for new modules.
- [ ] Build packages and update API snapshots.
- [ ] Add a changeset.
- [ ] Run the required checks and record results, skipped checks, and gaps in `verification.md`.
