# Tasks

## Planning

- [x] Read the DOCX Editor document refresh guide, types, and command contract.
- [x] Create the proposal, design, and specification delta.
- [x] Self-review against DOCX Editor's code. Apply the findings to the proposal and design.
- [x] Confirm the license boundary: change reporting in Pro, matcher and controller in `core`.
- [ ] Resolve review findings and the open questions in `design.md`.

## Stage 1: comparison, reload mode, highlights

- [x] Copy the shared contracts from DOCX Editor with source headers: refresh types, `AnchorHighlightOptions`, `ScrollToAnchorOptions`, `DocumentHandle`, `DocumentChange`.
- [x] Add `editor.getDocumentHandle()` from the store revision and a load id.
- [x] Pass `DocumentChange` to `change` listeners, with `source` for load, refresh, and recovery.
- [x] Implement the package matcher in a DOM-free `core` lane: parts, slides, and shapes, groups included. Record the lane in `core-lane-graph.ts`. Done in `store/store/package-diff.ts`, inside the existing `store` lane; no new lane.
- [x] Add the `refresh` hook to `EditorModule`.
- [x] Implement change reporting in `@pptx-editor.dev/pro`: aspects, text ranges, notes, layouts, masters, themes, and stable ids.
- [x] Register change reporting in `reviewModule()`.
- [x] Export `comparePresentations` from `@pptx-editor.dev/pro` and test it on Node.js.
- [x] Implement `createDocumentRefresh` with reload mode, sequence rules, refusals, cancel, and recovery.
- [x] Port the DOCX Editor controller tests for the shared behavior.
- [x] Validate `RefreshLocation` and report diagnostics.
- [x] Reload in place: open the returned file in a second session off the page, keep the loading screen hidden, and swap in one frame. Done differently: the returned file applies in place as a diff without history, so no second session is needed. Fallback: a quiet whole-file load without the loading screen.
- [ ] Paint the new current slide off the page and wait for its font faces before the swap. Not done.
- [x] Keep thumbnails of unchanged slides. Repaint changed thumbnails off the page. Never empty a painted container during a refresh. Unchanged parts keep their identity, so their thumbnails and canvas items do not repaint. Changed thumbnails repaint in place.
- [ ] Key rail and sorter items by slide match in React and Vue, and keep the rail scroll position. Rail items keep their slide id key, and `e2e/document-refresh.spec.ts` shows that no thumbnail node is replaced. Gap: after an inserted slide, the thumbnail of each later slide paints again in place at its new index. The rail scroll position is not tested.
- [x] Split parse, comparison, layout, and paint into steps under 50 milliseconds. `applyUpdate()` reads the file, takes the difference, commits with layout and paint, and computes the changes in separate tasks, and saves nothing during an update. Layout and paint of the commit stay one step. In Chromium, no task exceeds 50 ms at 20, 100, and 300 slides (`bun run bench:refresh:browser`). Before the split, a 100-slide reload took one task of 64 to 71 ms.
- [x] Keep the old paint and return `load-failed` with `recovered: true` when the new file fails.
- [x] Add a Playwright frame test for both modes and both adapters: no loading screen, no empty canvas, no emptied thumbnail, and no long task. `e2e/document-refresh.spec.ts` checks these on 20 slides, and long tasks on 100 slides, in React and Vue.
- [x] Flash changed thumbnails in the rail and sorter, with the ring, wash, and "Changed" marker.
- [x] Flash shape rings and text bands on the canvas, and the notes pane, from one start time. The notes pane gets a ring around the pane, not a text band.
- [x] Add `RefreshHighlightOptions.surfaces` and the `pptx-refresh-*` classes. Apply reduced motion.
- [x] Test that a thumbnail that scrolls into view during the hold shows the remaining flash only. `e2e/document-refresh.spec.ts` scrolls an offscreen rail item into view and opens the slide sorter during the hold.
- [x] Implement `navigateToChange` with `block`, `behavior`, `offsetPx`, and `focus`.
- [x] Add `refresh.change.*` strings to `en.json` and the `ja`, `ko`, and `zh-CN` catalogs.
- [x] Add a benchmark for comparison time on large fixtures. Set the time budget from it. `bun run bench:refresh` times each step at 20, 100, and 300 slides, and fails when a step median exceeds 50 ms at 100 slides. At 300 slides, the read, the comparison, and the change computation each take 25 to 30 ms.

## Stage 2: edit mode

- [x] Verify whether one transaction can remove and add the same part name. Add a `replacePart` op if it cannot.
- [x] Implement `replaceDocument` from the comparison, with part replacement as fallback.
- [ ] Map the current slide, the shape selection, and the text caret after replacement. The current slide follows its slide part, also to a new index. The shape selection and the caret stay when their nodes stay, with the caret clamped into its text. Mapping the caret through changed text is not done: an edit before the caret does not move it.
- [x] Connect `applyUpdate({ mode: 'edit' })` to `replaceDocument`.
- [x] Test undo, redo, save equality with `load()`, unknown content, and refusals.

## Stage 3: collaboration

- [x] Run `replaceDocument` with an attached `collaborationModule`.
- [x] Refuse reload with `collaboration`, and map room refusals to `collaboration-refused`.
- [ ] Add convergence tests with concurrent edits, offline edits, and undo. Concurrent and undo tests exist in `replace-document-peers.test.ts`. Offline edits are not tested.
- [x] Update the collaboration baseline if the op set changes. No op set change.

## Documentation and release

- [x] Write `docs/site/content/guides/document-refresh.mdx` with the DOCX Editor page structure, `FrameworkTabs` samples, and a highlight simulation. Done without the highlight simulation component.
- [x] Add `comparePresentations` to the Pro documentation.
- [x] Add `examples/document-refresh/`, as in DOCX Editor. It has the main page with both update modes and an agent replay page (`agent.html`) that applies seven scripted steps with computed changes, Undo step, and Restart.
- [x] Add the `replaceDocument` section to `docs/site/content/guides/loading-and-saving.mdx`.
- [x] Register the page in both `meta.json` files.
- [x] Add a feature matrix row with evidence fixtures.
- [ ] Apply a server update with highlights in `examples/chat/`. Not done.
- [x] Update `docs/repo-map.md` for new modules.
- [x] Build packages and update API snapshots.
- [x] Add a changeset.
- [ ] Run the required checks and record results, skipped checks, and gaps in `verification.md`. Not done.
