# Document refresh

## Why

Hosts generate and change presentations on a server, for example with an agent. The user keeps the presentation open in the editor. The host must show the returned file without a full reset, and the user must see what changed.

DOCX Editor ships this as the document refresh API (`createDocumentRefresh`) and the undoable `replaceDocument` command. PPTX Editor has no equivalent. The only file import is `Presentation.insertSlidesFromBase64`, which adds slides and does not replace or compare a presentation.

## Change

Ship the DOCX Editor document refresh API in PPTX Editor with the same names, types, methods, refusal codes, and lifecycle. Code that a host writes for DOCX Editor's controller must work with PPTX Editor's controller, except for the location fields that describe a place in a presentation.

1. `createDocumentRefresh(editor)` returns a `DocumentRefresh` controller: `capture()`, `applyUpdate()`, `finish()`, `cancel()`, `snapshot()`, `subscribe()`, `onResult()`, `highlightChanges()`, `clearHighlights()`, `navigateToChange()`, `recover()`, and `recoveryBytes()`.
2. `applyUpdate({ mode: 'reload' })` reloads the returned file and keeps the current slide and the scroll position where they map.
3. `applyUpdate({ mode: 'edit' })` runs a new `replaceDocument` command. The command applies the difference between the open package and the returned package as one transaction of `PptxOp`s. One undo returns to the captured presentation. A collaboration session shares the edit as an ordinary edit.
4. `editor.getDocumentHandle()` returns `{ loadId, revision }`, and the `change` event receives a `DocumentChange` with `source: 'load' | 'refresh' | 'recovery'`.
5. Updates are smooth in both modes. The editor shows no loading screen, no blank frame, and no font swap, and it runs no long task. It swaps the old paint for the new paint in one frame and keeps unchanged thumbnails.
6. `highlightChanges()` flashes the thumbnails of changed slides in the rail and the sorter. It also flashes changed shapes and text on the current slide. The flash is editing chrome only.

## Differences from DOCX Editor

These differences come from the format. Each one is additive or is limited to location fields.

- **Locations.** `RefreshLocation` addresses a slide, a shape on that slide, and optionally a text range in one paragraph of that shape. DOCX Editor's location addresses a body paragraph.
- **Source of computed changes.** In DOCX Editor, when `changes` is omitted, a registered `reviewModule()` supplies the tracked revisions of the returned file. PresentationML has no revision markup. In PPTX Editor, a registered `reviewModule()` compares the captured and returned packages instead. `RefreshChange.computed` describes each computed change in structured fields that the adapters localize. Without the module, the result reports `changeInformation: 'unavailable'`, as in DOCX Editor. Server metadata, when present, replaces the computed list.
- **`comparePresentations(original, revised)`.** A DOM-free export of the same comparison from `@pptx-editor.dev/pro`, so a server can compute, filter, and describe changes before it sends them. It has the place and role of DOCX Editor's `compareDocuments`, but it returns a change list, not redline bytes.
- **Command field.** `replaceDocument` takes `pptx` where DOCX Editor takes `docx`.

Change reporting is in `@pptx-editor.dev/pro` under the EigenPal Pro License, because the DOCX Editor equivalents (`compareDocuments` and change derivation in `reviewModule()`) are Pro. The controller, `replaceDocument` with its package matcher, highlights, and navigation are in `core`, as in DOCX Editor.

## Effect on DOCX Editor

This change needs no DOCX Editor change. Shared host code compiles against both controllers, because `computed` is an optional member. DOCX Editor hosts get change locations today through tracked revisions, for example from a server `compareDocuments` call. A DOCX Editor change that highlights computed changes without tracked changes is a possible follow-up. `design.md` records it as an open question.

## Scope

- Core: the package matcher, `replaceDocument`, the controller, the document handle, the `change` source, a module hook for computed changes, and highlight and navigation overlays.
- React and Vue: re-export `createDocumentRefresh` from both adapters, with equal behavior, and render highlights in the rail, sorter, and canvas.
- Pro: change reporting, `comparePresentations`, computed changes in `reviewModule()`, and `replaceDocument` with an attached `collaborationModule`, with convergence checks.
- Documentation: `guides/document-refresh.mdx` with the same page name, structure, and simulation as DOCX Editor's, the `replaceDocument` section of `guides/loading-and-saving.mdx`, a feature matrix row with evidence, translations, and a changeset.
- Examples: `examples/document-refresh/`, as in DOCX Editor, and a server-generated update with highlights in the chat example.

## Out of scope

- A public Office.js method for whole-presentation replacement. The Office.js `PowerPoint` namespace has no such method, and the editor API does not invent one.
- Tracked changes for presentations. PresentationML has no revision markup. Computed changes are not saved and do not reach other participants.
- Change detection inside embedded objects (OLE), media, and fonts. A changed binary part reports one change for the part, not a location inside it.

## Delivery

Deliver in three reviewed stages, as `tasks.md` lists:

1. The controller in reload mode, the comparison, highlights, and navigation.
2. `replaceDocument` and edit mode.
3. Edit mode in collaboration sessions.

Keep review findings and verification results in this directory. Do not claim completion while a required task remains open.
