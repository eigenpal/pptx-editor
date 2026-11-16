## ADDED Requirements

### Requirement: Document refresh API parity

PPTX Editor SHALL export `createDocumentRefresh` and the DOCX Editor document refresh types with the same names, members, refusal codes, and lifecycle. Only these differences SHALL exist: the fields of `RefreshLocation`, the `pptx` field of `ReplaceDocumentRequest`, and optional additions such as `RefreshChange.computed`, `RefreshHighlightOptions.surfaces`, and added diagnostic codes. React and Vue SHALL re-export the same API with equal behavior.

#### Scenario: Shared host code

- **WHEN** a host calls `capture()`, `applyUpdate()`, and `finish()` as the DOCX Editor guide shows
- **THEN** the calls compile and behave as in DOCX Editor
- **AND** each refusal returns the same `code` for the same situation

#### Scenario: Stale result

- **WHEN** the user edits the presentation after `capture()`, then undoes the edit
- **THEN** `applyUpdate()` for that submission returns `local-edits`
- **AND** the presentation does not change

#### Scenario: Sequence order

- **WHEN** a result arrives with a sequence equal to or lower than the last accepted sequence
- **THEN** `applyUpdate()` returns `out-of-order`
- **AND** the presentation does not change

### Requirement: Computed changes

When an update omits `changes` and `reviewModule()` is registered, the controller SHALL report changes from the captured package to the returned package at slide, shape, and text level. Each computed change SHALL carry `computed` fields and a stable id across cumulative results. Supplied `changes` SHALL replace the computed list. `comparePresentations` SHALL produce the same changes without a DOM.

#### Scenario: No module

- **WHEN** an update omits `changes` and no module supplies changes
- **THEN** the result reports `changeInformation: 'unavailable'`
- **AND** `changes` is empty

#### Scenario: Cumulative results

- **WHEN** two cumulative results both contain the same text change
- **THEN** the change has the same id in both results
- **AND** `isNew` is false for it in the second result

#### Scenario: Changed text

- **WHEN** the returned file changes one word in one shape
- **THEN** the result has one available change of scope `text`
- **AND** its location names the slide, the shape, the paragraph, and the changed range

#### Scenario: Added and removed slides

- **WHEN** the returned file adds one slide and removes another
- **THEN** the result has one `added` slide change with a location
- **AND** one `removed` slide change with status `deleted`

#### Scenario: Layout change

- **WHEN** the returned file changes a layout that three slides use
- **THEN** the result has one `layout` change
- **AND** `affectedSlides` lists the three slides

#### Scenario: Ambiguous shapes

- **WHEN** a slide's shape ids are duplicated so that shapes cannot match
- **THEN** the result reports that slide as one `slide` change
- **AND** it reports no shape-level guess

#### Scenario: Unchanged parts

- **WHEN** the returned file equals the captured file
- **THEN** the result has no changes

### Requirement: Undoable replacement

`replaceDocument` and `applyUpdate({ mode: 'edit' })` SHALL apply the difference as one store transaction. One undo SHALL restore the captured presentation. A save after the replacement SHALL equal a save of the returned file opened with `load()`, with the same binary parts and the same canonical XML parts. The controller SHALL NOT report `local-edits` for its own accepted update. A refused replacement SHALL leave the presentation and its history unchanged.

#### Scenario: Undo and redo

- **WHEN** an update with `mode: 'edit'` succeeds
- **THEN** one undo restores the captured presentation
- **AND** one redo applies the returned presentation again
- **AND** earlier undo steps stay available

#### Scenario: Selection survives

- **WHEN** the selected shape exists in the returned file
- **THEN** the shape stays selected after the update
- **AND** the current slide stays the same

#### Scenario: Unknown content

- **WHEN** the returned file holds extension content that the engine does not interpret
- **THEN** the replacement keeps that content
- **AND** save and reopen keep it

#### Scenario: Change source

- **WHEN** an update applies
- **THEN** the `change` event receives `source: 'refresh'`
- **AND** user edits, undo, and redo receive no `source`

### Requirement: Collaborative replacement

With a collaboration session attached, an update with `mode: 'edit'` SHALL reach other participants as an ordinary edit and converge with their concurrent edits. A reload SHALL refuse with `collaboration`.

#### Scenario: Convergence

- **WHEN** one participant applies an update while another participant edits a different slide
- **THEN** all participants converge to the same presentation
- **AND** undo on the first participant reverts only the update

### Requirement: Smooth update

An accepted update SHALL replace the visible presentation without a blank frame, a loading screen, a fallback font swap, or a long task. The old paint SHALL stay until the new current slide is painted, and the swap SHALL happen in one animation frame. Unchanged thumbnails SHALL keep their paint, and the view state SHALL stay.

#### Scenario: No blank frame

- **WHEN** an update in either mode is accepted
- **THEN** every recorded frame shows either the old or the new current slide
- **AND** no frame shows the loading screen or an empty canvas

#### Scenario: Unchanged thumbnails

- **WHEN** an update changes one slide of a long presentation
- **THEN** only that slide's thumbnail paints again
- **AND** no painted thumbnail container is emptied

#### Scenario: Responsive input

- **WHEN** an update applies to a benchmark presentation
- **THEN** no task of the update takes more than 50 milliseconds
- **AND** the rail scroll position, the zoom, and the current slide stay

#### Scenario: Failed update

- **WHEN** the returned file cannot open
- **THEN** the old paint stays on screen
- **AND** `applyUpdate()` returns `load-failed` with `recovered: true`

### Requirement: Change presentation

Highlights SHALL mark available changes on the canvas, in the slide rail, in the slide sorter, and in the notes pane. Each mark SHALL have a non-color signal and an accessible label. Highlights SHALL add no saved bytes, no undo entries, and no paint in the slide show or exports. A document edit SHALL remove them.

#### Scenario: Navigate to a change

- **WHEN** the host calls `navigateToChange(id, { focus: true })` for a shape change on another slide
- **THEN** the editor shows that slide
- **AND** the shape is selected
- **AND** the call returns `true`

#### Scenario: Flash changed thumbnails

- **WHEN** an update changes slides 2 and 5, and the host calls `highlightChanges()`
- **THEN** the thumbnails of slides 2 and 5 flash in the rail and the sorter
- **AND** each shows a marker with the accessible label "Changed"
- **AND** no other thumbnail flashes

#### Scenario: Flash the current slide

- **WHEN** the update changes text in a shape on the current slide, and the host calls `highlightChanges()`
- **THEN** a band flashes over the changed text on the canvas
- **AND** no slide or thumbnail paints again for the flash

#### Scenario: Reduced motion

- **WHEN** the user prefers reduced motion
- **THEN** the flash fades take at most 125 milliseconds
- **AND** the ring and the marker still show

#### Scenario: Unavailable change

- **WHEN** the host calls `navigateToChange` for a removed slide
- **THEN** the call returns `false`
- **AND** the view does not change

#### Scenario: Export after highlight

- **WHEN** highlights are visible and the host saves or exports the presentation
- **THEN** the output contains no highlight

### Requirement: Refresh security

The controller and `comparePresentations` SHALL read returned files through `readPptx` and its limits. They SHALL enforce change, description, text comparison, and time limits. They SHALL NOT fetch external relationships or run embedded content.

#### Scenario: Hostile file

- **WHEN** the returned file exceeds a package limit
- **THEN** `applyUpdate()` returns `invalid-document`
- **AND** the open presentation stays visible
