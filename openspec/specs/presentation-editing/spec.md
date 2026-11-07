# presentation-editing Specification

## Purpose

Created by archiving `pptx-editor-foundation`, `group-and-multi-edit` on 29 October 2026.

## Requirements

### Requirement: Store-only mutation

Every edit SHALL be a transaction of serializable ops on the document store. Unsupported edits SHALL be refused with a typed reason.

#### Scenario: Refused edit

- **WHEN** a command targets an element the engine cannot edit
- **THEN** the store does not change and the command reports a reason

### Requirement: Re-layout after edits

Edits SHALL re-lay out the affected slide with the autofit rules of its text bodies (`a:normAutofit`, `a:spAutoFit`).

#### Scenario: Typing in an autofit shape

- **WHEN** text grows in a shape with `spAutoFit`
- **THEN** the shape height grows to fit the text

### Requirement: Undo and redo

Each user action SHALL be one history entry, and typing SHALL group by intent.

#### Scenario: Undo typing

- **WHEN** a user types a word and presses undo
- **THEN** the whole word is removed in one step

### Requirement: Group scope

The editor SHALL let a user enter a group and edit its children, and SHALL keep the group's bounds fitted to its children.

#### Scenario: Move a child

- **WHEN** a user enters a group and drags one child
- **THEN** only that child moves, and the group's bounds become the union of its children

### Requirement: Multi-selection editing

Moving, nudging, aligning, or resizing several selected shapes SHALL be one undoable transaction.

#### Scenario: Move three shapes

- **WHEN** three shapes are selected and dragged
- **THEN** all three move by the same distance, and one undo restores all three

### Requirement: Shape drags write once

A pointer drag that moves, resizes, or rotates shapes, or moves a line end, SHALL change the deck once, on release. While the pointer moves, the editor SHALL show a preview and SHALL NOT write to the store.

#### Scenario: Drag and release

- **WHEN** a user drags a shape across the slide and releases the pointer
- **THEN** the store gets one transaction, collaboration peers get one set of ops, and one undo restores the shape

#### Scenario: Escape or focus loss cancels

- **WHEN** a user presses Escape, or the window loses focus, after the pointer moved and before the release
- **THEN** the preview goes, the deck does not change, and the live region says that the drag was canceled

#### Scenario: An edit under the drag cancels

- **WHEN** an undo or a redo runs, or a command or a collaborator changes a dragged shape or a group that holds it, before the release
- **THEN** the preview goes, and the release writes nothing

#### Scenario: An edit to another shape keeps the drag

- **WHEN** a command or a collaborator changes another shape, or another part of the deck, before the release
- **THEN** the preview stays, and the release writes the drag as one edit

#### Scenario: The browser cancels the pointer

- **WHEN** the browser cancels the pointer during a drag (`pointercancel`)
- **THEN** the drag writes its last preview as one undo step

### Requirement: Deferred thumbnails

The slide rail and the slide sorter SHALL NOT paint every thumbnail before the editor can be used. They SHALL ask for deferred thumbnails (`renderThumbnail` with `defer`). The current slide's thumbnail and the thumbnails in view SHALL paint at once. The other thumbnails SHALL paint when they scroll into view or when the browser is idle. A thumbnail SHALL NOT show another slide than its own. A `renderThumbnail` call without `defer` SHALL paint at once.

#### Scenario: A host paints thumbnails

- **WHEN** a host calls `renderThumbnail` for each slide with no options and reads the containers in the same task
- **THEN** every container holds its slide

#### Scenario: Open a long deck

- **WHEN** a deck of 800 slides opens
- **THEN** the rail paints about one screen of thumbnails before the editor can be used, and the rest later

#### Scenario: A slide is deleted above the view

- **WHEN** the rail is scrolled down and a slide above the view is deleted or restored
- **THEN** no thumbnail in view is empty at a frame

### Requirement: Group and ungroup keep positions

Grouping and ungrouping SHALL NOT move any shape on the slide.

#### Scenario: Ungroup a rotated group

- **WHEN** a rotated group is ungrouped
- **THEN** every child renders where it rendered inside the group, with the group's rotation added to its own

### Requirement: Clipboard

Copying and pasting shapes SHALL carry their pictures and refuse to carry external links.

#### Scenario: Paste a picture into another deck

- **WHEN** a picture is copied from one deck and pasted into another
- **THEN** the second deck gains the image part, a relationship to it, and the picture renders

### Requirement: Chart data editing

An edit of chart data SHALL change the series names, category labels, values, point counts, and cell references of the chart part in one undo step. It SHALL write the same cells into an embedded workbook only when the workbook's structure allows it, and SHALL otherwise keep the workbook unchanged. It SHALL turn off the automatic update of a linked workbook, and SHALL never open it.

#### Scenario: Add a category

- **WHEN** a category is added to a column chart with an embedded workbook
- **THEN** every series gains a cached point, its references cover the new row, and the workbook gains the same row

#### Scenario: Remove a category with point formatting

- **WHEN** a category is removed from a series with a formatted data point after it
- **THEN** the formatted point moves to its new index, and the formatting of the removed point goes away

#### Scenario: A workbook with formulas

- **WHEN** the chart's embedded workbook has a formula in its data cells
- **THEN** the chart part changes, and the workbook saves unchanged

### Requirement: Progressive open

A long deck SHALL show its start slide and take input before its other slides parse, and SHALL parse them in background tasks that give way to input. Every read, edit, undo, search, and save before then SHALL give the result of a deck that parsed whole, and a save SHALL parse the rest first. A slide that does not parse SHALL refuse the deck (or, after edits, make it read-only), and SHALL never be saved. An editor with a collaboration module SHALL parse the whole deck before its first slide shows.

#### Scenario: Edit before the deck is ready

- **WHEN** a user types on the first slide of a long deck and saves before the other slides parse
- **THEN** the saved file holds every slide, and is the file a whole open with the same edits saves

#### Scenario: A damaged slide in a long deck

- **WHEN** a slide or a notes slide that has not parsed yet holds XML that does not parse, and the user has not edited the deck
- **THEN** the deck is refused when the slide parses, and a save before then is refused too

#### Scenario: A damaged slide after edits

- **WHEN** a slide or a notes slide of a long deck that the user has edited does not parse
- **THEN** the deck stays open in view mode with its edits, and `readOnlyReason` is `damaged-slide` while `openFailure` stays `null`
- **AND** commands, undo, save, and automation writes and exports are refused, and the other slides go on opening in background tasks

#### Scenario: A damaged slide that the user deleted

- **WHEN** the user deletes a slide that has not parsed yet and does not parse
- **THEN** the deck is not refused, and saves without it, unless an undo brings the slide back and a reader reads it
