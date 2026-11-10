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

#### Scenario: Undo text copied by a peer

- **WHEN** a participant undoes typing after another participant formats or splits that text
- **THEN** known copies of those typed characters are removed, and unrelated characters remain

#### Scenario: Deliver an undo in separate updates

- **WHEN** a provider delivers updates for an undo separately
- **THEN** each restored character arrives with its identity and deletion metadata

#### Scenario: Concurrent table removal and text editing

- **WHEN** a row or cell disappears while another participant edits its text
- **THEN** surviving cells retain peer text, and explicit text deletions still apply

#### Scenario: Recover into a cell without a text body

- **WHEN** a concurrent row or column deletion removes a cell with unseen peer typing
- **AND** the surviving destination cell has no text body
- **THEN** recovery places the source text body and typing into that cell before another edit
- **AND** text already observed by the deletion does not qualify the paragraph for recovery

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

### Requirement: Canvas selection

Node ids are unique in one part only. A layout or master shape can have the node id of a slide shape. A press SHALL select the slide shape under the pointer. Selections, commands, and selection marks SHALL name slide shapes only. The canvas around the slide SHALL act as empty slide area. A drag there SHALL select the slide shapes that it encloses. A click there SHALL clear the selection. A right-click there SHALL open the slide menu. When several shapes are selected, the outline of each shape SHALL paint over the common box.

#### Scenario: A slide shape shares its id with a layout shape

- **WHEN** a user clicks a slide title whose node id is also the node id of a layout header
- **THEN** the title is selected, and its outline is on the title's box

#### Scenario: Drag from outside the slide

- **WHEN** a user drags from the canvas above and left of the slide across two shapes
- **THEN** the two shapes are selected

#### Scenario: Click or right-click outside the slide

- **WHEN** a shape is selected and a user clicks or right-clicks the canvas around the slide
- **THEN** the selection clears, and a right-click opens the slide menu

#### Scenario: A selected shape is removed

- **WHEN** a collaborator removes a selected slide shape whose node id is also the node id of a layout shape
- **THEN** the selection clears, and no mark shows on the layout shape

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

### Requirement: Equation editing

Insert Equation and Edit Equation SHALL write Office Math from a linear form in one undo step, and SHALL refuse a linear form that the parser does not read, with no change. A new equation SHALL be a text box in `mc:AlternateContent`, with the math in a `mc:Choice` that requires `a14` and the equation as plain text in a `mc:Fallback`. An edit of the math SHALL keep the alignment and the first run's format, SHALL remove the stale fallback, and SHALL refuse an equation whose math the linear form cannot write.

#### Scenario: Insert an equation

- **WHEN** `insertEquation` runs with the linear form `A=\pi r^2`
- **THEN** the slide gains a selected text box whose math has a superscript, and its fallback holds `A=πr^2`

#### Scenario: Edit an equation

- **WHEN** `setEquation` runs on a selected equation with the linear form `y=(a+1)/b`
- **THEN** the math is a fraction, the run font stays, and the wrapper has no fallback

#### Scenario: Unreadable linear form

- **WHEN** `insertEquation` runs with the linear form `x^`
- **THEN** the command refuses, and the undo history does not change

### Requirement: Diagram editing

A text edit in a diagram shape SHALL write the drawing shape's text and the text of the data model point that the shape shows, in one transaction. An added or deleted diagram shape SHALL change the data model's points and connections, renumber the sibling order, and lay the saved drawing out again, in one undo step. A diagram whose saved drawing is not one line of shapes with up to one level of children SHALL refuse structure edits, and SHALL write nothing.

#### Scenario: Type in a diagram shape

- **WHEN** text is typed at the end of a process shape
- **THEN** the drawing shape and its data point hold the new text, with the run formatting they had

#### Scenario: Add a shape after a node

- **WHEN** a shape is added after the third node of a four-node process
- **THEN** the data model holds five top-level nodes in order, each named by one connection, and the drawing holds five boxes and four arrows

#### Scenario: Delete a node

- **WHEN** a node is deleted
- **THEN** its subtree, transition points, connections, and presentation points leave the data model, and no connection names a missing point

### Requirement: Slide master view

The slide master view SHALL show a slide layout or a slide master on the canvas, and SHALL edit its own shapes, placeholders, and background through `PptxDocumentStore` operations on its part. A layout SHALL show its master's shapes under its own shapes. Every slide that uses an edited layout or master SHALL show the edit. Commands that act on slides SHALL refuse with `disabled.slideMaster` while the view is open. Closing the view SHALL return to the slide that was current when it opened. Undo, redo, and collaboration SHALL treat a template edit as any other edit of the deck.

#### Scenario: Move a layout shape

- **WHEN** a user opens the view on a slide, and moves a shape of its layout
- **THEN** the layout part holds the new position, and every slide on that layout shows the shape there

#### Scenario: A slide command in the view

- **WHEN** a user adds a slide, writes speaker notes, or starts Find while the view is open
- **THEN** the command is refused with `disabled.slideMaster`, and the deck does not change

#### Scenario: A layout edit in a collaboration room

- **WHEN** one collaborator moves a layout shape in the view while another moves a slide shape on that layout
- **THEN** both replicas converge, and the slide shows the moved layout shape on both

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

### Requirement: Collaboration format compatibility

Rooms SHALL record protocol, shared schema, repair, and canonical model versions. Sessions SHALL reject mismatched fields before applying room content and after metadata changes. Format inspection SHALL preserve shared state and SHALL support legacy version strings. Hosts SHALL check connection compatibility before synchronization.

#### Scenario: Upgrade an incompatible saved room

- **WHEN** a release changes the collaboration format
- **THEN** release checks restore a published baseline, refuse incompatible admission, and export through the compatible release
- **AND** the candidate seeds a new room with preserved presentation content and empty undo history

#### Scenario: Capture a stable release

- **WHEN** a stable package release is published
- **THEN** the post-release workflow captures immutable package integrity, installation locks, and a saved room fixture
- **AND** missing baseline evidence blocks the next stable publication

### Requirement: Concurrent case changes and selection continuity

Case changes SHALL retain character identities. Concurrent case changes SHALL NOT duplicate characters. Each participant's undo SHALL remove only that participant's case choice. Formatting from a peer SHALL preserve caret positions and both endpoints of text selections through run splits.

#### Scenario: Two peers capitalize the same text

- **WHEN** two disconnected peers change the same text to uppercase and synchronize
- **THEN** each original character appears once
- **AND** undoing either case change preserves the other participant's case choice

#### Scenario: A peer formats text around a caret

- **WHEN** formatting from a peer splits the run containing a caret or selection endpoint
- **THEN** the endpoint follows its character into the visible run

#### Scenario: A rescued cell retains vertical merge continuation

- **WHEN** concurrent table operations rescue text into an orphaned cell with `vMerge="1"`
- **THEN** the table layout includes that cell and its rescued text

### Requirement: Selection-aware formatting sections

The Format panel SHALL expose sections for the selected object's supported edits. React and Vue SHALL use shared applicability rules. Hosts SHALL be able to order, omit, and replace sections without replacing the complete panel.

#### Scenario: Format a selected object

- **WHEN** a user selects a shape, picture, connector, table, chart, diagram, or equation
- **THEN** Format shows applicable controls or editing entries
- **AND** selecting another object does not replace an explicitly chosen Comments or Animations tab

#### Scenario: Mixed object types

- **WHEN** several object types are selected
- **THEN** Format does not offer a common property that silently edits only some selected objects

#### Scenario: Customize the toolbar independently

- **WHEN** a host hides or replaces a toolbar control
- **THEN** the inspector retains its independently configured control
- **AND** inspector control rows expose `data-format-slot` for scoped selectors

### Requirement: Partial gradient and effect editing

Detailed gradient, shadow, and reflection edits SHALL preserve source XML for unspecified properties. Existing preset commands SHALL remain available for intentional replacement. Browser commands and server automation SHALL use the same partial operations.

#### Scenario: Change one gradient stop

- **WHEN** a user changes an existing stop's opacity
- **THEN** its source color and unrelated transforms remain unchanged
- **AND** other stops remain unchanged through save and reopen

#### Scenario: Edit an inherited gradient stop

- **WHEN** the first edit copies an inherited gradient into the slide
- **THEN** the inspector retains the selected stop and field focus
- **AND** subsequent edits and undo keep the same stop selected

#### Scenario: Change a custom shadow

- **WHEN** a user changes a shadow's distance
- **THEN** scale, skew, alignment, unrelated effects, and unknown content remain unchanged
- **AND** one undo restores the previous value

### Requirement: Inspector draft targets

Inspector drafts SHALL remain bound to their document, part, object, and property identities. Selection changes SHALL discard stale drafts. Invalid or unsupported edits SHALL provide a reason without changing the document.

#### Scenario: Select another object before committing

- **WHEN** a user enters a draft value and selects another object
- **THEN** the draft does not change the newly selected object

#### Scenario: Change documents before committing

- **WHEN** a command uses a formatting target from a replaced document
- **THEN** the command refuses without writing

### Requirement: Primary inspector tasks

The default inspector SHALL put the selection's primary controls before secondary controls. Host-provided section order SHALL take precedence over default ordering. Section expansion SHALL remain local to the editor instance and selection category.

#### Scenario: Format multiple lines

- **WHEN** compatible lines with different endpoints are selected
- **THEN** the Line section exposes exact width, dash, Start, and End controls
- **AND** differing endpoint values display Mixed
- **AND** changing an endpoint type preserves its dimensions and unknown content

#### Scenario: Display an imported line width

- **WHEN** an imported width has more than two decimal places in points
- **THEN** the inspector displays a rounded value without changing the source
- **AND** mixed-value comparisons use the exact resolved widths
- **AND** the public geometry snapshot retains its visible-only, rounded width

#### Scenario: Use a custom section order

- **WHEN** a host supplies an ordered section list
- **THEN** applicable sections follow that order
- **AND** omitted sections stay hidden
- **AND** expansion changes do not write presentation or collaboration state

### Requirement: Inspector field recovery

Numeric fields SHALL show units and retain invalid drafts with an accessible reason. Enter or blur SHALL commit a valid draft once. Escape SHALL restore the source value. Selection changes SHALL discard stale drafts. Updates SHALL NOT steal field focus.

#### Scenario: Correct an invalid geometry value

- **WHEN** a user submits an invalid geometry value
- **THEN** the field keeps the draft and identifies the error
- **AND** Escape restores the document value without a document edit

#### Scenario: Select a gradient stop without dragging

- **WHEN** a user edits a supported gradient with a keyboard
- **THEN** stop selection exposes Color, Position, and Opacity together
- **AND** stop removal keeps focus in the remaining stop controls

### Requirement: Picture picker targets

Picture replacement and picture fill pickers SHALL capture their document and selection before opening. They SHALL validate that target after file loading and refuse stale results without writing.

#### Scenario: Replace the document while a picker is open

- **WHEN** a new document has matching selected node identifiers
- **AND** an earlier picture picker returns a file
- **THEN** the file does not change the new document

### Requirement: Unused picture export

Saving SHALL omit unused image relationships from parsed presentation slide, layout, master, and notes parts. Saving SHALL omit image parts with no remaining relationships. Unknown relationships and opaque owners SHALL retain their image targets. Unreadable relationships SHALL prevent cleanup. Export cleanup SHALL NOT change the live document or collaboration room.

#### Scenario: Replace a picture and undo after saving

- **WHEN** a participant replaces a picture and saves
- **THEN** the saved file omits the old image unless another relationship uses it
- **AND** undo and pending peer edits can still use the original image in the live document

### Requirement: Visible control boundaries

Default themes SHALL provide at least 3:1 contrast for keyboard focus indicators and enabled form field boundaries against adjacent surfaces. React and Vue SHALL share these styles.

#### Scenario: Focus a dialog field

- **WHEN** a user focuses a text field, number field, select, or description text area with the keyboard
- **THEN** its focus indicator remains visible in light and dark themes
- **AND** invalid fields retain a distinct visible error boundary
