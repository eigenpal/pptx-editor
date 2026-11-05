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
