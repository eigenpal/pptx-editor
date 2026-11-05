# table-rendering Specification

## Purpose

Created by archiving `table-rendering` on 29 October 2026.

## Requirements

### Requirement: Tables render

The editor SHALL render `a:tbl` graphic frames with their grid, merged cells, fills, borders, and text.

#### Scenario: A table with merged cells

- **WHEN** a slide has a table with a cell that spans two columns and a cell that spans two rows
- **THEN** each merged cell draws once across its span, and the covered cells draw nothing

### Requirement: Table styles

The editor SHALL apply the table style from `ppt/tableStyles.xml` to each cell, in the part order of the design, controlled by the `a:tblPr` flags.

#### Scenario: Header row and banded rows

- **WHEN** a table has `firstRow="1"` and `bandRow="1"` and a style with `firstRow`, `band1H`, and `wholeTbl` parts
- **THEN** the first row uses the `firstRow` fill and text, the next row uses `band1H`, and the row after it uses `wholeTbl`

### Requirement: Rows fit their text

A row SHALL be at least as tall as `a:tr/@h` and as tall as the text of its cells needs.

#### Scenario: Long text

- **WHEN** a cell's text needs more height than its row's `h`
- **THEN** the row grows, and the rows below it move down

### Requirement: Shape style font reference

A shape style's `a:fontRef` SHALL take precedence over `defaultTextStyle` and the master text styles for text color and font.

#### Scenario: Inserted shape

- **WHEN** a shape has `a:fontRef idx="minor"` with `lt1` and the presentation's default text color is `tx1`
- **THEN** its text renders in `lt1`
