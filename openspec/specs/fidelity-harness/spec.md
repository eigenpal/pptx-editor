# fidelity-harness Specification

## Purpose

Created by archiving `pptx-editor-foundation` on 29 October 2026.

## Requirements

### Requirement: Measured fidelity

The harness SHALL print every slide to PDF with the editor's own pipeline and compare that PDF with the reference PDF on pixels, line breaks, extra text, and baselines, with the same tools for both PDFs and thresholds from one config file.

#### Scenario: No reference PDF

- **WHEN** a deck or fixture has no reference PDF
- **THEN** its slides are reported as "no reference", and the fixture gate fails

#### Scenario: Save consistency

- **WHEN** a deck is saved and reopened
- **THEN** each slide paints the same SVG markup as before the save

#### Scenario: Extra text

- **WHEN** our page holds a text line that the reference page does not hold
- **THEN** the slide fails on extra text

#### Scenario: Numbers and values

- **WHEN** a text line holds two or more letters or digits, such as a number, a date, or a chart or table value
- **THEN** the line check pairs it with the line of the same text in the other PDF, by the same rules for both PDFs, and a line of one letter or digit is checked by pixels only

#### Scenario: Spaces in a text item

- **WHEN** one PDF writes a space between two glyphs of a line and the other PDF does not
- **THEN** both lines get the same line check, because the rule reads the text without whitespace

#### Scenario: A font that the harness did not load

- **WHEN** our PDF draws with a font that is not a fixture font, an engine font, or a harness fallback face
- **THEN** the fixture gate fails and names the font

#### Scenario: A reference font that the deck does not name

- **WHEN** a fixture's reference PDF embeds a font that is not an open font that its deck names
- **THEN** the fixture gate fails and names the font
