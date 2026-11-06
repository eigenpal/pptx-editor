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

#### Scenario: A font that the harness did not load

- **WHEN** our PDF draws with a font that is not a fixture font, an engine font, or a harness fallback face
- **THEN** the fixture gate fails and names the font

#### Scenario: A reference font that the deck does not name

- **WHEN** a fixture's reference PDF embeds a font that is not an open font that its deck names, and the fixture is not in `NON_OPEN_FONT_DECKS`
- **THEN** the fixture gate fails and names the font
