# fidelity-harness Specification

## Purpose

Created by archiving `pptx-editor-foundation` on 29 October 2026.

## Requirements

### Requirement: Measured fidelity

The harness SHALL print every slide to PDF with the editor's own pipeline and compare that PDF with the reference PDF on pixels, line breaks, and baselines, with the same tools for both PDFs and thresholds from one config file.

#### Scenario: No reference PDF

- **WHEN** a deck or fixture has no reference PDF
- **THEN** its slides are reported as "no reference", and the fixture gate fails

#### Scenario: Save consistency

- **WHEN** a deck is saved and reopened
- **THEN** each slide paints the same SVG markup as before the save
