# fidelity-harness Specification

## Purpose

Created by archiving `pptx-editor-foundation` on 29 October 2026.

## Requirements

### Requirement: Measured fidelity

The harness SHALL render every slide with the editor's own pipeline and compare it with the reference PDF on pixels, line breaks, and baselines, using thresholds from one config file.

#### Scenario: No reference PDF

- **WHEN** a deck or fixture has no reference PDF
- **THEN** its slides are reported as "no reference" and listed by `bun run fidelity:missing`

#### Scenario: Save consistency

- **WHEN** a deck is saved and reopened
- **THEN** each slide renders the same as before the save
