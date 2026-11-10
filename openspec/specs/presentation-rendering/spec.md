# presentation-rendering Specification

## Purpose

Created by archiving `pptx-editor-foundation`, `default-font-coverage`, `effects-fidelity`, `fill-and-line-fidelity` on 29 October 2026.

## Requirements

### Requirement: Bounded, lossless open and save

The engine SHALL read a `.pptx` through bounded zip and XML readers and SHALL save it without losing any content the model carries.

#### Scenario: Round trip

- **WHEN** a deck is opened and saved without edits
- **THEN** the reopened package has the same canonical fingerprint, and every binary part is byte-identical

#### Scenario: Hostile package

- **WHEN** a package contains a DTD, a traversal path, or exceeds a size or ratio cap
- **THEN** the open is refused with a typed reason and nothing is rendered

### Requirement: Placeholder and style inheritance

The cascade SHALL resolve placeholders, color maps, style references, theme fonts, and list styles by the PresentationML inheritance chain: master, layout, then slide.

#### Scenario: Placeholder text style

- **WHEN** a slide title placeholder has no run properties
- **THEN** its size, font, and color come from the layout and master placeholders and the master title style

### Requirement: Spec-driven geometry

Preset shapes SHALL be drawn from the ECMA-376 guide formulas, and custom geometry SHALL use the same evaluator.

#### Scenario: Adjusted rounded rectangle

- **WHEN** a `roundRect` has `adj` set
- **THEN** the corner radius follows the preset's guide formulas

### Requirement: Embedded fonts

Embedded fonts SHALL be decoded in memory for the open document. Editing sessions SHALL use only embedded faces with installable or editable embedding permission for measurement and paint. They SHALL refuse restricted, preview-and-print-only, and bitmap-only faces without changing the original package parts.

#### Scenario: MTX-compressed EOT

- **WHEN** a deck embeds MicroType Express compressed fonts with installable or editable embedding permission
- **THEN** text measures and paints with the decoded faces, and no font bytes are written to disk

#### Scenario: Preview-and-print-only face

- **WHEN** a deck embeds a face with preview-and-print-only permission
- **THEN** the session reports the refused face and uses an available host face or substitute
- **AND** editing, save, and reopen preserve the original embedded font bytes and relationship
- **AND** viewer mode uses the same admission rule, so changing modes cannot enable the refused face

### Requirement: Host-provided fonts

The engine SHALL ask a host font provider for every family the deck names that no host or embedded face satisfies, and SHALL use the returned faces for measurement and paint before any substitute.

#### Scenario: Licensed face supplied

- **WHEN** a deck uses Aptos and the host provider returns Aptos faces
- **THEN** text measures and paints with Aptos and Aptos is not reported as substituted

### Requirement: Fonts added after open

The editor SHALL accept extra faces after a deck is open, re-measure and repaint, and keep the document and its undo history unchanged.

#### Scenario: Local fonts granted

- **WHEN** the user grants local font access and the substituted families are added
- **THEN** the slides re-layout with those faces and undo still restores earlier edits

### Requirement: Substitution is visible

The editor SHALL report which named families were substituted, and the React adapter SHALL show a notice listing them.

#### Scenario: Missing family

- **WHEN** a deck names a family that no source provides
- **THEN** the notice lists that family and the substitute is used

### Requirement: Upright East Asian vertical text

In `eaVert` text bodies, East Asian characters SHALL be drawn upright while the line runs top to bottom.

#### Scenario: Japanese title in eaVert

- **WHEN** a text body with `vert="eaVert"` contains Japanese text
- **THEN** each ideograph is drawn upright, stacked top to bottom

### Requirement: Outer shadows

The renderer SHALL draw `a:outerShdw` as the shape's content scaled by `sx`/`sy` and skewed by `kx`/`ky` about the `algn` point, offset by `dist` at `dir`, blurred by `blurRad`, in the shape's frame when `rotWithShape` is true and in the slide's frame otherwise.

#### Scenario: Perspective shadow

- **WHEN** a shape has `<a:outerShdw sy="23000" kx="800400" algn="br" …/>`
- **THEN** the shadow is 23 % of the shape's height, anchored at its bottom-right corner, and leans by 13.34°

#### Scenario: Text box shadow

- **WHEN** a shape without fill has text and an outer shadow
- **THEN** the text casts the shadow

### Requirement: Reflection

The renderer SHALL draw `a:reflection` as a transformed copy of the shape's content that fades from `stA` at `stPos` to `endA` at `endPos` along `fadeDir`.

#### Scenario: Tight reflection

- **WHEN** a shape has `<a:reflection blurRad="6350" stA="50000" endA="300" endPos="55000" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>`
- **THEN** a mirrored copy appears directly below the shape, at half opacity at its top, fading out at 55 % of its height

### Requirement: Fills

The renderer SHALL draw rectangular path gradients as concentric rectangles, preset patterns as bitmaps in their two colors, gradient text across the text's bounds, and tiled pictures at the image's natural size.

#### Scenario: Rectangular gradient

- **WHEN** a shape has `<a:path path="rect"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>`
- **THEN** the first stop is at the center, the last stop is on the box edge, and the level lines are rectangles

#### Scenario: Tiled picture

- **WHEN** a 128×96-pixel image without a resolution is tiled with `sx="50000" sy="50000"`
- **THEN** each tile is 64×48 pt, because an image without a resolution reads at 72 DPI

### Requirement: Compound lines

The renderer SHALL draw `dbl`, `thickThin`, `thinThick`, and `tri` lines as separate bands across the line width.

#### Scenario: Double line

- **WHEN** a rectangle has `<a:ln w="38100" cmpd="dbl">`
- **THEN** two lines of 1/3 of the width draw with a gap of 1/3 of the width between them

### Requirement: Line ends and dashes

The editor and the PDF export SHALL draw dashes and line ends from the same geometry (`layout/line-strokes.ts`), as the `line-ends-dashes` reference PDF shows them. A line end SHALL scale with the line width, but never with less than 2 pt.

#### Scenario: Triangle on a thin line

- **WHEN** an open path has `<a:ln w="9525"><a:tailEnd type="triangle" w="sm" len="lg"/></a:ln>`
- **THEN** the triangle is 10 pt long and 4 pt wide with its tip on the end point, and the line stops 9 pt before the end point

#### Scenario: Square caps on a preset dash

- **WHEN** a line has `cap="sq"` and `<a:prstDash val="dot"/>`
- **THEN** each dot is one line width long, and only the ends of the line get a square cap

### Requirement: Equations

The layout SHALL draw an equation (`a14:m`) whose constructs it reads, within its depth and size limits, with the rules of TeX, in the editor canvas and in the PDF export from one layout. It SHALL draw the fallback of the text box for any other equation.

#### Scenario: A supported equation

- **WHEN** a slide holds a fraction in `mc:AlternateContent` with a fallback picture
- **THEN** the numerator, the rule, and the denominator draw, and the fallback picture does not

#### Scenario: An unsupported construct

- **WHEN** an equation holds a construct that the layout does not read
- **THEN** the `mc:Fallback` branch draws

### Requirement: Contextual RTL PDF text

PDF export SHALL use the layout engine's contextual RTL glyphs and mark offsets. Text effects SHALL use the same glyphs. The text layer SHALL retain searchable words and spaces without duplicating combining marks.

#### Scenario: Arabic fallback

- **WHEN** Arabic text names an unavailable font and the host enables Google Fonts
- **THEN** an available Amiri face supplies its glyphs without changing the saved presentation
- **AND** PDF letters join and extracted words retain their spaces

### Requirement: Explicit network fonts

PDF export SHALL make Google Fonts requests only when `googleFonts` is enabled. It SHALL resolve names through the pinned catalog, verify font hashes, and preserve supplied and permitted embedded faces before network fallback faces.

#### Scenario: Offline default

- **WHEN** an export omits `googleFonts`
- **THEN** font resolution makes no Google Fonts requests

#### Scenario: Invalid network bytes

- **WHEN** a returned face fails its pinned hash or structural check
- **THEN** the exporter does not shape or embed those bytes
