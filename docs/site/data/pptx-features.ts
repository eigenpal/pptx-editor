/**
 * PPTX feature support matrix: single source of truth.
 *
 * Adapted from DOCX Editor `docs/site/data/word-features.ts` at commit 29535a640510 (the type
 * block, `FEATURE_TIERS`, and the by-id lookup). The row type has the same fields as
 * `WordFeature`; only the categories differ. Add PPTX categories, never new fields or status
 * values.
 *
 * The docs site renders this file on `/docs/pptx-fidelity` through the `<FeatureMatrix>`,
 * `<FeatureSummary>`, and `<FeatureBadge>` components, as for DOCX Editor. The site syncs this
 * file as it is, so it imports nothing. `docs/site/README.md` says what the site provides. The
 * `tier` field drives feature access and pricing pages; the components do not show it.
 *
 * The evidence for each row (its fixtures and its known approximations) is in
 * `pptx-feature-evidence.ts`, keyed by row `id`. Add an entry there for each new row.
 *
 * Status axes:
 * - editing:   can the user (or code driving the editor) change it in the editor?
 * - rendering: does it display like the reference PDF of the same slide?
 * - roundTrip: does it survive open -> edit -> save -> reopen without loss?
 *
 * Honesty rule: when in doubt, downgrade. A "partial" that turns out to be full delights; a
 * "full" that turns out to be partial burns trust.
 *
 * Evidence rule (checked by pptx-features.test.ts):
 * - `rendering: 'full'` needs a current reference PDF next to every fixture of the row
 *   (`fixtures/<name>.pdf` with its identity in `fixtures/references.json`, see CONTRIBUTING.md
 *   "Fidelity harness"), and no slide of those fixtures in `FIXTURE_KNOWN_GAPS`: the fixture
 *   passes the harness in every gated environment. The row also owns no approximation but the
 *   three page-wide text ones (`PAGE_WIDE_APPROXIMATIONS` in pptx-features.test.ts), and its
 *   notes name no rendering gap, such as a substitute font or a part that does not draw. Until
 *   then, `partial` is the most a row can claim, however good the render looks.
 * - `roundTrip: 'full'` needs at least one fixture, and its saved copies opened with no repair
 *   prompt (`SAVE_VERIFIED_FIXTURES` in `save-verified-fixtures.ts`). That is necessary, not
 *   enough: the row's own edits must also survive save and reopen.
 * - `roundTrip: 'full'` and `roundTrip: 'preserved'` need a recipe in `FEATURE_EDITS`
 *   (`scripts/fidelity/feature-edits.ts`) on a fixture of the row, whose copy opened with no
 *   repair prompt (`SAVE_VERIFIED_FEATURE_EDITS`). A `preserved` recipe edits next to the kept
 *   content and checks that the content is as it was.
 *
 * Notes rule: notes render inside a table cell, so use at most three sentences. Link to a guide for details. Write Simplified Technical
 * English: active voice, one idea per sentence, 20 words or fewer per sentence. Name the
 * observable behavior, not the internal module or change proposal. The notes of a `premium` row
 * name the EigenPal Pro License, as in DOCX Editor.
 */

export type FeatureStatus =
  | 'full'
  | 'partial'
  | 'render-only'
  | 'preserved' // round-trips losslessly as inert content; editing/rendering may be absent
  | 'planned'
  | 'none';

/** The tiers, as values, so the test can check every row at runtime. Same values as DOCX. */
export const FEATURE_TIERS = ['community', 'premium'] as const;

export type FeatureTier = (typeof FEATURE_TIERS)[number];

export type FeatureCategory =
  | 'text'
  | 'paragraphs'
  | 'lists'
  | 'shapes'
  | 'fills-lines'
  | 'effects'
  | 'pictures'
  | 'tables'
  | 'objects'
  | 'slides'
  | 'themes'
  | 'presenting'
  | 'review'
  | 'collaboration'
  | 'export';

export interface PptxFeature {
  /** Stable key, e.g. 'shapes.preset-geometry'. Never rename; gating may reference it. */
  id: string;
  name: string;
  category: FeatureCategory;
  editing: FeatureStatus;
  rendering: FeatureStatus;
  roundTrip: FeatureStatus;
  tier: FeatureTier;
  notes?: string;
  /** Docs page that covers the feature, e.g. '/docs/guides/tables'. */
  docsLink?: string;
}

export const FEATURE_CATEGORY_LABELS: Record<FeatureCategory, string> = {
  text: 'Text & formatting',
  paragraphs: 'Paragraphs & text boxes',
  lists: 'Bullets & numbering',
  shapes: 'Shapes, groups & connectors',
  'fills-lines': 'Fills & lines',
  effects: 'Effects',
  pictures: 'Pictures',
  tables: 'Tables',
  objects: 'Charts, diagrams, media & OLE',
  slides: 'Slides, layouts & notes',
  themes: 'Themes, masters & inheritance',
  presenting: 'Presenting',
  review: 'Comments & review',
  collaboration: 'Collaboration & automation',
  export: 'Export & print',
};

export const pptxFeatures: PptxFeature[] = [
  // --- Text & formatting -----------------------------------------------
  {
    id: 'text.basic-formatting',
    name: 'Bold, italic, underline, strikethrough',
    category: 'text',
    editing: 'full',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Apply bold, italic, underline, and strikethrough, including underline colors. Every underline style renders, with its underline color and outline.',
  },
  {
    id: 'text.font-family-size',
    name: 'Font family and size',
    category: 'text',
    editing: 'full',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Choose theme fonts, embedded fonts, or fonts supplied by your application. Missing fonts use substitutes, and embedding permissions control which embedded faces load. Font substitutions can affect fidelity.',
    docsLink: '/docs/customization/fonts',
  },
  {
    id: 'text.color-highlight',
    name: 'Text color and highlight',
    category: 'text',
    editing: 'full',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Gradient and pattern text fills render. Their geometry is approximate. The Font dialog sets theme font colors.',
  },
  {
    id: 'text.caps-baseline',
    name: 'Small caps, all caps, superscript, subscript',
    category: 'text',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Superscript and subscript toggles. The Font dialog sets the offset, caps, and strikethrough kinds. Superscript and subscript offsets are approximate.',
  },
  {
    id: 'text.character-spacing',
    name: 'Character spacing and kerning',
    category: 'text',
    editing: 'partial',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'The Character Spacing menu has five presets, from Very Tight to Very Loose. The Font dialog sets any spacing and the kerning size.',
  },
  {
    id: 'text.change-case',
    name: 'Change case',
    category: 'text',
    editing: 'partial',
    rendering: 'none',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Change case while keeping formatting, or cycle through three modes with Shift+F3. Letters whose case conversion changes their length stay unchanged. Conversion does not use the run language.',
  },
  {
    id: 'text.hyperlinks',
    name: 'Hyperlinks',
    category: 'text',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Link text to websites, email addresses, slides, custom shows, or files. The slide show follows slide links but does not open custom show or file links. Program links never run, and link colors are approximate.',
    docsLink: '/docs/recipes/slide-email-links',
  },
  {
    id: 'text.input-method',
    name: 'Typing and input-method composition',
    category: 'text',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Type and compose text at the caret, with selection and deletion by grapheme cluster. Composed text shows in place before it commits as one undo step, but conflicting peer edits can cancel unfinished text. Validate your target input methods.',
  },
  {
    id: 'text.complex-scripts',
    name: 'Right-to-left and complex scripts',
    category: 'text',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Edit East Asian, right-to-left, and mixed-direction text with script-specific fonts and navigation. Line layout remains approximate. Arabic letters join in the editor and PDF export when a suitable font is available.',
    docsLink: '/docs/guides/east-asian-text',
  },
  {
    id: 'text.find-replace',
    name: 'Find and replace',
    category: 'text',
    editing: 'full',
    rendering: 'none',
    roundTrip: 'full',
    tier: 'community',
    notes: 'Searches slide text and speaker notes.',
  },
  {
    id: 'text.clipboard',
    name: 'Copy and paste',
    category: 'text',
    editing: 'full',
    rendering: 'none',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Copy text, shapes, and slides between editors, or paste formatted HTML from other applications. In browsers, shapes from other applications paste as pictures. Desktop hosts can provide editable shapes through system clipboard access.',
    docsLink: '/docs/guides/copy-and-paste',
  },
  {
    id: 'text.format-painter',
    name: 'Format painter',
    category: 'text',
    editing: 'partial',
    rendering: 'none',
    roundTrip: 'full',
    tier: 'community',
  },

  // --- Paragraphs & text boxes -----------------------------------------
  {
    id: 'paragraphs.alignment',
    name: 'Horizontal and vertical alignment',
    category: 'paragraphs',
    editing: 'full',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
  },
  {
    id: 'paragraphs.spacing',
    name: 'Line and paragraph spacing',
    category: 'paragraphs',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Set line spacing and paragraph spacing in the Paragraph dialog. Line heights, baselines, and glyph advances are approximate.',
  },
  {
    id: 'paragraphs.indent-tabs',
    name: 'Indents and tab stops',
    category: 'paragraphs',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Set first-line and hanging indents, or change list levels. Existing tab stops align text, but you cannot edit them. The editor has no ruler.',
  },
  {
    id: 'paragraphs.autofit',
    name: 'Autofit (shrink text, resize shape)',
    category: 'paragraphs',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Choose no autofit, shrink text on overflow, or resize the shape to fit. Autofit recalculates after edits, with approximate results for rotated text, multiple columns, and nonrectangular shapes. Titles shrink to 90% at most, then overflow.',
  },
  {
    id: 'paragraphs.text-direction-columns',
    name: 'Vertical text and columns',
    category: 'paragraphs',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Set horizontal, rotated, stacked, or upright East Asian text, with up to 16 columns. Upright East Asian text does not use vertical font metrics. Mongolian and right-to-left stacked text render approximately and cannot be selected through editing commands.',
  },
  {
    id: 'paragraphs.text-warp',
    name: 'Text warp',
    category: 'paragraphs',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Apply preset text warps, with adjustment values available through commands. Warped glyphs lose ligatures and complex-script forms, while the caret and selection follow flat text. Vertical, rotated, and diagram text stay flat.',
    docsLink: '/docs/guides/text-warp',
  },

  // --- Bullets & numbering ---------------------------------------------
  {
    id: 'lists.bullets-numbering',
    name: 'Bullets and numbering',
    category: 'lists',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Set bullet characters, pictures, numbering schemes, colors, sizes, and starting numbers. Keyboard controls continue lists and change levels. Symbol substitutions, picture bullet sizing, and some numbering schemes are approximate.',
  },

  // --- Shapes, groups & connectors -------------------------------------
  {
    id: 'shapes.preset-geometry',
    name: 'Preset shapes',
    category: 'shapes',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'All 187 presets use the spec guide formulas. Insert, move, resize, rotate, and drag the adjust handles (`setAdjustments`). A shape cannot change to another preset.',
    docsLink: '/docs/recipes/shape-adjustments',
  },
  {
    id: 'shapes.custom-geometry',
    name: 'Freeform shapes',
    category: 'shapes',
    editing: 'none',
    rendering: 'partial',
    roundTrip: 'preserved',
    tier: 'community',
    notes: 'No edit points and no freehand tools.',
  },
  {
    id: 'shapes.groups',
    name: 'Groups',
    category: 'shapes',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Group, ungroup, and edit inside a group. Group effects draw over the whole group as one picture and are approximate.',
  },
  {
    id: 'shapes.connectors',
    name: 'Lines and connectors',
    category: 'shapes',
    editing: 'partial',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Attach connector ends to shapes so they follow moves and resizes. Drag elbow or curved connector segments to adjust their route. Moving a connector independently or deleting its attached shape releases the connection.',
  },
  {
    id: 'shapes.arrange',
    name: 'Align, distribute, order, and nudge',
    category: 'shapes',
    editing: 'partial',
    rendering: 'none',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Align, order, and nudge objects, or distribute three or more objects. Alignment and distribution use unrotated boxes. Smart guides appear during dragging, but there is no Align to Slide setting.',
  },
  {
    id: 'shapes.flip-rotate',
    name: 'Flip, rotate, exact size, and aspect lock',
    category: 'shapes',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Flip, rotate, or set exact object dimensions from the controls. Aspect locks preserve proportions when you resize one object. Resizing several objects ignores individual aspect locks.',
  },
  {
    id: 'shapes.selection-pane',
    name: 'Selection pane',
    category: 'shapes',
    editing: 'partial',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Select, rename, hide, show, and reorder objects in the Selection pane, including objects inside groups. Some edits to markup-compatibility content remove alternate copies. See the editing approximations below.',
    docsLink: '/docs/customization/layout#open-the-selection-pane',
  },
  {
    id: 'shapes.styles-gallery',
    name: 'Shape styles gallery',
    category: 'shapes',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes: 'Apply six theme style rows in seven colors. Transparent preset styles are unavailable.',
  },

  // --- Fills & lines ---------------------------------------------------
  {
    id: 'fills-lines.solid-gradient',
    name: 'Solid and gradient fills',
    category: 'fills-lines',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Apply solid colors or gradients. Edit gradient stops, opacity, and direction without replacing unrelated properties. Gradient colors, focus, tiling, and rotation can differ from the reference. Rectangle and shape gradients can show seams or bands.',
  },
  {
    id: 'fills-lines.pattern-picture',
    name: 'Pattern and picture fills',
    category: 'fills-lines',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Apply preset patterns or stretched picture fills, with tiling available through commands. Picture fill crop, offset, scale, and transparency are not editable. Tile sizing and placement are approximate in some contexts.',
    docsLink: '/docs/recipes/pattern-picture-fills',
  },
  {
    id: 'fills-lines.lines',
    name: 'Line color, width, dash, and arrowheads',
    category: 'fills-lines',
    editing: 'full',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Set line colors, widths, dashes, and arrowheads. Compound lines and dashes on curves are approximate. Excessively complex dash patterns display as solid lines but remain unchanged on save.',
  },
  {
    id: 'fills-lines.theme-colors',
    name: 'Theme colors and color transforms',
    category: 'fills-lines',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes: 'The color picker shows the theme palette. Transform color spaces are approximate.',
  },

  // --- Effects ---------------------------------------------------------
  {
    id: 'effects.outer-shadow',
    name: 'Outer shadow',
    category: 'effects',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Apply or clear shadow presets on supported shapes and pictures. Edit existing shadow color, opacity, blur, distance, and direction. Groups, tables, charts, and theme effect graphs reject these edits. A group shadow draws as the silhouette of the whole group. Perspective shadows are unavailable, and blur is measured only for small radii.',
  },
  {
    id: 'effects.glow-reflection',
    name: 'Glow, inner shadow, soft edge, and reflection',
    category: 'effects',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'premium',
    notes:
      'Display and edit these effects with @pptx-editor.dev/pro under the EigenPal Pro License. Without the drawing module, the editor omits these effects but preserves them on save. Blur remains approximate, especially for reflections.',
    docsLink: '/docs/pro/drawing',
  },
  {
    id: 'effects.fill-overlay',
    name: 'Fill overlay',
    category: 'effects',
    editing: 'none',
    rendering: 'partial',
    roundTrip: 'preserved',
    tier: 'community',
    notes:
      'The overlay blends with the shape fill only. The multiply, screen, and darken modes are measured, and so is the opaque result over a transparent fill. Other cases are approximate.',
  },
  {
    id: 'effects.3d-and-other',
    name: '3-D, preset shadows, blur, and text effects',
    category: 'effects',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'premium',
    notes:
      'Display approximate 3-D shapes and blur with @pptx-editor.dev/pro under the EigenPal Pro License. Preset shadows display without the module, but bottom bevels and text effects do not render. PDF export draws blur approximately, with text inside the shape kept sharp.',
    docsLink: '/docs/guides/3d-effects',
  },

  // --- Pictures --------------------------------------------------------
  {
    id: 'pictures.raster-svg',
    name: 'PNG, JPEG, GIF, and SVG pictures',
    category: 'pictures',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Insert and transform pictures, or replace them with PNG, JPEG, or GIF files. Simple SVG pictures render as vectors; other SVG pictures use their PNG fallback. Large raster pictures can display at reduced resolution without changing stored images.',
    docsLink: '/docs/guides/images',
  },
  {
    id: 'pictures.adjustments',
    name: 'Corrections, color, and transparency',
    category: 'pictures',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Apply correction, recolor, and transparency presets, or reset a picture. Saturation, color tone, and artistic effects are unavailable. Rendering is approximate, and some existing picture effects prevent corrections or reset.',
  },
  {
    id: 'pictures.crop',
    name: 'Crop, crop to shape, fill, and fit',
    category: 'pictures',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Crop with handles, preset shapes, aspect ratios, Fill, or Fit. Crop commands respect picture locks, but canvas handles can still resize and rotate locked pictures. Tiled or inset fills reject cropping, and JPEG orientation metadata is ignored.',
    docsLink: '/docs/guides/images',
  },
  {
    id: 'pictures.styles-border',
    name: 'Picture styles and picture border',
    category: 'pictures',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Apply five picture styles and set border color, weight, and dashes. Style and border fidelity remains approximate. Round border dots display as squares.',
  },
  {
    id: 'pictures.alt-text',
    name: 'Alt text and decorative objects',
    category: 'pictures',
    editing: 'partial',
    rendering: 'none',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Set a description, title, or decorative status for a selected object. Screen readers announce picture descriptions and skip decorative pictures. The editor does not generate descriptions.',
  },
  {
    id: 'pictures.emf-wmf',
    name: 'EMF and WMF pictures',
    category: 'pictures',
    editing: 'none',
    rendering: 'partial',
    roundTrip: 'preserved',
    tier: 'community',
    notes:
      'The editor, image export, and PDF export draw each EMF and WMF picture as a PNG. The placeholder shows until the PNG is ready. The picture file stays unchanged.',
  },

  // --- Tables ----------------------------------------------------------
  {
    id: 'tables.render',
    name: 'Tables and table styles',
    category: 'tables',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Tables render with merges, borders, and table styles. A cell with a merge marker that no span reaches draws as a normal cell (an approximation). The rows below list what you can edit.',
    docsLink: '/docs/guides/tables',
  },
  {
    id: 'tables.cell-text',
    name: 'Text in table cells',
    category: 'tables',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Edit text and formatting in cells, cell ranges, or the whole table. Tab moves between cells and adds a row after the last cell. Paste tab-separated text into a range, expanding the table when needed.',
    docsLink: '/docs/guides/tables',
  },
  {
    id: 'tables.structure',
    name: 'Insert tables, rows, and columns',
    category: 'tables',
    editing: 'partial',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Insert tables and add or delete rows and columns. New tables fill an available content or table placeholder. Deleting a table does not restore its placeholder.',
    docsLink: '/docs/guides/tables',
  },
  {
    id: 'tables.merge-split',
    name: 'Merge and split cells',
    category: 'tables',
    editing: 'partial',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Merge cells while keeping their text, or split one cell at a time. Merged cells split only along even grid lines. Inconsistent merge markup follows the approximation rules below.',
    docsLink: '/docs/guides/tables',
  },
  {
    id: 'tables.styles',
    name: 'Table styles and style options',
    category: 'tables',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      "Toggle header, total, banding, and first or last column options. The gallery contains the presentation's styles and the default style only.",
    docsLink: '/docs/guides/tables',
  },
  {
    id: 'tables.cell-format',
    name: 'Cell shading, borders, margins, and alignment',
    category: 'tables',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Set cell shading and border color, dash, and weight. Diagonal borders and cell effects are not editable.',
    docsLink: '/docs/guides/tables',
  },
  {
    id: 'tables.sizes',
    name: 'Row height and column width',
    category: 'tables',
    editing: 'partial',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Drag grid lines, enter dimensions, distribute rows and columns, or resize the table frame. Row heights are minimums, so rows grow to fit text.',
    docsLink: '/docs/guides/tables',
  },

  // --- Charts, diagrams, media & OLE -----------------------------------
  {
    id: 'objects.charts',
    name: 'Charts',
    category: 'objects',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'premium',
    notes:
      'Insert and edit common chart types, data, and formatting with @pptx-editor.dev/pro under the EigenPal Pro License. Unsupported chart types or features show placeholders or remain read-only. Linked workbooks never open automatically.',
    docsLink: '/docs/guides/charts',
  },
  {
    id: 'objects.diagrams',
    name: 'Diagrams',
    category: 'objects',
    editing: 'partial',
    rendering: 'partial',
    notes:
      'Display diagrams from their saved drawings and edit supported node text. Add or remove nodes in simple, single-line diagrams with one child level. Missing drawings show placeholders, and formatting and layout changes are unavailable.',
    roundTrip: 'full',
    tier: 'community',
    docsLink: '/docs/guides/diagrams',
  },
  {
    id: 'objects.media-ole',
    name: 'Audio, video, and OLE objects',
    category: 'objects',
    editing: 'none',
    rendering: 'render-only',
    roundTrip: 'preserved',
    tier: 'community',
    notes:
      'Audio, video, and OLE objects show a stored poster picture or a labeled placeholder. They remain in the file but never play or run.',
  },
  {
    id: 'objects.equations',
    name: 'Equations',
    category: 'objects',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Insert and edit equations through a linear text form. Unsupported equations remain unchanged and read-only. Equations use STIX Two Math, with approximate spacing and positioning.',
    docsLink: '/docs/guides/equations',
  },

  // --- Slides, layouts & notes -----------------------------------------
  {
    id: 'slides.structure',
    name: 'Add, delete, duplicate, and reorder slides',
    category: 'slides',
    editing: 'full',
    rendering: 'none',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'New slide takes a layout from a gallery that the engine draws. A delete shows what went, with an Undo button.',
  },
  {
    id: 'slides.hidden',
    name: 'Hidden slides',
    category: 'slides',
    editing: 'partial',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Hide slide works from the rail and the menus, and the slide show skips hidden slides. The PDF export leaves hidden slides out.',
  },
  {
    id: 'slides.placeholder-prompts',
    name: 'Placeholder prompt text',
    category: 'slides',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Empty placeholders show prompts and dotted outlines only in the editor. Insert tables or pictures through placeholder buttons or the Insert menu. The inserted object replaces the placeholder.',
    docsLink: '/docs/guides/images#fill-an-empty-placeholder',
  },
  {
    id: 'slides.notes',
    name: 'Speaker notes',
    category: 'slides',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Edit speaker notes and apply bold, italic, underline, or strikethrough. Other formatting stays in the file but is not editable. Formatting at a caret outside a word does not carry into new text.',
  },
  {
    id: 'slides.header-footer',
    name: 'Slide numbers, dates, and footers',
    category: 'slides',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Show or hide slide numbers, dates, and footers through layout or master placeholders. Apply changes to selected slides or all slides, optionally excluding title slides. New date fields use the datetime1 format only.',
  },
  {
    id: 'slides.size',
    name: 'Slide size',
    category: 'slides',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Choose a preset or custom slide size from 1 to 56 inches in either orientation. Content scales to fit by default, but some effects, diagram text, and default values keep their sizes. Portrait and very wide slide fidelity is unmeasured, and PDF page dimensions round to whole points.',
  },
  {
    id: 'slides.sections',
    name: 'Sections',
    category: 'slides',
    editing: 'partial',
    rendering: 'full',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Add, rename, remove, move, and collapse sections through section headers or custom toolbar controls. The default editor has no Section menu. Sections cannot be dragged.',
  },
  {
    id: 'slides.sorter',
    name: 'Slide sorter view',
    category: 'slides',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes: 'Select, move, duplicate, delete, and hide slides by section. No zoom.',
  },

  // --- Themes, masters & inheritance -----------------------------------
  {
    id: 'themes.placeholder-inheritance',
    name: 'Placeholder inheritance from layout and master',
    category: 'themes',
    editing: 'none',
    rendering: 'partial',
    roundTrip: 'preserved',
    tier: 'community',
    notes: 'Content placeholders are approximate.',
  },
  {
    id: 'themes.backgrounds',
    name: 'Slide backgrounds',
    category: 'themes',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'The Background dialog sets a color, gradient, pattern, or stretched or tiled picture. OK applies it as one undo step.',
  },
  {
    id: 'themes.change-layout-theme',
    name: 'Change layout, apply a theme, edit the master',
    category: 'themes',
    editing: 'partial',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes:
      'Change or reset slide layouts, and edit shapes, placeholders, and backgrounds in the slide master view. Placeholder matching and placement have known approximations. You cannot add, delete, or switch layouts, masters, or themes through the master view.',
    docsLink: '/docs/guides/slide-master',
  },

  // --- Presenting ------------------------------------------------------
  {
    id: 'presenting.slideshow',
    name: 'Slide show',
    category: 'presenting',
    editing: 'none',
    rendering: 'partial',
    roundTrip: 'none',
    tier: 'community',
    notes:
      'Keyboard navigation only. No presenter view. The show writes nothing to the file, so round trip does not apply.',
  },
  {
    id: 'presenting.animations-transitions',
    name: 'Animations and transitions',
    category: 'presenting',
    editing: 'partial',
    rendering: 'none',
    roundTrip: 'full',
    tier: 'community',
    docsLink: '/docs/guides/animations',
    notes:
      'Configure 13 animation presets and five transition types, including timing and order. Motion paths and interactive sequences cannot be added, and markup-compatibility timing remains read-only. The editor preserves effects but does not play them.',
  },

  // --- Comments & review -----------------------------------------------
  {
    id: 'review.comments',
    name: 'Comments',
    category: 'review',
    editing: 'partial',
    rendering: 'none',
    roundTrip: 'full',
    tier: 'premium',
    notes:
      'Add, edit, reply to, and resolve comments with @pptx-editor.dev/pro under the EigenPal Pro License. Legacy comments cannot resolve, and unsupported thread content remains preserved but read-only. Concurrent edits can shift text anchors, and importing slides with threaded comments is refused.',
    docsLink: '/docs/pro/comments',
  },

  // --- Collaboration & automation --------------------------------------
  {
    id: 'collaboration.realtime',
    name: 'Real-time collaboration',
    category: 'collaboration',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'full',
    tier: 'premium',
    notes:
      'Synchronize edits, typing, and peer selections through Yjs with @pptx-editor.dev/pro under the EigenPal Pro License. Undo applies to your own changes, with limits when peers edit the same content. Concurrent case changes preserve character identities. Peer formatting preserves carets and selections. Rescued text remains visible in orphaned vertical merge cells. Concurrent undo can reorder text; see the collaboration guide.',
    docsLink: '/docs/pro/collaboration',
  },
  {
    id: 'collaboration.automation-api',
    name: 'Office.js-shaped automation API',
    category: 'collaboration',
    editing: 'partial',
    rendering: 'none',
    roundTrip: 'full',
    tier: 'premium',
    notes:
      'Automate presentation edits on servers or in live editors with @pptx-editor.dev/editor-api under the EigenPal Pro License. The API implements a selected Office.js subset with batched operations. Server image rendering requires a rasterizer supplied by your application.',
    docsLink: '/docs/editor-api',
  },

  // --- Export & print --------------------------------------------------
  {
    id: 'export.images',
    name: 'Export slides as PNG and SVG',
    category: 'export',
    editing: 'none',
    rendering: 'partial',
    roundTrip: 'none',
    tier: 'community',
    notes: 'The export writes image files, not the presentation, so round trip does not apply.',
  },
  {
    id: 'export.pdf',
    name: 'PDF export and printing',
    category: 'export',
    editing: 'none',
    rendering: 'planned',
    roundTrip: 'none',
    tier: 'premium',
    notes:
      'Export PDFs with @pptx-editor.dev/pptx-to-pdf under the EigenPal Pro License. Hidden slides are omitted, and approximated or unsupported content produces diagnostics. Fidelity gaps remain, so the rendering status stays planned.',
    docsLink: '/docs/export/pdf',
  },
];

/** Lookup by stable id; used by <FeatureBadge id="..."/>. */
export const pptxFeatureById: Record<string, PptxFeature> = Object.fromEntries(
  pptxFeatures.map((f) => [f.id, f])
);
