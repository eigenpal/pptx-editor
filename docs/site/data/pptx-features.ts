/**
 * PPTX feature support matrix: single source of truth.
 *
 * Copied from DOCX Editor `docs/site/data/word-features.ts` at commit 29535a640510 (the type
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
 * Notes rule: notes render inside a table cell, so keep them short. Write Simplified Technical
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

const SAVE_CHECKED = 'Saved files pass schema validation and open with no repair prompt.';

const TABLE_SAVE = 'A saved copy, before and after table edits, opens with no repair prompt.';

export const pptxFeatures: PptxFeature[] = [
  // --- Text & formatting -----------------------------------------------
  {
    id: 'text.basic-formatting',
    name: 'Bold, italic, underline, strikethrough',
    category: 'text',
    editing: 'full',
    rendering: 'partial',
    roundTrip: 'full',
    tier: 'community',
    notes: `Underline styles other than single are approximate. The Font dialog sets the underline color (a:uFill). ${SAVE_CHECKED}`,
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
      "Theme fonts (+mj-lt, +mn-lt) resolve through the theme; the font picker's Headings and Body rows write these references. Missing fonts use metric-compatible substitutes. Embedded fonts decode in memory. Only installable and editable embedded faces enter the session. Other embedded faces use available host faces or substitutes; save preserves their original bytes. Hosts supply more faces through the `fonts` option and `setFonts` (the font configuration API of DOCX Editor); host faces rank after embedded faces and before the substitutes. A host face changes measurement only, and its fidelity is not measured.",
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
      'Superscript and subscript toggles. The Font dialog sets the offset, caps, and strikethrough kinds. Small caps and script size are approximate.',
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
      'Five modes rewrite the letters and keep the formatting; at a caret, the word around it. Shift+F3 steps through Capitalize Each Word, UPPER CASE, and lower case. A letter whose other case has a different length (ß) stays, and the run language is not used.',
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
      'Link selected text to a web address, an email address with a subject, a slide, a custom show, or another file. A slide link goes to a slide number, or the first, last, next, or previous slide. A custom show link can return to its slide after the show. A file link takes a checked relative path, and the editor never opens it. A click into a link shows its target with Copy, Edit, and Remove, and a slide link shows a button that goes to its slide. In the slide show, a slide link moves the show; custom show and file links do nothing. A slide link follows its slide when slides move. Links to programs are kept but never written or run, as a security limit. URLs pass an allowlist. Hyperlink text color is approximate.',
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
      'Typing re-lays out with HarfBuzz. The caret steps and deletes by grapheme cluster. Left and Right collapse a selected range to its start or end before typing. Alt or Ctrl with Left and Right steps by word, Ctrl or Cmd with Home and End goes to the start or end of the text, and a double-click or a format painter click takes a word (Unicode word boundaries) and a triple-click the paragraph; a drag after a double-click or triple-click extends by whole words or paragraphs. Home and End go to the edges of the laid-out line. An input method composes at the caret: the text being composed shows there underlined, and the result goes in once at composition end as one undo step. Keys during a composition, and the key that ends it, run no shortcut. Validate your target input methods.',
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
      "East Asian characters measure and draw with the run's ea font, complex-script characters with its cs font, and symbol characters with its sym font. A theme with an empty ea or cs font uses its script font (a:font) for the run's language (lang, altLang). Chinese and Japanese text breaks between characters and Korean text at spaces. eaLnBrk, latinLnBrk, hangingPunct, the presentation kinsoku lists (p:kinsoku), and justified and distributed East Asian lines follow the reference rules. A missing East Asian font falls back to a host CJK face, such as Noto Sans CJK JP from @docx-editor.dev/fonts-cjk, and a character that the run's face lacks uses another loaded face. Right-to-left text is approximate. Bidi line layout is not built, so the caret moves in logical order in right-to-left text. Caret steps keep combining marks and clusters whole, and word steps use dictionary words in Thai, Lao, Khmer, Myanmar, and CJK text.",
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
      'Copy and paste text, shapes, and slides in the editor and between editors. Text from other applications pastes as HTML with its formatting, list levels, and bullets; a picture pastes at the size its resolution gives. Copied text leaves as HTML, and copied shapes as a picture, or as text for text boxes. Shapes copied in another application paste as a picture: browsers cannot read their shape format.',
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
      'Line heights and baselines follow measured rules, and each baseline rounds to a whole point from the top of the text. Glyph advances are approximate. The Paragraph dialog sets space before and after in points or lines, and line spacing, on mixed selections too.',
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
      'The Paragraph dialog sets the indent before text and a first-line or hanging indent. Indent and outdent change the list level. Text at a tab stop aligns left, centered, right, or on its decimal point. Tab stops are not editable. No ruler.',
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
      'Stored font scale and line spacing reduction render with whole-point sizes and a subtracted reduction, as measured. Refit after an edit starts from 100 %, not from the last state: a title shrinks to 90 % at most and then overflows, as measured; other text uses the measured step order down to a 55 % font scale, and estimated steps below it. Empty paragraphs at the end of the text do not count, as measured. The font size box shows the size the text is drawn at; the Font dialog shows and sets the size the text asks for. Increase steps up from that size, and Decrease on whole selected shapes steps down from the drawn size. The toolbar and the Format panel switch between no autofit, shrink text on overflow, and resize shape to fit text. Overflowing centered text is approximate. Resize to fit text across rotated text, several columns, or a shape that is not a rectangle is approximate.',
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
      'The Text Direction and Columns menus set horizontal, rotated, stacked, and upright East Asian text, and one to three columns. More Columns sets up to 16 columns and the spacing. Mongolian and right-to-left stacked text are refused and approximate. Upright East Asian text has no vertical metrics.',
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
      'Text effects > Transform sets or removes a preset warp. Adjust values are set by command only. Glyphs bend into the warp. The caret and the selection use flat text. Lines in some multi-band warps sit lower. Warped glyphs are not shaped, so ligatures and complex scripts lose their forms. Vertical, rotated, and diagram text draws flat.',
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
      'Toggle bullets and numbering. The galleries set common bullet characters and a:buAutoNum numbering schemes, and the Bullets and Numbering dialog sets the size, color, start number, and a custom character and font. Enter continues the list, also on an empty item. At the start of a list item or across paragraphs, Tab demotes a level and Shift+Tab promotes one; elsewhere both keys type a tab. Backspace at the start of a list item removes its bullet and keeps its level. Typing 1., 1), *, or - and a space at the start of a paragraph starts a list (AutoCorrect). Where a symbol font is not installed, a Wingdings or Symbol bullet draws its Unicode twin (approximate). The dialog sets a picture bullet from a PNG, JPEG, or GIF file. A picture bullet draws 0.7 times the bullet size high, on the baseline. A missing or unreadable bullet picture draws no bullet.',
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
    notes: 'Group, ungroup, and edit inside a group. Group effects do not draw.',
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
      'Drag a connector end near a shape to glue it to a connection site. Glued connectors follow a shape that moves or resizes, and elbow and curved connectors route again. A connector that moves on its own lets go of its ends. Deleting a shape unglues the connector ends attached to it.',
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
      'Smart guides show while you drag. Distribute needs three objects. Align and distribute use the unrotated box of a turned object. There is no Align to Slide setting.',
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
      'The Rotate control flips and turns 90° on the slide, also inside groups. Every resize handle of one object keeps a locked ratio; a resize of several objects ignores it. The Format tab of the side panel has a Lock aspect ratio check box, and a typed width or height of a locked object sizes the other side too. `setLockAspectRatio` sets the lock.',
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
      'The side panel has a Selection tab that lists the objects of the slide as a tree. It selects objects, hidden ones too, shows or hides them with the eye button or the H key, renames them, and moves them forward or backward. The toolbar Selection pane button turns a mounted side panel to that tab, and a host can listen to show a closed panel. The commands also work inside mc:AlternateContent. There, a move, a resize, a rotation, or a flip changes the copies in the other branches too. Other edits of a drawn choice remove the out-of-date copies in the other branches: an approximation.',
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
    notes:
      'The six theme style rows in seven colors. The cells get the same p:style that the table style defines for them. No preset (transparent) styles.',
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
      'Linear and circle path gradients are measured. Two-stop gradient colors (a measured rule that the specification does not state), rectangle path gradients (seams at the joins), shape path gradients (bands), a rectangle or shape gradient focus outside the shape, gradient tile rectangles on linear, rectangle, and shape gradients and inside the shape on circle gradients, gradient flips, and fill rotation are approximate.',
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
      'The fill menu sets a preset pattern with two colors (`setShapePattern`), or a stretched picture (`setShapePicture`, which can also tile). Picture crop, offset, scale, and transparency of a fill are not editable. Tiles use the image resolution, or 72 DPI when it states none, and keep their recolor; the 72 DPI rule is measured on PNG only, and the tile size of a PNG that states no resolution is approximate. Pattern tiles are 8 × 8 pt on the slide grid; 12 of the 54 preset bitmaps are measured, and the others are approximate.',
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
      'Dashes and line ends draw at measured sizes, at every line width. Compound line bands are not measured.',
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
      'Gallery presets set and clear the shadow. Inherited and theme effects stay. Groups, tables, and charts are refused, and so is a theme effect style with a:effectDag. No perspective shadows. The shadow blur is measured for small radii.',
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
      'The drawing module of @pptx-editor.dev/pro paints these effects and uses the EigenPal Pro License. Evaluation is permitted. Production use requires a commercial agreement. Without it, shapes and pictures paint without them. PDF export always paints them. With the module, the Shape Effects menu sets and clears each effect, glows in the six accent colors, and Picture Styles offers the reflection and soft edge styles. Without it, the chrome does not offer the choices that add these effects. It still offers No reflection, No glow, and No soft edges when the selection has that effect, and No shadow also removes an inner shadow. The commands still set and clear each effect. Save keeps them in every edition. No Preset, Bevel, or 3-D Rotation. Glow, inner shadow, and soft edge blurs are measured for small radii. The reflection blur is not measured.',
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
    editing: 'none',
    rendering: 'none',
    roundTrip: 'preserved',
    tier: 'community',
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
      'Insert, move, resize, rotate, and Change Picture. Change Picture fits the new image in the old frame; this rule is not measured. Change Picture takes PNG, JPEG, and GIF only, not SVG.',
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
      'Corrections, Recolor, Transparency, and Reset Picture in the `image` group. Reset and Corrections refuse a picture with `a14:imgProps` picture effects. Saturation, color tone, and artistic effects are not built. Luminance, recolor, and picture shape fills are approximate.',
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
      'Crop mode with crop handles, crop to a preset shape, Fill, Fit, aspect ratios, and crop position in Format Picture. A picture locked against cropping (`a:picLocks/@noCrop`) refuses crop mode, crops, the Format Picture crop fields, and Reset Picture & Size of a cropped picture; `noResize` also refuses a crop that changes the frame size (Fill and Fit pass); `noResize`, `noRot`, and `noChangeShapeType` refuse a new size, rotation, or shape (Crop to Shape, Picture Styles, Reset Picture) in the commands, and Change Picture honors them, but the canvas handles still resize and rotate. A tiled or inset picture fill (`a:tile`, an inset `a:fillRect`) refuses crop mode and crops. Fill, Fit, and aspect ratio centering are not measured. The editor ignores the EXIF orientation of a JPEG and draws and sizes the stored pixels, as measured.',
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
      'Five of the gallery styles, and border color, weight, and dashes. The written style XML is not measured. Dashes on a picture border are not measured, and round dots draw as squares.',
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
      'Description, title, and "Mark as decorative" for one selected object. On the slide and in the slide show, a screen reader reads the description of a picture. It skips a decorative picture and a picture with no description. When Tab selects an object, the screen reader also hears its description. Automatic descriptions are not built.',
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
    notes: `Click in a cell and type. Text commands work in a cell, a cell range, or the whole table. Tab and Shift+Tab move between cells. Tab in the last cell adds a row. Shift+Arrow at a cell's edge selects a cell range. Copy, cut, and paste a cell range as tab-separated text; a paste larger than the table adds rows and columns. ${TABLE_SAVE}`,
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
    notes: `Insert a table from a grid. Insert and delete rows and columns, and select a row, a column, or the table. A new table fills an empty content or table placeholder: the selected one, or else the first on the slide. Deleting the table does not bring the empty placeholder back. ${TABLE_SAVE}`,
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
    notes: `Merge keeps the text of every cell. Split works on one cell. A merged cell splits only along even grid lines. A table whose merge markup disagrees with itself follows the inconsistent cell merges rule under Approximations. When peers in a collaboration session change the rows, columns, or merged cells of one table at once, the canvas draws the table as a save writes it, with valid merge markers. ${TABLE_SAVE}`,
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
    notes: `Toggle Header Row, Total Row, Banded Rows and Columns, and First and Last Column. The gallery lists the deck's styles and the default style only. ${TABLE_SAVE}`,
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
    notes: `Shading takes a color or No Fill. Borders draw with a pen of a color, a preset dash, and a weight. Diagonal borders and cell effects are not editable. ${TABLE_SAVE}`,
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
    notes: `Drag a grid line (the line under the pointer is marked), distribute rows and columns, type a row height or column width in the Format panel, or resize the frame by its handles (the columns and rows scale; the share of each column and row in that edit is not measured). Row heights stay minimums, so rows grow to fit text. ${TABLE_SAVE}`,
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
      'The drawing module of @pptx-editor.dev/pro draws charts from the values cached in the chart part, in the editor and in the PDF export. It draws bar and column charts (clustered, stacked, and percent stacked), line, area, scatter, pie, doughnut, and radar charts, with titles, legends, axes, gridlines, data labels, and number formats. 3-D, of-pie, stock, bubble, and surface charts, and charts with trendlines or error bars, show a labeled placeholder. Without the module, every chart shows the placeholder. The Edit Data dialog and the setChartData command change series names, category labels, and values, and add or remove series and categories. They edit bar, column, line, area, pie, doughnut, radar, and scatter charts. Bubble, stock, surface, of-pie, and combination charts, and filtered or multi-level data, stay read-only. The embedded workbook gets the same cells. A workbook with formulas, filters, or merged cells keeps its old cells. A linked workbook is never opened, and its automatic update is turned off. Dates show as serial numbers. Other chart formatting is not editable. The module uses the EigenPal Pro License. Evaluation is permitted. Production use requires a commercial agreement.',
    docsLink: '/docs/guides/charts',
  },
  {
    id: 'objects.diagrams',
    name: 'Diagrams',
    category: 'objects',
    editing: 'none',
    rendering: 'partial',
    notes:
      'A diagram draws from its saved drawing. A diagram with no saved drawing shows a placeholder. Diagrams are read-only.',
    roundTrip: 'preserved',
    tier: 'community',
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
      'Video and audio show their poster picture. An OLE object shows the PNG, JPEG, GIF, or BMP picture that its p:oleObj/p:pic stores, or a labeled placeholder when it has none. Nothing plays or runs. In mc:AlternateContent, a choice that requires only p14 and holds shapes draws; otherwise the fallback draws.',
  },
  {
    id: 'objects.equations',
    name: 'Equations',
    category: 'objects',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'preserved',
    tier: 'community',
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
      'Empty placeholders show their prompt and a dotted outline in the editor only. A click in an empty text placeholder or text box places the caret. Empty content and table placeholders show a Table button, and empty content and picture placeholders show a Picture button. The new object goes in place of the placeholder. The buttons cover tables and pictures only, and a keyboard user inserts from the Insert menu, which fills the placeholder too.',
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
      'Edit notes text, bold, italic, underline, and strikethrough, from the notes pane or the toolbar. A format at a caret outside a word does not carry to the next typed text. Other notes formatting is kept but not editable.',
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
      'Turn slide numbers, the footer, and the date on or off (`setSlideNumbers`, `setFooter`, `setDateTime`), through the placeholders of each layout or its master. The commands act on all slides or on `indices`, and `hideOnTitleSlides` leaves title slides out. Turning one on without `text` keeps what the slides show. A slide that shows a footer or a date keeps its box and format. New date text is a `datetime1` field, or fixed text with `fixed: true`. Editing is partial: a new date field always has the `datetime1` format.',
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
      'Pick Widescreen, Standard, A4, Letter, or a custom size from 1 to 56 inches, in either orientation, in **Page setup** (`setSlideSize`). The content scales to fit by default: boxes, tables, chart text, and the text sizes, spacing, and insets that the file sets scale by one factor and stay centered. Editing is partial: line widths, effect distances, and values left to their defaults keep their values, diagram text keeps its size, and the notes pages and drawing guides keep their old layout. Rendering is measured on 16:9, ISO A4 (842 × 595.5 pt), and 400.25 × 300.5 pt slides, which export to PDF pages rounded to whole points. Rendering is partial: portrait and very wide slides are not measured, and the strip past a rounded-down page is not compared.',
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
      'Add, rename, remove, move, and collapse sections from the section header menu in the rail and the slide sorter, and from the toolbar section parts (`section.*` slots). The default editor has no Section menu. Sections cannot be dragged.',
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
      'Change layout and reset slide work. Placeholders without a place, animated empty placeholders, and placeholders in mc:AlternateContent keep their box: an approximation. The placeholder match order is not measured. Themes and the slide master cannot be edited.',
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
      'You can add, reorder, time, and remove the entrance, emphasis, and exit effects of the main sequence, from 13 presets. You can set fade, push, wipe, split, and cut transitions with a duration and advance timing. Motion paths and interactive sequences are listed, and you cannot add them. Timing in markup-compatibility content is read-only. The editor does not play effects or transitions. Deleting, cutting, or ungrouping shapes, deleting a table, and turning slide numbers off remove the animations of the removed shapes. A paragraph animation range (`p:pRg`) follows its paragraphs when you add, delete, split, or join paragraphs. An effect whose paragraphs are all deleted goes too. Between paragraphs with the same text, an edit in the editor keeps each effect with its own paragraph. An automation edit that names no selection decides by the text, so an effect can stay with the wrong one of two identical paragraphs.',
  },

  // --- Comments & review -----------------------------------------------
  {
    id: 'review.comments',
    name: 'Comments',
    category: 'review',
    editing: 'none',
    rendering: 'none',
    roundTrip: 'preserved',
    tier: 'premium',
    notes:
      'Comments are kept on save. The review module of @pptx-editor.dev/pro reads them and uses the EigenPal Pro License. Evaluation is permitted. Production use requires a commercial agreement. No review UI yet.',
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
      'The Yjs session of @pptx-editor.dev/pro syncs edits and concurrent typing. Undo reverts your own changes and keeps what a peer wrote into an element you created: the attributes, children, and text that the peer wrote stay, without yours. Text a peer typed into a run whose format you undo stays, and a paragraph you add keeps what a peer typed in it. Text a peer typed into a shape, a table row, or a cell that you add goes with your undo and comes back with your redo. Two property elements written at once (a:pPr, a:rPr, a:tcPr, a:ln) merge; for a fill or effects, one version shows and the other comes back when the winner undoes. A first line width also writes a default line color, which can show instead of the line color that a peer sets at the same time. One attribute written by two peers at once keeps one value. Slides, pictures, links, shapes, tables, and comments that two peers add or copy at once keep separate ids; one notes slide or slide-number placeholder that two peers add first can still show twice. Two peers who format one run at once keep the text once with the format of one peer; when one of them changed the text of the run first, the two formats do not count as one split, so the text can show twice, and a piece the losing peer edited shows twice. Text that a peer types, deletes, or splits with Enter while a format of another peer moves it goes with the moved text. Two peers who press Enter in one paragraph at once keep the text of each paragraph once. A run that a peer deletes keeps the text another peer types into it at the same time, also when an Enter moved the run and a peer typed into it at once. Text typed into a paragraph that a peer removes at the same time (a delete across paragraphs, a join) stays, in a shape and in a table cell: the paragraph shows again at its old place with only that text, and an undo of the removal puts the rest back. Two peers who type at once into a table cell that holds no text body yet both keep their text in the cell, in the same order for every peer, and each undo removes only the text of its own peer. An undo or a redo after a peer formatted, split, or pressed Enter in the same text shows each character once, next to the characters around it, and two peers who undo their delete of one character show it once; a run that the split of the peer made keeps the format it copied, so the undo of a format can leave it on that run. Three peers whose undo, redo, and Enter meet over the same text show each character once, but such edits can rarely lose a character that a peer typed beside them. A format and an Enter of a peer at once inside the same text keep the text in order, and so does an undo of either one; an undo of the format and then of the Enter can show the text out of order. Content moved into content that a peer deletes at once goes back to its old place in the file. Two peers who join neighboring paragraphs at once keep the text of all of them, and an undo of one join puts the text back in its paragraphs, also after a third peer typed into the joined paragraph. A paragraph that two peers move at once shows in one place, and a later delete of that place removes it for every peer. Text in a cell that a peer merges shows once in the merged cell, also when another peer types into that cell or inserts, deletes, splits, or merges rows and columns across it at the same time; some sequences of several such edits can still lose text, for example a merge and deletes of two of its rows by two peers at once. When peers delete every row of a merged cell at once, the merged text goes with the rows. After a vertical merge and a concurrent delete of one of its covered rows, the merged cell ends at its last row, and the next row shows on every replica and in the saved file. The canvas shows the selections of peers with name and color. No provider entries yet. The session uses the EigenPal Pro License. Evaluation is permitted. Production use requires a commercial agreement.',
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
      '@pptx-editor.dev/editor-api edits slides, shapes, fill and line, text, tables, hyperlinks of slides and text ranges, tags, document properties and custom document properties, custom XML parts, slides imported from another presentation, slide, layout, and master backgrounds, theme colors, table cell borders and margins, table style options, merged areas, and preset adjust values on a server or in an open editor, and binds shapes by id in memory for the lifetime of a runtime, never in the file. The request context gives sync pass-through values, loadRecursive, trace, and debugInfo. Every editing method and property write of the selected Office.js subset has its exact signature. Slide and shape pictures are drawn in a browser, or on a server through a rasterizer that the host supplies, and a shape picture is cut to the shape box; an imported slide with threaded comments is refused, while other comments come along; a background picture is stretched, never tiled, and ShadeFromTitle is a read state; a table style must be defined in the presentation, except the built-in default. The package uses the EigenPal Pro License. Evaluation is permitted. Production use requires a commercial agreement.',
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
      '@pptx-editor.dev/pptx-to-pdf uses the EigenPal Pro License. Evaluation is permitted. Production use requires a commercial agreement. The PDF export draws shapes, solid and gradient fills, cropped PNG and JPEG pictures, tables, shadows, glow, inner shadows, reflections, and searchable text. It exports in one call or from a reusable session, and leaves hidden slides out. The React and Vue demos export through a server function: Export > Presentation as PDF. It draws dashes and line ends with the same geometry as the editor. Approximations, each with a diagnostic code: shadow and glow blur, opaque silhouettes, inner shadows, soft edges, reflection blur, gradient lines, gradients with more than 64 stops, fills past the shape box, tiled pictures, picture color effects, compound lines, underline styles other than single and double, and some text styles; SVG pictures draw from their PNG fallback. The border of a picture is drawn outside the picture. With the optional @docx-editor.dev/fonts-cjk installed, East Asian text in a missing font draws with Noto Sans CJK JP. Not drawn yet: pattern fills, picture fills on text, text outlines, other picture formats, diagrams, OLE objects with no picture, and other graphic frames, and table frames with no grid. Charts draw through the drawing module of @pptx-editor.dev/pro, which the export always installs; chart types outside that module show a placeholder with the frame.chart code. The export fidelity is not measured yet, so rendering stays planned. The export writes a PDF, not the presentation, so round trip does not apply.',
  },
];

/** Lookup by stable id; used by <FeatureBadge id="..."/>. */
export const pptxFeatureById: Record<string, PptxFeature> = Object.fromEntries(
  pptxFeatures.map((f) => [f.id, f])
);
