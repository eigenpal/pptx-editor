// The evidence of the feature matrix, kept out of `pptx-features.ts` so that the row type stays
// equal to the DOCX Editor row type and the docs site syncs only public data. Each entry is
// keyed by a row `id` of `pptx-features.ts`; `pptx-features.test.ts` checks both lists.

/** The evidence for one row of the feature matrix. */
export interface FeatureEvidence {
  /** Base names of the synthetic decks in `fixtures/` that show the feature. */
  readonly fixtures: readonly string[];
  /**
   * Ids of the known approximations (`APPROXIMATIONS` in `packages/core/src/layout/
   * approximations.ts`) that the feature owns. Name an approximation on the rows whose behavior
   * it changes, not on every row whose fixture shows it: a slide title makes layout report
   * `line-model`, but only text rows own it.
   */
  readonly approximations?: readonly string[];
}

/** The evidence of each row, by row `id`, in matrix order. */
export const FEATURE_EVIDENCE: Readonly<Record<string, FeatureEvidence>> = {
  // --- Text & formatting -------------------------------------------------
  'text.basic-formatting': {
    fixtures: ['text-runs', 'font-effects'],
    approximations: ['underline-styles'],
  },
  'text.font-family-size': {
    fixtures: ['text-runs', 'font-substitution'],
    approximations: ['line-model', 'glyph-advances', 'missing-glyphs'],
  },
  'text.color-highlight': {
    fixtures: ['text-runs', 'colors', 'fills-lines'],
    approximations: ['gradient-text'],
  },
  'text.caps-baseline': {
    fixtures: ['text-runs', 'font-effects'],
    approximations: ['line-model', 'small-caps', 'script-size'],
  },
  'text.character-spacing': {
    fixtures: ['text-runs', 'font-effects'],
    approximations: ['glyph-advances'],
  },
  'text.change-case': { fixtures: ['font-effects'] },
  'text.hyperlinks': { fixtures: [], approximations: ['hyperlink-text'] },
  'text.input-method': { fixtures: ['text-runs'] },
  'text.complex-scripts': {
    fixtures: [
      'rtl-text',
      'east-asian-line-breaking',
      'east-asian-theme-fonts',
      'east-asian-typography',
    ],
    approximations: ['script-fonts', 'rtl'],
  },
  'text.find-replace': { fixtures: [] },
  'text.format-painter': { fixtures: [] },
  // --- Paragraphs & text boxes -------------------------------------------
  'paragraphs.alignment': {
    fixtures: ['alignment-anchor', 'east-asian-line-breaking'],
    approximations: ['line-model', 'glyph-advances', 'anchoring'],
  },
  'paragraphs.spacing': {
    fixtures: ['spacing', 'line-model', 'paragraph-format'],
    approximations: ['line-model', 'trailing-break'],
  },
  'paragraphs.indent-tabs': {
    fixtures: ['bullets-numbering', 'paragraph-format'],
    approximations: ['line-model', 'tab-stops'],
  },
  'paragraphs.autofit': { fixtures: ['autofit', 'line-model'], approximations: ['line-model'] },
  'paragraphs.text-direction-columns': {
    fixtures: ['vertical-columns', 'vertical-east-asian', 'text-direction'],
    approximations: ['line-model', 'script-fonts', 'ea-vert', 'word-art-vert', 'other-vert'],
  },
  'paragraphs.text-warp': { fixtures: [], approximations: ['text-warp'] },
  // --- Bullets & numbering -----------------------------------------------
  'lists.bullets-numbering': {
    fixtures: ['bullets-numbering', 'bullets-gallery', 'rtl-text'],
    approximations: [
      'line-model',
      'aligned-bullets',
      'numbering-schemes',
      'symbol-bullets',
      'trailing-break',
    ],
  },
  // --- Shapes, groups & connectors ---------------------------------------
  'shapes.preset-geometry': { fixtures: ['preset-shapes'], approximations: ['path-fill-modes'] },
  'shapes.custom-geometry': {
    fixtures: ['custom-geometry', 'follow-ups'],
    approximations: ['path-start'],
  },
  'shapes.groups': { fixtures: ['groups-transforms'] },
  'shapes.connectors': { fixtures: ['lines-arrows', 'line-ends-dashes'] },
  'shapes.arrange': { fixtures: ['arrange-flip-hidden'] },
  'shapes.flip-rotate': {
    fixtures: ['arrange-flip-hidden', 'groups-transforms'],
    approximations: ['flipped-text', 'fill-rotation'],
  },
  'shapes.selection-pane': { fixtures: ['arrange-flip-hidden', 'preserve-unsupported'] },
  'shapes.styles-gallery': { fixtures: ['shape-styles'], approximations: ['blur'] },
  // --- Fills & lines -----------------------------------------------------
  'fills-lines.solid-gradient': {
    fixtures: ['fills-lines', 'gradients', 'colors', 'fill-fidelity'],
    approximations: [
      'gradient-colors',
      'path-gradient-focus',
      'gradient-tiles',
      'rect-path-gradient',
      'shape-path-gradient',
      'fill-rotation',
    ],
  },
  'fills-lines.pattern-picture': {
    fixtures: ['fills-lines', 'fill-fidelity', 'follow-ups'],
    approximations: ['pattern-bitmaps', 'fill-rotation', 'tiled-pictures'],
  },
  'fills-lines.lines': {
    fixtures: ['lines-arrows', 'line-ends-dashes', 'fills-lines'],
    approximations: ['inset-lines', 'compound-lines'],
  },
  'fills-lines.theme-colors': {
    fixtures: ['colors', 'follow-ups'],
    approximations: ['color-transforms', 'gray-transform'],
  },
  // --- Effects -----------------------------------------------------------
  'effects.outer-shadow': {
    fixtures: ['effects', 'shape-effects-presets'],
    approximations: ['blur', 'shadow-skew'],
  },
  // Painted by the drawing module of @pptx-editor.dev/pro: the harness installs it
  // (`examples/vite/src/harness.ts`), and `packages/pro/src/drawing/__tests__/effects-paint.test.ts`
  // proves that the free engine paints none of these effects.
  'effects.glow-reflection': {
    fixtures: ['effects', 'shape-effects-presets', 'picture-editing'],
    approximations: ['blur', 'glow-soft-edge'],
  },
  'effects.fill-overlay': {
    fixtures: ['fill-fidelity'],
    approximations: ['fill-overlay', 'effects-not-drawn'],
  },
  'effects.3d-and-other': { fixtures: ['effects'], approximations: ['effects-not-drawn'] },
  // --- Pictures ----------------------------------------------------------
  'pictures.raster-svg': {
    fixtures: ['pictures', 'fill-fidelity'],
    approximations: ['picture-shape-fill'],
  },
  'pictures.adjustments': {
    fixtures: ['pictures', 'fill-fidelity', 'picture-editing'],
    approximations: ['picture-luminance', 'picture-recolor'],
  },
  'pictures.crop': { fixtures: ['pictures', 'picture-editing'] },
  'pictures.styles-border': {
    fixtures: ['picture-editing'],
    approximations: ['blur', 'glow-soft-edge', 'picture-border-dots', 'inset-lines'],
  },
  'pictures.alt-text': { fixtures: ['picture-editing'] },
  'pictures.emf-wmf': { fixtures: [], approximations: ['emf-wmf'] },
  // --- Tables ------------------------------------------------------------
  'tables.render': {
    fixtures: ['tables', 'table-editing'],
    approximations: ['default-table-style', 'coincident-borders', 'inconsistent-merges'],
  },
  'tables.cell-text': {
    fixtures: ['tables', 'table-editing'],
    approximations: ['line-model', 'rotated-cell-text'],
  },
  'tables.structure': { fixtures: ['table-editing'] },
  'tables.merge-split': { fixtures: ['tables', 'table-editing'] },
  'tables.styles': {
    fixtures: ['tables', 'table-editing'],
    approximations: ['default-table-style'],
  },
  'tables.cell-format': {
    fixtures: ['tables', 'table-editing'],
    approximations: ['gradient-colors', 'table-cell-effects', 'coincident-borders'],
  },
  'tables.sizes': { fixtures: ['tables', 'table-editing'] },
  // --- Charts, diagrams, media & OLE -------------------------------------
  'objects.charts': {
    fixtures: ['charts', 'chart-sizes', 'preserve-unsupported', 'slide-structure'],
  },
  'objects.diagrams': {
    fixtures: ['diagrams'],
    approximations: ['diagram-drawing', 'line-model'],
  },
  'objects.media-ole': { fixtures: ['preserve-unsupported', 'slide-structure'] },
  'objects.equations': { fixtures: [] },
  // --- Slides, layouts & notes -------------------------------------------
  'slides.structure': { fixtures: ['slide-structure', 'slide-layouts'] },
  'slides.hidden': { fixtures: ['slide-layouts'] },
  'slides.placeholder-prompts': { fixtures: ['slide-layouts'] },
  'slides.notes': { fixtures: ['slide-layouts'] },
  'slides.header-footer': { fixtures: ['placeholders-background'] },
  'slides.size': { fixtures: ['slide-size', 'slide-size-rounding'] },
  'slides.sections': { fixtures: ['slide-structure', 'slide-layouts'] },
  'slides.sorter': { fixtures: ['slide-layouts'] },
  // --- Themes, masters & inheritance -------------------------------------
  'themes.placeholder-inheritance': {
    fixtures: ['placeholders-background'],
    approximations: ['content-placeholders'],
  },
  'themes.backgrounds': {
    fixtures: ['placeholders-background'],
    approximations: ['gradient-colors'],
  },
  'themes.change-layout-theme': { fixtures: ['slide-layouts'] },
  // --- Presenting --------------------------------------------------------
  'presenting.slideshow': { fixtures: [] },
  'presenting.animations-transitions': { fixtures: [] },
  // --- Comments & review -------------------------------------------------
  'review.comments': { fixtures: ['slide-structure'] },
  // --- Collaboration & automation ----------------------------------------
  'collaboration.realtime': { fixtures: [] },
  'collaboration.automation-api': { fixtures: [] },
  // --- Export & print ----------------------------------------------------
  'export.images': { fixtures: [] },
  'export.pdf': {
    fixtures: [
      'text-runs',
      'preset-shapes',
      'pictures',
      'gradients',
      'tables',
      'effects',
      'fills-lines',
      'groups-transforms',
    ],
  },
};
