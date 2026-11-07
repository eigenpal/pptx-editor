// The approximations the engine knows it makes (AGENTS.md: "Approximations are listed"). Each
// entry's title is a bold item under "Known approximations" on the public fidelity page
// (docs/site/content/pptx-fidelity.mdx); a test keeps the two lists equal. Layout reports the
// entries a slide uses as `approximation: <id>` diagnostics, and the fidelity harness lists them
// for every slide. Entries that layout cannot see in a laid-out slide (`perSlide: false`) are
// listed once, in the report header.

import type {
  RelRect,
  ResolvedEffects,
  ResolvedFill,
  ResolvedLine,
  ResolvedTextBody,
} from '../cascade/index.ts';
import type { LayoutItem, Matrix } from './slide-layout.ts';
import { isExactScheme, isSymbolFont } from './text/autonumber.ts';
import { hasCharacters } from './text/text-layout.ts';
import { isIdeographic } from './text/line-break.ts';

/** One known approximation. */
export interface Approximation {
  readonly id: string;
  /** The bold title of its item on the fidelity page, without the period. */
  readonly title: string;
  /** Whether layout detects it on a slide. */
  readonly perSlide: boolean;
}

/** Every known approximation, in the fidelity page order. */
export const APPROXIMATIONS: readonly Approximation[] = Object.freeze([
  { id: 'line-model', title: 'Line model', perSlide: true },
  { id: 'glyph-advances', title: 'Glyph advances', perSlide: false },
  { id: 'color-transforms', title: 'Color transform spaces', perSlide: false },
  { id: 'gray-transform', title: 'Gray transform', perSlide: false },
  { id: 'blur', title: 'Blur', perSlide: true },
  { id: 'shadow-skew', title: 'Shadow skew', perSlide: true },
  { id: 'glow-soft-edge', title: 'Glow and soft edge', perSlide: true },
  { id: 'preset-shadow', title: 'Preset shadows', perSlide: true },
  { id: 'three-d', title: '3-D shapes', perSlide: true },
  { id: 'effects-not-drawn', title: 'Effects not drawn', perSlide: true },
  { id: 'fill-overlay', title: 'Fill overlay', perSlide: true },
  { id: 'gradient-colors', title: 'Gradient colors', perSlide: true },
  { id: 'path-gradient-focus', title: 'Path gradient focus', perSlide: true },
  { id: 'gradient-tiles', title: 'Gradient tiles', perSlide: true },
  { id: 'rect-path-gradient', title: 'Rectangle path gradients', perSlide: true },
  { id: 'shape-path-gradient', title: 'Shape path gradients', perSlide: true },
  { id: 'pattern-bitmaps', title: 'Pattern bitmaps', perSlide: true },
  { id: 'path-fill-modes', title: 'Path fill modes', perSlide: true },
  { id: 'fill-rotation', title: 'Fill rotation', perSlide: true },
  { id: 'gradient-text', title: 'Gradient text', perSlide: true },
  { id: 'tiled-pictures', title: 'Tiled pictures', perSlide: true },
  { id: 'picture-shape-fill', title: 'Picture shape fill', perSlide: false },
  { id: 'picture-luminance', title: 'Picture brightness and contrast', perSlide: false },
  { id: 'picture-recolor', title: 'Picture recolor', perSlide: true },
  { id: 'picture-border-dots', title: 'Dashes on picture borders', perSlide: true },
  { id: 'emf-wmf', title: 'EMF and WMF pictures', perSlide: false },
  { id: 'inset-lines', title: 'Inset lines', perSlide: false },
  { id: 'hyperlink-text', title: 'Hyperlink text', perSlide: true },
  { id: 'compound-lines', title: 'Compound lines', perSlide: true },
  { id: 'path-start', title: 'Paths without a move', perSlide: false },
  { id: 'flipped-text', title: 'Text in flipped shapes', perSlide: true },
  { id: 'content-placeholders', title: 'Content placeholders', perSlide: true },
  { id: 'small-caps', title: 'Small caps', perSlide: true },
  { id: 'underline-styles', title: 'Underline styles', perSlide: true },
  { id: 'missing-glyphs', title: 'Missing glyphs', perSlide: false },
  { id: 'script-fonts', title: 'East Asian, complex-script, and symbol fonts', perSlide: true },
  { id: 'rtl', title: 'Right-to-left text', perSlide: true },
  { id: 'tab-stops', title: 'Tab stops', perSlide: true },
  { id: 'anchoring', title: 'Anchoring and alignment variants', perSlide: true },
  { id: 'ea-vert', title: 'Upright East Asian text (`eaVert`)', perSlide: true },
  { id: 'word-art-vert', title: 'Stacked text (`wordArtVert`)', perSlide: true },
  { id: 'other-vert', title: 'Other vertical text types', perSlide: true },
  { id: 'script-size', title: 'Superscript and subscript', perSlide: true },
  { id: 'aligned-bullets', title: 'Bullets on aligned lines', perSlide: true },
  { id: 'numbering-schemes', title: 'Numbering schemes', perSlide: true },
  { id: 'symbol-bullets', title: 'Symbol font bullets', perSlide: true },
  { id: 'trailing-break', title: 'Trailing line break', perSlide: true },
  { id: 'default-table-style', title: 'Default table style', perSlide: false },
  { id: 'rotated-cell-text', title: 'Rotated cell text', perSlide: true },
  { id: 'table-cell-effects', title: 'Table cell effects', perSlide: false },
  { id: 'coincident-borders', title: 'Coincident cell borders', perSlide: true },
  { id: 'inconsistent-merges', title: 'Inconsistent cell merges', perSlide: false },
  { id: 'text-warp', title: 'Text warp', perSlide: true },
  { id: 'diagram-drawing', title: 'Diagram drawings', perSlide: true },
  { id: 'equations', title: 'Equations', perSlide: true },
] satisfies Approximation[]);

const CONTENT_PLACEHOLDERS = new Set(['pic', 'tbl', 'chart', 'clipArt', 'dgm', 'media']);
const ROTATED_TEXT = new Set(['vert', 'vert270', 'eaVert', 'wordArtVert', 'mongolianVert']);
const DRAWN_UNDERLINES = new Set(['none', 'sng', 'dbl']);
const RECOLOR_EFFECTS = new Set(['grayscale', 'duotone', 'biLevel']);

/**
 * Whether text holds characters that take the East Asian (`a:ea`) or complex-script (`a:cs`)
 * font slot: CJK, kana, and Hangul; Hebrew, Arabic, and the other right-to-left scripts; Indic
 * scripts; Thai and Lao.
 */
function usesScriptFont(text: string): boolean {
  for (let i = 0; i < text.length; i += 1) {
    const cp = text.codePointAt(i) ?? 0;
    if (cp > 0xffff) i += 1;
    if (cp < 0x0590) continue;
    if (
      isIdeographic(cp) ||
      cp <= 0x08ff || // Hebrew, Arabic, Syriac, Thaana, N'Ko, Samaritan
      (cp >= 0x0900 && cp <= 0x0eff) || // Indic scripts, Sinhala, Thai, Lao
      (cp >= 0x1100 && cp <= 0x11ff) || // Hangul Jamo
      (cp >= 0xfb1d && cp <= 0xfdff) || // Hebrew and Arabic presentation forms
      (cp >= 0xfe70 && cp <= 0xfeff)
    )
      return true;
  }
  return false;
}

/** Collects the approximation ids a slide's content uses. */
class Collector {
  readonly found = new Set<string>();

  add(id: string): void {
    this.found.add(id);
  }

  /** The fill that `a:grpFill` paints in the items being visited (`output/paint.ts`). */
  private group: ResolvedFill | undefined;

  fill(given: ResolvedFill | undefined, rotated: boolean): void {
    const fill = given?.type === 'group' ? this.group : given;
    if (!fill) return;
    if (fill.type === 'pattern') this.add('pattern-bitmaps');
    if (fill.type === 'path' && fill.path === 'rect') this.add('rect-path-gradient');
    if (fill.type === 'path' && fill.path === 'shape') this.add('shape-path-gradient');
    if (twoStops(fill)) this.add('gradient-colors');
    if (fill.type === 'path' && fill.path !== 'circle' && focusOutside(fill.fillToRect))
      this.add('path-gradient-focus');
    if ((fill.type === 'linear' || fill.type === 'path') && fill.tileRect && unmeasuredTile(fill))
      this.add('gradient-tiles');
    if (fill.type === 'picture' && fill.mode === 'tile') this.add('tiled-pictures');
    if (fill.type === 'picture' && fill.effects.some((e) => RECOLOR_EFFECTS.has(e.kind)))
      this.add('picture-recolor');
    if (
      rotated &&
      (fill.type === 'linear' || fill.type === 'path' || fill.type === 'picture') &&
      !fill.rotWithShape
    )
      this.add('fill-rotation');
  }

  line(line: ResolvedLine | undefined, picture = false): void {
    if (!line || line.fill.type === 'none') return;
    this.fill(line.fill, false);
    // Any dash on a picture border: dashes outside a picture are not checked against a reference PDF.
    if (picture && line.dash !== 'solid') this.add('picture-border-dots');
    if (line.compound !== 'sng') this.add('compound-lines');
  }

  /** `overlay` is false when paint drops the fill overlay (more than `MAX_OVERLAY_PATHS` paths). */
  effects(fx: ResolvedEffects | undefined, overlay = true): void {
    if (!fx) return;
    if (
      (fx.outerShadow?.blur ?? 0) > 0 ||
      (fx.reflection?.blur ?? 0) > 0 ||
      (fx.innerShadow?.blur ?? 0) > 0 ||
      (fx.blur?.radius ?? 0) > 0
    )
      this.add('blur');
    if (fx.secondShadow || (fx.outerShadow && fx.outerShadow.blur === 0 && fx.outerShadow.sy !== 1))
      this.add('preset-shadow');
    if (fx.threeD) this.add('three-d');
    if (fx.outerShadow && (fx.outerShadow.kx !== 0 || fx.outerShadow.ky !== 0))
      this.add('shadow-skew');
    if (fx.glow || fx.softEdge) this.add('glow-soft-edge');
    if (fx.fillOverlay && overlay) {
      this.add('fill-overlay');
      this.fill(fx.fillOverlay.fill, false);
    } else if (fx.fillOverlay) this.add('effects-not-drawn');
    if (fx.unsupported.length > 0) this.add('effects-not-drawn');
  }

  text(body: ResolvedTextBody | undefined): void {
    if (!body) return;
    const bp = body.body;
    if (bp.vert === 'eaVert') this.add('ea-vert');
    if (bp.vert === 'wordArtVert') this.add('word-art-vert');
    if (bp.vert === 'mongolianVert' || bp.vert === 'wordArtVertRtl') this.add('other-vert');
    if (bp.anchorCtr || bp.anchor === 'just' || bp.anchor === 'dist') this.add('anchoring');
    if (bp.warp) this.add('text-warp');

    for (const p of body.paragraphs) {
      if (p.runs.length === 0) continue;
      // The line rules are measured on Latin text in the reference fonts; other scripts,
      // vertical text, substituted faces, and overflowing anchored text are not (the
      // fidelity page, "Line model").
      this.add('line-model');
      if (p.rtl) this.add('rtl');
      if (p.align === 'justLow' || p.align === 'thaiDist') this.add('anchoring');
      if (p.bullet && (p.align === 'ctr' || p.align === 'r')) this.add('aligned-bullets');
      if (p.bullet?.type === 'autonum' && !isExactScheme(p.bullet.scheme))
        this.add('numbering-schemes');
      if (p.bullet?.type === 'char' && isSymbolFont(p.bullet.font ?? ''))
        this.add('symbol-bullets');
      if (p.runs[p.runs.length - 1].kind === 'break') this.add('trailing-break');
      // Centered, right, and decimal stops are measured in wrapped left-to-right text
      // (`svg-tab-stops`); right-to-left and unwrapped text are not.
      const tabbed = p.tabs.some((t) => t.align !== 'l') && (p.rtl || bp.wrap !== 'square');
      for (const run of p.runs) {
        if (run.kind === 'math') this.add('equations');
        const props = run.props;
        if (tabbed && run.text.includes('\t')) this.add('tab-stops');
        if (props.caps === 'small') this.add('small-caps');
        if (props.baseline !== 0) this.add('script-size');
        if (props.hyperlink || props.hyperlinkAction) this.add('hyperlink-text');
        if (!DRAWN_UNDERLINES.has(props.underline)) this.add('underline-styles');
        const fill = props.fill.type === 'group' ? this.group : props.fill;
        if (fill?.type === 'linear' || fill?.type === 'path') this.add('gradient-text');
        this.fill(fill, false);
        if (props.sym || usesScriptFont(run.text)) this.add('script-fonts');
      }
    }
  }

  item(item: LayoutItem): void {
    const el = item.element;
    if (el.placeholder && CONTENT_PLACEHOLDERS.has(el.placeholder.type))
      this.add('content-placeholders');
    switch (item.kind) {
      case 'group': {
        // A group with `a:grpFill` passes its own group's fill on (ECMA-376 Part 1, 20.1.8.35).
        const outer = this.group;
        if (item.element.fill.type !== 'group') this.group = item.element.fill;
        for (const c of item.children) this.item(c);
        this.group = outer;
        return;
      }
      case 'frame': {
        if (item.poster) this.item(item.poster);
        if (item.drawing) this.add('diagram-drawing');
        for (const c of item.drawing ?? []) this.item(c);
        // A chart's labels set text like any shape, so their approximations count.
        if (item.chart) for (const c of item.chart) this.item(c);
        const table = item.table;
        if (!table) return;
        this.fill(item.element.table?.background, false);
        for (const cell of table.cells) {
          const c = cell.cell;
          this.text(c.text);
          this.fill(c.fill, false);
          if (ROTATED_TEXT.has(c.text.body.vert) && c.text.paragraphs.some((p) => p.runs.length))
            this.add('rotated-cell-text');
          const b = c.borders;
          for (const l of [b.left, b.right, b.top, b.bottom, b.tlToBr, b.blToTr])
            this.fill(l?.fill, false);
          if (b.left || b.right || b.top || b.bottom || b.tlToBr || b.blToTr)
            this.add('coincident-borders');
        }
        return;
      }
      case 'shape':
      case 'picture': {
        const rotated = el.xfrm.rot % 360 !== 0;
        this.effects(item.element.effects, item.geometry.paths.length <= MAX_OVERLAY_PATHS);
        this.fill(item.element.fill, rotated);
        this.line(item.element.line, item.kind === 'picture');
        for (const p of item.geometry.paths)
          if (p.fill !== 'norm' && p.fill !== 'none') this.add('path-fill-modes');
        if (item.kind === 'shape' && item.text && hasCharacters(item.text)) {
          this.text(item.element.text);
          if (item.textMatrix && flipped(item.matrix)) this.add('flipped-text');
        }
        return;
      }
    }
  }
}

const flipped = (m: Matrix): boolean => m[0] * m[3] - m[1] * m[2] < 0;

/** The most paths of a shape that paint draws a fill overlay on (`output/paint.ts`, test-checked). */
export const MAX_OVERLAY_PATHS = 64;

/** A gradient with two stops at 0 and 100 %: its colors bend (`paint-gradient.ts`). */
const twoStops = (fill: ResolvedFill): boolean =>
  (fill.type === 'linear' || fill.type === 'path') &&
  fill.stops.length === 2 &&
  fill.stops[0].pos <= 0 &&
  fill.stops[1].pos >= 1;

/**
 * A rectangle or shape path gradient whose focus lies outside the fill box: not measured. (A circle
 * gradient with an outside focus is measured: `fixtures/fill-fidelity.pdf`, slide 2.)
 */
function focusOutside(f: RelRect): boolean {
  const fx = (f.l + 1 - f.r) / 2;
  const fy = (f.t + 1 - f.b) / 2;
  return fx < 0 || fx > 1 || fy < 0 || fy > 1;
}

/**
 * A gradient tile rectangle that paint does not draw as measured on reference PDFs: one smaller
 * than the box (reference PDFs show it tiled), or one on a linear, rectangle, or shape gradient. A tile rectangle that
 * extends the box of a circle gradient is measured (`fixtures/fill-fidelity.pdf`, slide 1).
 */
function unmeasuredTile(fill: Extract<ResolvedFill, { type: 'linear' | 'path' }>): boolean {
  const t = fill.tileRect;
  if (!t || (t.l === 0 && t.t === 0 && t.r === 0 && t.b === 0)) return false;
  const inset = t.l > 0 || t.t > 0 || t.r > 0 || t.b > 0;
  return inset || fill.type === 'linear' || fill.path !== 'circle';
}

/** The ids of the approximations a laid-out slide's items and background use, in catalog order. */
export function slideApproximations(
  items: readonly LayoutItem[],
  background?: ResolvedFill
): string[] {
  const c = new Collector();
  c.fill(background, false);
  for (const item of items) c.item(item);
  return APPROXIMATIONS.filter((a) => c.found.has(a.id)).map((a) => a.id);
}
