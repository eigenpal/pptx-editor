// The insert-shape gallery: preset shapes by category, and previews drawn by the same geometry
// engine that paints slides, so a gallery icon is the shape the user gets.

import type { TranslationKey } from '@pptx-editor.dev/i18n';
import { evaluateGeometry } from '../layout/index.ts';
import { pathData } from '../output/index.ts';
import { isInsertablePreset } from '../automation/commands/shape-presets.ts';

/** Presets with no inside to show a picture in: lines and connectors (`cropToShape` refuses them). @internal */
export const NO_INSIDE_PRESET =
  /^(line|lineInv|straightConnector1|bentConnector[2-5]|curvedConnector[2-5])$/;

/** One gallery category. @public */
export interface ShapeCategory {
  readonly id: string;
  /** i18n key of the category name. */
  readonly labelKey: TranslationKey;
  /** Preset names (`a:prstGeom/@prst`); each has the i18n label `shapes.<preset>`. */
  readonly presets: readonly string[];
}

/** The gallery, in menu order. @public */
export const SHAPE_GALLERY: readonly ShapeCategory[] = [
  {
    id: 'basic',
    labelKey: 'shapeGallery.basic',
    presets: [
      'rect',
      'roundRect',
      'snip1Rect',
      'snip2SameRect',
      'snip2DiagRect',
      'snipRoundRect',
      'round1Rect',
      'round2SameRect',
      'round2DiagRect',
      'ellipse',
      'triangle',
      'rtTriangle',
      'parallelogram',
      'trapezoid',
      'diamond',
      'pentagon',
      'hexagon',
      'heptagon',
      'octagon',
      'decagon',
      'dodecagon',
      'pie',
      'chord',
      'teardrop',
      'frame',
      'halfFrame',
      'corner',
      'diagStripe',
      'plus',
      'plaque',
      'can',
      'cube',
      'bevel',
      'donut',
      'noSmoking',
      'blockArc',
      'foldedCorner',
      'smileyFace',
      'heart',
      'lightningBolt',
      'sun',
      'moon',
      'cloud',
      'arc',
      'bracketPair',
      'bracePair',
      'leftBracket',
      'rightBracket',
      'leftBrace',
      'rightBrace',
    ],
  },
  {
    id: 'arrows',
    labelKey: 'shapeGallery.arrows',
    presets: [
      'rightArrow',
      'leftArrow',
      'upArrow',
      'downArrow',
      'leftRightArrow',
      'upDownArrow',
      'quadArrow',
      'leftRightUpArrow',
      'bentArrow',
      'uturnArrow',
      'leftUpArrow',
      'bentUpArrow',
      'curvedRightArrow',
      'curvedLeftArrow',
      'curvedUpArrow',
      'curvedDownArrow',
      'stripedRightArrow',
      'notchedRightArrow',
      'homePlate',
      'chevron',
      'rightArrowCallout',
      'leftArrowCallout',
      'upArrowCallout',
      'downArrowCallout',
      'leftRightArrowCallout',
      'quadArrowCallout',
      'circularArrow',
    ],
  },
  {
    id: 'callouts',
    labelKey: 'shapeGallery.callouts',
    presets: [
      'wedgeRectCallout',
      'wedgeRoundRectCallout',
      'wedgeEllipseCallout',
      'cloudCallout',
      'borderCallout1',
      'borderCallout2',
      'borderCallout3',
    ],
  },
  {
    id: 'equation',
    labelKey: 'shapeGallery.equation',
    presets: ['mathPlus', 'mathMinus', 'mathMultiply', 'mathDivide', 'mathEqual', 'mathNotEqual'],
  },
  {
    id: 'flowchart',
    labelKey: 'shapeGallery.flowchart',
    presets: [
      'flowChartProcess',
      'flowChartAlternateProcess',
      'flowChartDecision',
      'flowChartInputOutput',
      'flowChartPredefinedProcess',
      'flowChartInternalStorage',
      'flowChartDocument',
      'flowChartMultidocument',
      'flowChartTerminator',
      'flowChartPreparation',
      'flowChartManualInput',
      'flowChartManualOperation',
      'flowChartConnector',
      'flowChartOffpageConnector',
      'flowChartPunchedCard',
      'flowChartPunchedTape',
      'flowChartSummingJunction',
      'flowChartOr',
      'flowChartCollate',
      'flowChartSort',
      'flowChartExtract',
      'flowChartMerge',
      'flowChartOnlineStorage',
      'flowChartDelay',
      'flowChartMagneticTape',
      'flowChartMagneticDisk',
      'flowChartMagneticDrum',
      'flowChartDisplay',
    ],
  },
  {
    id: 'stars',
    labelKey: 'shapeGallery.stars',
    presets: [
      'irregularSeal1',
      'irregularSeal2',
      'star4',
      'star5',
      'star6',
      'star7',
      'star8',
      'star10',
      'star12',
      'star16',
      'star24',
      'star32',
      'ribbon2',
      'ribbon',
      'ellipseRibbon2',
      'ellipseRibbon',
      'verticalScroll',
      'horizontalScroll',
      'wave',
      'doubleWave',
    ],
  },
  { id: 'lines', labelKey: 'shapeGallery.lines', presets: ['line'] },
];

/**
 * The Crop to Shape gallery: the shape gallery without the presets that have no inside to show
 * a picture in (lines and connectors), by the rule `cropToShape` applies. @public
 */
export const CROP_SHAPE_GALLERY: readonly ShapeCategory[] = SHAPE_GALLERY.map((c) => ({
  ...c,
  presets: c.presets.filter((p) => !NO_INSIDE_PRESET.test(p)),
})).filter((c) => c.presets.length > 0);

/** A gallery preview: SVG path data in a 0…100 box. @public */
export interface ShapePreview {
  readonly viewBox: string;
  readonly paths: readonly {
    readonly d: string;
    readonly fill: boolean;
    readonly stroke: boolean;
  }[];
}

const previews = new Map<string, ShapePreview>();

/** Wide presets preview at 2:1, tall ones at 1:2; the rest are square. */
const WIDE =
  /^(right|left|leftRight|striped|notched|bent|uturn|curvedUp|curvedDown|quad)?Arrow|Arrow$|ArrowCallout$|homePlate|chevron|ibbon|horizontalScroll|wave|Wave|mathEqual|mathNotEqual|mathMinus|flowChartTerminator|flowChartPunchedTape|flowChartMagneticDrum|flowChartDisplay|flowChartDelay|flowChartAlternateProcess|flowChartProcess|flowChartPredefinedProcess|flowChartDocument|flowChartMultidocument|flowChartInputOutput|flowChartPreparation|flowChartManualInput|flowChartManualOperation|flowChartOnlineStorage/;
const SQUARE =
  /^(quadArrow|leftRightUpArrow|leftUpArrow|bentUpArrow|uturnArrow|quadArrowCallout|circularArrow)$/;
const TALL =
  /^(upArrow|downArrow|upDownArrow|upArrowCallout|downArrowCallout|upDownArrowCallout|curvedLeftArrow|curvedRightArrow|verticalScroll|can|leftBracket|rightBracket|leftBrace|rightBrace)$/;

/** The preview of a preset, from its spec geometry. Null for an unknown preset. @public */
export function shapePreview(preset: string): ShapePreview | null {
  const hit = previews.get(preset);
  if (hit) return hit;
  if (!isInsertablePreset(preset)) return null;
  const tall = TALL.test(preset);
  const w = tall ? 50 : 100;
  const h = !tall && !SQUARE.test(preset) && WIDE.test(preset) ? 50 : 100;
  const g = evaluateGeometry({ kind: 'preset', preset, adjust: new Map() }, w, h);
  // Frame every point the paths reach (callout leader lines leave the box), square and padded.
  let x0 = 0;
  let y0 = 0;
  let x1 = w;
  let y1 = h;
  // Curves are sampled: their control points can lie far outside the drawn shape.
  const visit = (x: number, y: number): void => {
    x0 = Math.min(x0, x);
    x1 = Math.max(x1, x);
    y0 = Math.min(y0, y);
    y1 = Math.max(y1, y);
  };
  for (const path of g.paths) {
    if (!path.stroke) continue;
    let px = 0;
    let py = 0;
    for (const c of path.commands) {
      if (c.op === 'Z') continue;
      if (c.op === 'C' || c.op === 'Q')
        for (let i = 1; i < 8; i += 1) {
          const t = i / 8;
          const u = 1 - t;
          if (c.op === 'C')
            visit(
              u * u * u * px + 3 * u * u * t * c.x1 + 3 * u * t * t * c.x2 + t * t * t * c.x,
              u * u * u * py + 3 * u * u * t * c.y1 + 3 * u * t * t * c.y2 + t * t * t * c.y
            );
          else
            visit(
              u * u * px + 2 * u * t * c.x1 + t * t * c.x,
              u * u * py + 2 * u * t * c.y1 + t * t * c.y
            );
        }
      visit(c.x, c.y);
      px = c.x;
      py = c.y;
    }
  }
  const size = Math.max(x1 - x0, y1 - y0) * 1.08;
  const cx = (x0 + x1) / 2;
  const cy = (y0 + y1) / 2;
  const preview: ShapePreview = {
    viewBox: [cx - size / 2, cy - size / 2, size, size]
      .map((v) => Math.round(v * 100) / 100)
      .join(' '),
    // A fill-only path (no outline) is a background region, such as the inside of a bracket;
    // gallery icons draw outlined paths only.
    paths: g.paths
      .filter((p) => p.stroke)
      .map((p) => ({ d: pathData(p), fill: p.fill !== 'none', stroke: true }))
      .filter((p) => p.d),
  };
  previews.set(preset, preview);
  return preview;
}

export { isInsertablePreset };
