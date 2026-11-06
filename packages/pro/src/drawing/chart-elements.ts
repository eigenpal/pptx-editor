/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The output step of a chart: its painted scene and its labels as resolved shapes in the
// frame's box. The engine lays these out with the same geometry and text pipeline as every
// other shape, so the editor and the PDF writer draw one layout. This module replaces DOCX
// Editor's `chart-graphic.ts`, which emits that engine's vector components and text stories.

import type {
  BodyProps,
  ResolvedFill,
  ResolvedGraphicFrame,
  ResolvedLine,
  ResolvedParagraph,
  ResolvedShape,
  RunProps,
  Theme,
} from '@pptx-editor.dev/core/cascade';
import type { GeometryPath, PathCommand } from '@pptx-editor.dev/core/layout';
import type { ChartComponent } from './chart-scene.ts';
import type { ChartPaint, ChartStroke } from './chart-style.ts';
import {
  chartTypeface,
  MAX_LABEL_UNITS,
  splitLines,
  type ChartLabel,
  type ChartTextRun,
} from './chart-text.ts';

const EMU_PER_POINT = 12_700;
/** Labels one chart may set; past it the chart does not render. */
const MAX_LABELS = 4_096;
/**
 * Room added to a label frame on the sides its alignment leaves free, in points, so a line
 * that fills its frame never wraps.
 */
const LABEL_SLACK = 2;

const emu = (pt: number): number => Math.round(pt * EMU_PER_POINT);

function rgba(paint: ChartPaint): { r: number; g: number; b: number; a: number } {
  const value = /^[0-9a-f]{6}$/i.test(paint.hex) ? Number.parseInt(paint.hex, 16) : 0;
  const a = Number.isFinite(paint.alpha) ? Math.min(1, Math.max(0, paint.alpha)) : 1;
  return { r: (value >> 16) & 255, g: (value >> 8) & 255, b: value & 255, a };
}

function solid(paint: ChartPaint): ResolvedFill {
  return { type: 'solid', color: rgba(paint) };
}

function line(stroke: ChartStroke, join: ChartComponent['join']): ResolvedLine {
  return {
    fill: solid(stroke),
    width: emu(stroke.widthPt),
    cap: stroke.cap === 'round' ? 'rnd' : stroke.cap === 'square' ? 'sq' : 'flat',
    join: join ?? 'miter',
    miterLimit: 8,
    // The scene splits open dashed lines into dashes; a closed outline keeps its pattern.
    dash: stroke.dash ? 'custom' : 'solid',
    ...(stroke.dash ? { dashes: stroke.dash } : {}),
    compound: 'sng',
  };
}

const usablePaint = (paint: ChartPaint | null): paint is ChartPaint =>
  paint !== null && /^[0-9a-f]{6}$/i.test(paint.hex);

function geometryPath(component: ChartComponent, fill: boolean, stroke: boolean): GeometryPath {
  const commands: PathCommand[] = [];
  for (const points of component.subpaths) {
    for (const [index, point] of points.entries())
      commands.push({ op: index === 0 ? 'M' : 'L', x: emu(point.x), y: emu(point.y) });
    if (component.closed) commands.push({ op: 'Z' });
  }
  return { commands, fill: fill ? 'norm' : 'none', stroke };
}

interface Owner {
  readonly frame: ResolvedGraphicFrame;
  readonly theme: Theme;
}

function base(owner: Owner, index: number, kind: string) {
  const frame = owner.frame;
  return {
    nodeId: `${frame.nodeId}#chart-${kind}:${index}`,
    partName: frame.partName,
    layer: frame.layer,
    name: '',
    hidden: false,
  } as const;
}

/** One scene component as a shape over the whole chart box, or null when it paints nothing. */
function componentShape(
  component: ChartComponent,
  index: number,
  owner: Owner
): ResolvedShape | null {
  const fill = component.closed && usablePaint(component.fill) ? component.fill : null;
  const stroke = component.stroke && usablePaint(component.stroke) ? component.stroke : null;
  if (!fill && !stroke) return null;
  const { cx, cy } = owner.frame.xfrm;
  return {
    ...base(owner, index, 'shape'),
    kind: 'shape',
    xfrm: { x: 0, y: 0, cx, cy, rot: 0, flipH: false, flipV: false },
    geometry: { kind: 'paths', paths: [geometryPath(component, !!fill, !!stroke)] },
    fill: fill ? solid(fill) : { type: 'none' },
    ...(stroke ? { line: line(stroke, component.join) } : {}),
  };
}

const LABEL_BODY: BodyProps = {
  lIns: 0,
  tIns: 0,
  rIns: 0,
  bIns: 0,
  anchor: 'ctr',
  anchorCtr: false,
  wrap: 'square',
  vert: 'horz',
  rot: 0,
  upright: false,
  numCol: 1,
  spcCol: 0,
  autofit: { kind: 'none' },
  spcFirstLastPara: false,
};

function runProps(run: ChartTextRun, theme: Theme): RunProps {
  const style = run.style;
  const latinTheme = style.latin.startsWith('+mj')
    ? 'major'
    : style.latin.startsWith('+mn')
      ? 'minor'
      : undefined;
  const latin = chartTypeface(style.latin, theme);
  return {
    size: Math.round(style.sizePt * 100),
    bold: style.bold,
    italic: style.italic,
    underline: style.underline ? 'sng' : 'none',
    strike: 'noStrike',
    caps: 'none',
    spacing: 0,
    baseline: 0,
    latin,
    ...(latinTheme ? { latinTheme } : {}),
    ea: style.eastAsian ? chartTypeface(style.eastAsian, theme) : latin,
    cs: style.complex ? chartTypeface(style.complex, theme) : latin,
    fill: solid({ hex: style.colorHex, alpha: 1 }),
  };
}

function paragraph(
  runs: readonly ChartTextRun[],
  align: ChartLabel['align'],
  id: string,
  theme: Theme
): ResolvedParagraph {
  const usable = runs.filter((run) => run.text !== '');
  // The paragraph mark takes the smallest run's size, so it does not raise the line.
  const smallest = usable.reduce((min, run) => (run.style.sizePt < min.style.sizePt ? run : min));
  return {
    nodeId: `${id}:p`,
    level: 0,
    align,
    marL: 0,
    marR: 0,
    indent: 0,
    lineSpacing: { pct: 1 },
    spaceBefore: { pts: 0 },
    spaceAfter: { pts: 0 },
    bullet: null,
    defaultTab: 914_400,
    tabs: [],
    rtl: false,
    eaLineBreak: true,
    latinLineBreak: false,
    hangingPunct: false,
    runs: usable.map((run, index) => ({
      kind: 'text',
      nodeId: `${id}:r${index}`,
      text: run.text.length > MAX_LABEL_UNITS ? run.text.slice(0, MAX_LABEL_UNITS) : run.text,
      props: runProps(run, theme),
    })),
    endProps: runProps(smallest, theme),
  };
}

/** One label as a text shape, or null when it has no text. */
function labelShape(label: ChartLabel, index: number, owner: Owner): ResolvedShape | null {
  const id = `${owner.frame.nodeId}#chart-label:${index}`;
  const paragraphs: ResolvedParagraph[] = [];
  for (const [line, runs] of splitLines(label.paragraphs).entries()) {
    if (!runs.some((run) => run.text !== '')) continue;
    paragraphs.push(paragraph(runs, label.align, `${id}:${line}`, owner.theme));
  }
  if (paragraphs.length === 0) return null;
  const left = label.align === 'l' ? 0 : LABEL_SLACK;
  const right = label.align === 'r' ? 0 : LABEL_SLACK;
  const rotation = ((label.rotation % 360) + 360) % 360;
  return {
    ...base(owner, index, 'label'),
    kind: 'shape',
    xfrm: {
      x: emu(label.x - left),
      y: emu(label.y),
      cx: Math.max(1, emu(label.width + left + right)),
      cy: Math.max(1, emu(label.height)),
      rot: Number.isFinite(rotation) ? rotation : 0,
      flipH: false,
      flipV: false,
    },
    geometry: { kind: 'preset', preset: 'rect', adjust: new Map() },
    fill: { type: 'none' },
    text: { body: LABEL_BODY, paragraphs },
  };
}

/**
 * The shapes of a painted chart: every scene component, then every label, in the frame's box.
 * Null past the label bound.
 */
export function chartElements(
  components: readonly ChartComponent[],
  labels: readonly ChartLabel[],
  frame: ResolvedGraphicFrame,
  theme: Theme
): ResolvedShape[] | null {
  if (labels.length > MAX_LABELS) return null;
  const owner: Owner = { frame, theme };
  const out: ResolvedShape[] = [];
  for (const [index, component] of components.entries()) {
    const shape = componentShape(component, index, owner);
    if (shape) out.push(shape);
  }
  for (const [index, label] of labels.entries()) {
    const shape = labelShape(label, index, owner);
    if (shape) out.push(shape);
  }
  return out;
}
