// The selected picture and the alt text of the selection, as the snapshot reports them for the
// Picture Format tab and the Alt Text dialog. Both read the slide's XML, so the chrome shows the
// values the file holds (`a:srcRect`, the blip effects, `p:cNvPr`).

import {
  attr,
  child,
  nodeIndex,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../store/index.ts';
import type {
  AltText,
  EditorSelection,
  SelectedImageState,
  SelectionGeometry,
} from '../contracts/index.ts';
import type { SlideLayout } from '../layout/index.ts';
import { findItem } from '../automation/commands/placement.ts';
import {
  adjustmentsOf,
  cropOf,
  isDecorative,
  nonVisualOf,
  pictureLock,
  toPercent,
  usableCrop,
} from './picture-commands.ts';
import { pictureStyleOf } from './picture-presets.ts';
import { chartStateAt } from './chart-commands.ts';
import { equationStateAt } from './equation-state.ts';
import { diagramStateAt } from './diagram-commands.ts';

export { chartDataAt, chartReads } from './chart-commands.ts';
export { afterInsertCommand } from './pane-requests.ts';

/** The selected picture, chart, and equation, as the snapshot reports them. */
export function selectedObjectsAt(
  pkg: PptxPackage | undefined,
  part: string | undefined,
  selection: EditorSelection,
  cropping: string | null,
  layout?: SlideLayout
): {
  image: SelectedImageState | null;
  chart: ReturnType<typeof chartStateAt>;
  equation: ReturnType<typeof equationStateAt>;
  diagram: ReturnType<typeof diagramStateAt>;
} {
  return {
    image: pictureStateAt(pkg, part, selection, cropping, layout),
    chart: chartStateAt(pkg, part, selection),
    equation: equationStateAt(pkg, part, selection),
    diagram: diagramStateAt(pkg, part, selection),
  };
}

function elementAt(
  pkg: PptxPackage | undefined,
  part: string | undefined,
  id: string
): Element | undefined {
  const xml = pkg && part ? xmlPartOf(pkg, part) : undefined;
  const node = xml ? nodeIndex(xml).get(id)?.node : undefined;
  return node?.type === 'element' ? node : undefined;
}

/**
 * The selected pictures, while every selected element is a picture; else null. The values are
 * the first picture's. The shape comes from the laid-out geometry, so a placeholder picture
 * reports the shape it inherits.
 */
export function pictureStateAt(
  pkg: PptxPackage | undefined,
  part: string | undefined,
  selection: EditorSelection,
  cropping: string | null,
  layout?: SlideLayout
): SelectedImageState | null {
  if (selection.kind !== 'shapes' || selection.nodeIds.length === 0) return null;
  const pics = selection.nodeIds.map((id) => elementAt(pkg, part, id));
  if (pics.some((p) => !p || p.kind !== 'pic')) return null;
  const pic = pics[0]!;
  const blipFill = child(pic, 'p:blipFill');
  const spPr = child(pic, 'p:spPr');
  const preset = child(spPr, 'a:prstGeom');
  const style = pictureStyleOf(spPr);
  const laid = layout ? findItem(layout.items, pic.id) : undefined;
  const geometry = laid?.kind === 'picture' ? laid.element.geometry : undefined;
  // The box on the slide, as the Size fields show it and `setImageProperties` takes it.
  const box = laid?.element.xfrm;
  return {
    nodeId: pic.id,
    count: pics.length,
    widthEmu: box?.cx ?? 0,
    heightEmu: box?.cy ?? 0,
    rotationDegrees: box?.rot ?? 0,
    // As paint and the commands read it: a merged axis that shows nothing reads as uncropped.
    crop: toPercent(usableCrop(cropOf(blipFill))),
    shape: geometry
      ? geometry.kind === 'preset'
        ? geometry.preset
        : 'custom'
      : preset
        ? (attr(preset, 'prst') ?? 'rect')
        : child(spPr, 'a:custGeom')
          ? 'custom'
          : 'rect',
    cropping: pics.length === 1 && cropping === pic.id,
    ...adjustmentsOf(child(blipFill, 'a:blip')),
    ...(style ? { style } : {}),
    lockAspect: pictureLock(pic, 'noChangeAspect'),
    canCrop: !pics.some((p) => pictureLock(p!, 'noCrop')),
    canResize: !pics.some((p) => pictureLock(p!, 'noResize')),
    canRotate: !pics.some((p) => pictureLock(p!, 'noRot')),
  };
}

/** The alt text of an element (`p:cNvPr/@descr`, `@title`, and `adec:decorative`), if it has a `p:cNvPr`. */
export function altTextAt(
  pkg: PptxPackage | undefined,
  part: string | undefined,
  id: string
): AltText | undefined {
  const el = elementAt(pkg, part, id);
  const cNvPr = el ? nonVisualOf(el) : undefined;
  if (!cNvPr) return undefined;
  return {
    description: attr(cNvPr, 'descr') ?? '',
    title: attr(cNvPr, 'title') ?? '',
    decorative: isDecorative(cNvPr),
  };
}

/** The selection geometry with the element's alt text (`p:cNvPr/@descr`, `@title`). */
export function withAltText(
  geometry: SelectionGeometry | null,
  pkg: PptxPackage | undefined,
  part: string | undefined
): SelectionGeometry | null {
  if (!geometry) return null;
  const altText = altTextAt(pkg, part, geometry.nodeId);
  return altText ? { ...geometry, altText } : geometry;
}
