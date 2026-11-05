// Office.js `Slide.getImageAsBase64` and `Shape.getImageAsBase64`: the check in the batch, and
// what the host's `renderImage` draws.
//
// A picture is drawn after the batch, because drawing waits for fonts and for the canvas. The
// batch only checks the target and the size, and answers the handle that `renderImage` takes.
// So a wrong target or size refuses in the batch, before anything is drawn.

import { shapeTreeOf, type PptxPackage } from '../store/index.ts';
import type { AutomationTarget } from './handles.ts';
import type { AutomationRef, AutomationValue } from './protocol.ts';
import {
  AutomationFailure,
  findShape,
  handleOf,
  slideOf,
  targetOf,
  type PlanContext,
} from './targets.ts';

/** What the port draws (`AutomationDocumentPort.renderImage`): a slide, or one shape of it. */
export interface AutomationImageTarget {
  /** The index of the slide in the presentation. */
  readonly slideIndex: number;
  /** The tree node id of the shape; absent for the whole slide. */
  readonly shapeNodeId?: string;
  /** The width of the box the picture fits into, in pixels, keeping its aspect ratio. */
  readonly width?: number;
  /** The height of that box. */
  readonly height?: number;
}

/** A size of a picture: absent, or a positive finite number of pixels. */
export function imageSize(value: unknown, name: 'width' | 'height'): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0)
    throw new AutomationFailure('invalid-argument', `invalid ${name}`, 'image-size');
  return value;
}

/**
 * The slide index and shape node id of a target, in `pkg`. A picture is of a slide or of a shape
 * on a slide; a shape of a layout or a master has no slide to draw it on.
 */
export function imageTarget(
  pkg: PptxPackage,
  target: AutomationTarget,
  width: unknown,
  height: unknown
): AutomationImageTarget {
  const size = {
    ...(imageSize(width, 'width') === undefined ? {} : { width: width as number }),
    ...(imageSize(height, 'height') === undefined ? {} : { height: height as number }),
  };
  if (target.kind === 'slide') return { slideIndex: slideOf(pkg, target.sldId).index, ...size };
  if (target.kind !== 'shape' || target.owner.kind !== 'slide')
    throw new AutomationFailure(
      'unsupported-content',
      'only a slide or a shape on a slide has a picture',
      'image-owner'
    );
  const { ref, index } = slideOf(pkg, target.owner.sldId);
  const hit = findShape(shapeTreeOf(pkg, ref.partName), target.shapeId);
  if (!hit) throw new AutomationFailure('invalid-handle', 'that shape is gone', 'gone');
  return { slideIndex: index, shapeNodeId: hit.el.id, ...size };
}

/** Check a picture request in the batch, and answer the handle `renderImage` takes. */
export function prepareImage(
  ctx: PlanContext,
  ref: AutomationRef,
  width: unknown,
  height: unknown
): AutomationValue {
  const target = targetOf(ctx, ref, 'slide', 'shape', 'slideLayout', 'slideMaster');
  imageTarget(ctx.batch.pkg, target, width, height);
  return { kind: 'handle', handle: handleOf(ctx, ref) };
}
