// Whether a slide shape's text fits its text area (`readTextFit`, the object model's
// `TextFrame.getTextFit()`). The answer comes from the slide layout that the shape's picture
// draws (`layoutSlide`, as `DocumentSession.layout`), measured with the host's measurer, so it
// needs no canvas: a server host measures with the measurer it was given.
//
// The text area is the text rectangle less the insets of `a:bodyPr` (CT_TextBodyProperties,
// ECMA-376 Part 1, 21.1.2.1.1). The text lays out at its stored autofit (`textFitOf`). A text
// write re-fits the body first (`refitText`), so after one `a:spAutoFit` (CT_TextShapeAutofit,
// 21.1.2.1.4) does not overflow, and `a:normAutofit` (CT_TextNormalAutofit, 21.1.2.1.3)
// overflows only when its last shrink step does not fit.
//
// Refused: a shape with no text frame (a picture, a connector, a group, a table: the table
// layout grows each row to its cells' text, `table-layout.ts`), a shape of a layout or a master
// (it has no slide to lay out on), a hidden shape (the slide does not draw it), and a host with
// no measurer or with the fixed measurer, whose heights model no font.

import { isFixedMeasurer, layoutSlide, type LayoutItem } from '../layout/index.ts';
import { textFitOf } from '../layout/autofit.ts';
import type { AutomationRef, AutomationValue } from './protocol.ts';
import { resolvedSlide } from './plan-reads.ts';
import { readTarget, type TextPlanContext } from './plan-text.ts';
import { AutomationFailure, slideOf, toPoints } from './targets.ts';

/** The laid-out item of the shape `nodeId` of `part`, through its groups. */
function itemOf(
  items: readonly LayoutItem[],
  part: string,
  nodeId: string,
  depth = 0
): LayoutItem | undefined {
  if (depth > 32) return undefined;
  for (const item of items) {
    if (item.element.nodeId === nodeId && item.element.partName === part) return item;
    if (item.kind !== 'group') continue;
    const inner = itemOf(item.children, part, nodeId, depth + 1);
    if (inner) return inner;
  }
  return undefined;
}

export function readTextFit(context: TextPlanContext, ref: AutomationRef): AutomationValue {
  const measurer = context.measurer;
  if (!measurer || isFixedMeasurer(measurer))
    throw new AutomationFailure(
      'unsupported-capability',
      'this host has no font measurement',
      'text-fit-measurer'
    );
  // A `p:sp` without a text body reads as the body a write would add, as `readTextFrame` does.
  const { t, ctx } = readTarget(context, ref, undefined, undefined)!;
  const { owner, part, el } = t.at;
  if (owner.kind !== 'slide')
    throw new AutomationFailure(
      'unsupported-content',
      'only a shape on a slide has a text layout',
      'text-fit-owner'
    );
  const pkg = ctx.batch.pkg;
  const slide = resolvedSlide(pkg, slideOf(pkg, owner.sldId).index);
  const item = slide && itemOf(layoutSlide(slide, measurer).items, part, el.id);
  if (!item)
    throw new AutomationFailure(
      'unsupported-content',
      'the slide does not draw that shape',
      'text-fit-hidden'
    );
  const fit = item.kind === 'shape' ? textFitOf(item, measurer) : null;
  if (!fit)
    throw new AutomationFailure('unsupported-content', 'that shape has no text frame', 'no-text');
  return {
    kind: 'textFit',
    textFit: {
      overflows: fit.overflows,
      textHeight: toPoints(fit.textHeight),
      availableHeight: toPoints(fit.availableHeight),
      substitutedFonts: [...fit.substitutions],
    },
  };
}
