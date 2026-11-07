// The link of a run as the link popup reads it: a slide or a custom show in its `#` form, a file
// in its `./` form, or the external target as the file stores it. `setLink` takes the same form
// back, so Edit and Apply keep the link's kind.

import { customShowLinkOf, fileLinkUrl } from '../automation/commands/custom-show-links.ts';
import { slideLinkOf } from '../automation/commands/slide-links.ts';
import { relationshipById, type PptxPackage } from '../store/index.ts';

/** The action of a link to a program: kept as the file has it, never written by the editor. */
const PROGRAM_ACTION = 'ppaction://program';

/**
 * The link of a run of slide part `part` with relationship `rId` and `action`, or undefined.
 * `slideParts` are the deck's slide parts in order.
 */
export function runLinkOf(
  pkg: PptxPackage | undefined,
  part: string | undefined,
  slideParts: readonly string[],
  rId: string | undefined,
  action: string | undefined
): string | undefined {
  const rel = pkg && part && rId ? relationshipById(pkg, part, rId) : undefined;
  const target = rel && !rel.external ? rel.target : undefined;
  const local = action
    ? (slideLinkOf(action, target, slideParts) ?? customShowLinkOf(action))
    : undefined;
  if (local) return local;
  if (!rel?.external) return undefined;
  // A program link shows its target as the file stores it, which the popup never opens.
  if (action?.startsWith(PROGRAM_ACTION)) return rel.externalTarget ?? rel.relativeTarget;
  return rel.externalTarget ?? fileLinkUrl(rel.relativeTarget);
}
