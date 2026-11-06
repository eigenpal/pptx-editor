// Pattern and picture fills (ECMA-376 Part 1, 20.1.8.47 `a:pattFill` and 20.1.8.14
// `a:blipFill`): the checked XML of each, and the shape fill commands that write them. The slide
// background writes the same XML (`slide-commands.ts`), and the Office.js lane checks its
// patterns and pictures with the same rules.

import { imageSize } from '../../store/index.ts';
import type { PatternFill } from '../../contracts/neutral.ts';
import { CommandRefusal } from './op-batch.ts';

/** `ST_PresetPatternVal` (ECMA-376 Part 1, 20.1.10.51). */
export const PATTERN_PRESETS: ReadonlySet<string> = new Set([
  'pct5',
  'pct10',
  'pct20',
  'pct25',
  'pct30',
  'pct40',
  'pct50',
  'pct60',
  'pct70',
  'pct75',
  'pct80',
  'pct90',
  'horz',
  'vert',
  'ltHorz',
  'ltVert',
  'dkHorz',
  'dkVert',
  'narHorz',
  'narVert',
  'dashHorz',
  'dashVert',
  'cross',
  'dnDiag',
  'upDiag',
  'ltDnDiag',
  'ltUpDiag',
  'dkDnDiag',
  'dkUpDiag',
  'wdDnDiag',
  'wdUpDiag',
  'dashDnDiag',
  'dashUpDiag',
  'diagCross',
  'smCheck',
  'lgCheck',
  'smGrid',
  'lgGrid',
  'dotGrid',
  'smConfetti',
  'lgConfetti',
  'horzBrick',
  'diagBrick',
  'solidDmnd',
  'openDmnd',
  'dotDmnd',
  'plaid',
  'sphere',
  'weave',
  'divot',
  'shingle',
  'wave',
  'trellis',
  'zigZag',
]);

const HEX = /^#[0-9a-f]{6}$/i;

/** The image type of picture bytes, by their first bytes: PNG, JPEG, or GIF. */
export function imageType(bytes: Uint8Array): string | undefined {
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47)
    return 'image/png';
  if (bytes[0] === 0xff && bytes[1] === 0xd8) return 'image/jpeg';
  if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) return 'image/gif';
  return undefined;
}

/**
 * The image type of picture bytes, or a refusal when they are no PNG, JPEG, or GIF with a
 * readable size. The type comes from the bytes, not from the caller's `mime`, which can be
 * wrong (a file name extension). The bytes are never decoded further.
 */
export function pictureType(bytes: unknown): string {
  const type = bytes instanceof Uint8Array ? imageType(bytes) : undefined;
  if (!type || !imageSize(bytes as Uint8Array))
    throw new CommandRefusal('errors.unsupported', 'picture');
  return type;
}

/** `a:pattFill` for a pattern, after checking it. */
export function patternXml(p: PatternFill): string {
  if (!p || !PATTERN_PRESETS.has(p.preset))
    throw new CommandRefusal('errors.unsupported', 'pattern');
  if (!HEX.test(p.foreground ?? '') || !HEX.test(p.background ?? ''))
    throw new CommandRefusal('errors.unsupported', 'pattern color');
  const clr = (hex: string): string => `<a:srgbClr val="${hex.slice(1).toUpperCase()}"/>`;
  return (
    `<a:pattFill prst="${p.preset}"><a:fgClr>${clr(p.foreground)}</a:fgClr>` +
    `<a:bgClr>${clr(p.background)}</a:bgClr></a:pattFill>`
  );
}

/**
 * `a:blipFill` over the image relationship `rId`: stretched over the fill box, or tiled from
 * the top left at the picture's own size (`a:tile`, 20.1.8.58).
 */
export function pictureFillXml(rId: string, tile: boolean): string {
  const mode = tile
    ? '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>'
    : '<a:stretch><a:fillRect/></a:stretch>';
  return `<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${rId}"/><a:srcRect/>${mode}</a:blipFill>`;
}
