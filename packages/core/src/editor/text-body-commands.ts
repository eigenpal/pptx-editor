// Text body settings of whole shapes: Home > Text Direction and Home > Columns, as store ops.
//
// ECMA-376 Part 1, DrawingML `CT_TextBodyProperties` (`a:bodyPr`): `vert` (`ST_TextVerticalType`),
// `numCol` (`ST_TextColumnCount`, 1 to 16), and `spcCol` (`ST_PositiveCoordinate32`, EMU).
// The API names a direction in Office.js words (`ShapeTextOrientation`) and maps
// it to the file token here. The editor sets only the directions that layout draws;
// `mongolianVertical` and `wordArtVerticalRtl` are refused with a typed reason
// (an approximation on the public fidelity page, "Other vertical text types").

import { child } from '../store/index.ts';
import type { TextOrientation } from '../contracts/index.ts';
import { CommandRefusal, type OpBatch } from '../automation/commands/op-batch.ts';
import { tblOf } from '../automation/commands/table-model.ts';

/** Each direction and its `ST_TextVerticalType` token. */
const TOKENS: Readonly<Record<TextOrientation, string>> = {
  horizontal: 'horz',
  vertical: 'vert',
  vertical270: 'vert270',
  wordArtVertical: 'wordArtVert',
  eastAsianVertical: 'eaVert',
  mongolianVertical: 'mongolianVert',
  wordArtVerticalRtl: 'wordArtVertRtl',
};
const BY_TOKEN: ReadonlyMap<string, TextOrientation> = new Map(
  (Object.entries(TOKENS) as [TextOrientation, string][]).map(([d, token]) => [token, d])
);

/** The directions the editor sets, in Text Direction menu order. */
export const TEXT_DIRECTIONS = [
  'horizontal',
  'vertical',
  'vertical270',
  'wordArtVertical',
  'eastAsianVertical',
] as const satisfies readonly TextOrientation[];
const SETTABLE: ReadonlySet<string> = new Set<string>(TEXT_DIRECTIONS);

/**
 * The direction of a file token. A token outside `ST_TextVerticalType` reads as horizontal,
 * as layout draws it.
 */
export function textDirectionOf(token: string): TextOrientation {
  return BY_TOKEN.get(token) ?? 'horizontal';
}

/**
 * The widest gap between columns the editor writes, EMU: 56 in, the cascade's bound
 * (`readBodyProps`). The schema allows more (`ST_PositiveCoordinate32`). @internal
 */
export const MAX_COLUMN_GAP = 51_206_400;

/** The `a:bodyPr` of each selected shape with text. Refuses when none has text. */
function bodies(batch: OpBatch, part: string, ids: readonly string[]): [string, string][] {
  const out: [string, string][] = [];
  for (const id of ids) {
    const bodyPr = child(child(batch.element(part, id), 'p:txBody'), 'a:bodyPr');
    if (bodyPr) out.push([id, bodyPr.id]);
  }
  // A table's cells have their own text bodies; their direction and columns are not built,
  // and the reason says so, not "select a shape that has text".
  if (out.length === 0)
    throw new CommandRefusal(
      !ids.length
        ? 'disabled.noSelection'
        : ids.some((id) => tblOf(batch.element(part, id)))
          ? 'disabled.tableText'
          : 'disabled.noText'
    );
  return out;
}

function setBodyAttr(batch: OpBatch, part: string, bodyPrId: string, name: string, value: string) {
  batch.push({ op: 'setAttribute', part, node: bodyPrId, name: { ns: '', local: name }, value });
}

/** Set the text direction of shapes. Returns the shapes changed, for autofit. */
export function setTextDirection(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  direction: TextOrientation
): string[] {
  // The detail names the field, as for every other refusal of this module.
  if (!Object.hasOwn(TOKENS, direction) || !SETTABLE.has(direction))
    throw new CommandRefusal('errors.unsupported', 'direction');
  const targets = bodies(batch, part, ids);
  for (const [, bodyPr] of targets) setBodyAttr(batch, part, bodyPr, 'vert', TOKENS[direction]);
  return targets.map(([id]) => id);
}

/** Set the text columns of shapes. Returns the shapes changed, for autofit. */
export function setColumns(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  count: number,
  gapEmu: number | undefined
): string[] {
  if (!Number.isInteger(count) || count < 1 || count > 16)
    throw new CommandRefusal('errors.unsupported', 'count');
  if (
    gapEmu !== undefined &&
    !(typeof gapEmu === 'number' && gapEmu >= 0 && gapEmu <= MAX_COLUMN_GAP)
  )
    throw new CommandRefusal('errors.unsupported', 'gapEmu');
  const targets = bodies(batch, part, ids);
  for (const [, bodyPr] of targets) {
    setBodyAttr(batch, part, bodyPr, 'numCol', String(count));
    if (gapEmu !== undefined)
      setBodyAttr(batch, part, bodyPr, 'spcCol', String(Math.round(gapEmu)));
  }
  return targets.map(([id]) => id);
}
