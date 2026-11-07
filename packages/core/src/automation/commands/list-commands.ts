// Bulleted and numbered lists: turning a list on or off, and changing the list level.
//
// A deck keeps list formatting in the inheritance chain (master text styles, placeholder
// and shape list styles) and writes paragraph overrides only where a paragraph differs:
//
// - Bullets off: `a:buNone` with `marL="0" indent="0"` when an inherited bullet must be
//   suppressed; otherwise the paragraph's own bullet and hanging indent are simply removed.
// - Bullets on: the paragraph's own `a:buNone` (and the zero margins that came with it, or
//   the margin that Backspace left at the bullet's position) are removed first, so a
//   placeholder takes its master's bullet, font, color, and margins. Only
//   a paragraph that inherits no list gets an explicit `a:buChar` or `a:buAutoNum` with a
//   hanging indent.
// - List level: `lvl` moves between 0 and 8. A paragraph with its own `marL` (an explicit
//   list in a text box) moves that margin by one level step, so the text moves with the level.

import { attr, child, type Element } from '../../store/index.ts';
import { resolveSlide, type ResolvedElement, type ResolvedParagraph } from '../../cascade/index.ts';
import type { OpBatch } from './op-batch.ts';
import { aSpec, ensureParagraphProps, paragraphsOf } from './text-model.ts';

/** The hanging indent and the margin step written for an explicit list. */
const HANG = 342_900;
const LEVEL_STEP = 457_200;
const BULLET_CHILDREN = ['buNone', 'buChar', 'buAutoNum', 'buBlip', 'buFont', 'buFontTx'];
/** Children that come before the bullet in `CT_TextParagraphProperties`. */
const BEFORE_BULLET = [
  'lnSpc',
  'spcBef',
  'spcAft',
  'buClrTx',
  'buClr',
  'buSzTx',
  'buSzPct',
  'buSzPts',
];

export type ListKind = 'none' | 'bullet' | 'number';

/** One paragraph a list command applies to. */
export interface ListTarget {
  readonly shapeId: string;
  readonly txBodyId: string;
  /** Paragraph indices in the text body. */
  readonly paragraphs: readonly number[];
}

function setPPrAttr(
  batch: OpBatch,
  part: string,
  pPrId: string,
  name: string,
  value: string | null
) {
  batch.push({ op: 'setAttribute', part, node: pPrId, name: { ns: '', local: name }, value });
}

const levelOf = (pPr: Element): number => Number(attr(pPr, 'lvl') ?? '0') || 0;

function removeBulletChildren(batch: OpBatch, part: string, pPrId: string): void {
  for (const c of batch.element(part, pPrId).children)
    if (c.type === 'element' && BULLET_CHILDREN.includes(c.local))
      batch.push({ op: 'removeNode', part, node: c.id });
}

/** Where a bullet child goes in `a:pPr`: after spacing, bullet color, and bullet size. */
function bulletIndex(batch: OpBatch, part: string, pPrId: string): number {
  let index = 0;
  batch.element(part, pPrId).children.forEach((c, i) => {
    if (c.type === 'element' && BEFORE_BULLET.includes(c.local)) index = i + 1;
  });
  return index;
}

function insertBullet(batch: OpBatch, part: string, pPrId: string, kind: 'bullet' | 'number') {
  const index = bulletIndex(batch, part, pPrId);
  const specs =
    kind === 'bullet'
      ? [aSpec('buFont', { typeface: 'Arial' }), aSpec('buChar', { char: '•' })]
      : [aSpec('buFont', { typeface: '+mj-lt' }), aSpec('buAutoNum', { type: 'arabicPeriod' })];
  specs.forEach((spec, i) =>
    batch.insert({ op: 'insertNode', part, parent: pPrId, index: index + i, spec })
  );
}

/** The resolved text body `txBodyId` of element `nodeId`: a shape's text or a table cell's. */
function findBody(
  elements: readonly ResolvedElement[],
  nodeId: string,
  txBodyId: string
): readonly ResolvedParagraph[] | undefined {
  for (const el of elements) {
    if (el.nodeId === nodeId && el.kind === 'shape') return el.text?.paragraphs;
    if (el.nodeId === nodeId && el.kind === 'graphicFrame')
      return el.table?.cells.find((c) => c.text.nodeId === txBodyId)?.text.paragraphs;
    if (el.kind === 'group') {
      const hit = findBody(el.children, nodeId, txBodyId);
      if (hit) return hit;
    }
  }
  return undefined;
}

/** The paragraphs of one text body as the cascade resolves the batch's current state. */
export function resolvedParagraphs(
  batch: OpBatch,
  slide: number,
  target: ListTarget
): readonly ResolvedParagraph[] {
  const resolved = resolveSlide(batch.pkg, slide);
  return (resolved && findBody(resolved.elements, target.shapeId, target.txBodyId)) ?? [];
}

/**
 * The paragraphs of each text body of slide `slide`, as the cascade resolves the batch's state
 * now: one resolve for many targets, where `resolvedParagraphs` resolves the slide each time.
 */
export function paragraphLookup(
  batch: OpBatch,
  slide: number
): (target: ListTarget) => readonly ResolvedParagraph[] {
  const shapes = new Map<string, readonly ResolvedParagraph[]>();
  const cells = new Map<string, readonly ResolvedParagraph[]>();
  const walk = (elements: readonly ResolvedElement[]): void => {
    for (const el of elements) {
      if (el.kind === 'shape') shapes.set(el.nodeId, el.text?.paragraphs ?? []);
      else if (el.kind === 'graphicFrame')
        for (const c of el.table?.cells ?? [])
          cells.set(`${el.nodeId}\u0000${c.text.nodeId}`, c.text.paragraphs);
      else if (el.kind === 'group') walk(el.children);
    }
  };
  walk(resolveSlide(batch.pkg, slide)?.elements ?? []);
  return (t) => shapes.get(t.shapeId) ?? cells.get(`${t.shapeId}\u0000${t.txBodyId}`) ?? [];
}

/**
 * Turn a list on or off for the target paragraphs. `resolve` is false for a probe ("can this
 * command run"), which only needs ops, not the inheritance check.
 */
export function setListKind(
  batch: OpBatch,
  part: string,
  slide: number,
  target: ListTarget,
  kind: ListKind,
  resolve = true
): void {
  const all = paragraphsOf(batch.element(part, target.txBodyId));
  const pPrs = target.paragraphs.map((i) => ensureParagraphProps(batch, part, all[i].id));
  // A margin dropped on trust in phase 1: put back unless it was the bullet's position.
  const dropped: (string | undefined)[] = [];
  // Phase 1: drop the paragraph's own list overrides, so inheritance shows through.
  for (const pPrId of pPrs) {
    const pPr = batch.element(part, pPrId);
    const suppressed = !!child(pPr, 'a:buNone');
    removeBulletChildren(batch, part, pPrId);
    const noIndent = (attr(pPr, 'indent') ?? '0') === '0';
    const zeroMargins = attr(pPr, 'marL') === '0' && noIndent;
    const ownHanging = attr(pPr, 'indent') === String(-HANG);
    // A bullet that Backspace removed left the text at the bullet's position (`marL`, no
    // indent). Bullets back on drop that margin too, so the level's hanging indent returns.
    // Phase 2 puts back a margin that is not the inherited bullet's position: the author's.
    const removedInPlace = suppressed && noIndent && !zeroMargins && kind !== 'none';
    dropped.push(removedInPlace ? attr(pPr, 'marL') : undefined);
    if ((suppressed && zeroMargins) || removedInPlace || (kind === 'none' && ownHanging)) {
      setPPrAttr(batch, part, pPrId, 'marL', null);
      setPPrAttr(batch, part, pPrId, 'indent', null);
    }
  }
  // Phase 2: add what inheritance does not give.
  const inherited = resolve ? resolvedParagraphs(batch, slide, target) : [];
  target.paragraphs.forEach((index, k) => {
    const pPrId = pPrs[k];
    const bullet = inherited[index]?.bullet ?? null;
    if (kind === 'none') {
      if (bullet || !resolve) {
        // CT_TextParagraphProperties order: `a:buNone` after the spacing, color, and size.
        const index = bulletIndex(batch, part, pPrId);
        batch.insert({ op: 'insertNode', part, parent: pPrId, index, spec: aSpec('buNone') });
        setPPrAttr(batch, part, pPrId, 'marL', '0');
        setPPrAttr(batch, part, pPrId, 'indent', '0');
      }
      return;
    }
    const own = dropped[k];
    const at = inherited[index] && Math.round(inherited[index].marL + inherited[index].indent);
    if (bullet && own !== undefined && String(at) !== own) {
      setPPrAttr(batch, part, pPrId, 'marL', own);
      setPPrAttr(batch, part, pPrId, 'indent', '0');
    }
    const wanted = kind === 'bullet' ? 'char' : 'autonum';
    if (bullet?.type === wanted) return;
    insertBullet(batch, part, pPrId, kind);
    if (!bullet) {
      // No inherited list: an explicit hanging indent at this paragraph's level.
      const level = levelOf(batch.element(part, pPrId));
      setPPrAttr(batch, part, pPrId, 'marL', String(HANG + level * LEVEL_STEP));
      setPPrAttr(batch, part, pPrId, 'indent', String(-HANG));
    }
  });
}

/**
 * Remove the bullet or number of each target paragraph that has one, and keep its text where
 * the bullet was: `a:buNone`, `marL` at the bullet's position (`marL + indent`), and no
 * hanging indent. The level stays, so the paragraph keeps its level's spacing and font. This
 * is Backspace at the start of a list item. A probe resolves too, so `can` and `exec` agree
 * on a paragraph with no bullet. Answers whether any paragraph had a bullet to remove.
 * `resolved` gives the paragraphs as the cascade resolves them (`paragraphLookup`).
 */
export function removeBulletInPlace(
  batch: OpBatch,
  part: string,
  slide: number,
  target: ListTarget,
  resolved = resolvedParagraphs(batch, slide, target)
): boolean {
  const all = paragraphsOf(batch.element(part, target.txBodyId));
  let removed = false;
  for (const index of target.paragraphs) {
    const r = resolved[index];
    if (!r?.bullet) continue;
    removed = true;
    const pPrId = ensureParagraphProps(batch, part, all[index].id);
    removeBulletChildren(batch, part, pPrId);
    const at = bulletIndex(batch, part, pPrId);
    batch.insert({ op: 'insertNode', part, parent: pPrId, index: at, spec: aSpec('buNone') });
    const marL = Math.max(0, Math.round(r.marL + r.indent));
    setPPrAttr(batch, part, pPrId, 'marL', String(marL));
    setPPrAttr(batch, part, pPrId, 'indent', '0');
  }
  return removed;
}

/**
 * `removeBulletInPlace` for each of `targets`, with one resolve of the slide: a removal changes
 * only its own paragraphs, never another target's. Answers whether any bullet went.
 */
export function removeBullets(
  batch: OpBatch,
  part: string,
  slide: number,
  targets: readonly ListTarget[]
): boolean {
  const lookup = paragraphLookup(batch, slide);
  let removed = false;
  for (const t of targets)
    removed = removeBulletInPlace(batch, part, slide, t, lookup(t)) || removed;
  return removed;
}

/** Move the list level of one paragraph by `direction` (±1). Returns false at a limit. */
export function shiftListLevel(
  batch: OpBatch,
  part: string,
  pPrId: string,
  direction: number
): boolean {
  const pPr = batch.element(part, pPrId);
  const level = levelOf(pPr);
  const next = Math.max(0, Math.min(8, level + direction));
  if (next === level) return false;
  setPPrAttr(batch, part, pPrId, 'lvl', next ? String(next) : null);
  const marL = attr(pPr, 'marL');
  if (marL !== undefined && /^-?\d+$/.test(marL)) {
    const moved = Math.max(0, Number(marL) + (next - level) * LEVEL_STEP);
    setPPrAttr(batch, part, pPrId, 'marL', String(moved));
  }
  return true;
}
