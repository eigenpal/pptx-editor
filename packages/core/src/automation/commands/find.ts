// Find and replace over the store's text model. The search reads the canonical tree: the text
// of every shape (also inside groups, and in the `mc:AlternateContent` branch the slide draws),
// every table cell, and the speaker notes, in deck order.
// Replacements compile to store ops through an OpBatch, so they undo like any other edit.
//
// Offsets are the text model's: a `slidenum` field is as long as the number it shows. Fields
// are atomic, so a match never covers one; a line break counts as one `\n` character.

import {
  ancestorsOf,
  attr,
  child,
  children,
  elements,
  path,
  presentationInfo,
  tableAsDrawn,
  shapeTreeOf,
  TABLE_MAX_COLUMNS,
  type Element,
  joinedTextBodies,
  type PptxPackage,
  xmlPartOf,
} from '../../store/index.ts';
import { contentElements } from '../../cascade/index.ts';
import type { FindMatch, FindOptions } from '../../contracts/neutral.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import {
  deleteRange,
  insertText,
  paragraphsOf,
  slideNumberOf,
  slotsOf,
  splitsPair,
} from './text-model.ts';
import { cleanText } from './text-units.ts';
import { joinTextBodies } from './table-model.ts';

/** One text body find reads, in deck order. */
export interface TextBodyRef {
  readonly slide: number;
  readonly where: 'slide' | 'notes';
  readonly part: string;
  readonly nodeId: string;
  readonly txBodyId: string;
}

/** The longest query and replacement find accepts. */
export const MAX_FIND_TEXT = 1000;
/** Group nesting find follows, as deep as the cascade does. */
const MAX_DEPTH = 32;
/** Stands in for a field's text, so no query matches inside a field. */
const FIELD_CHAR = '￼';

/**
 * Text bodies of a table: one per drawn cell, row by row, by the cascade's merge grid. A cell
 * that a merged cell covers is not drawn, and its text (such as the copy a merge in a
 * collaboration room keeps) is not found. A cell with a merge marker that no span reaches is
 * drawn, and found.
 */
function tableBodies(frame: Element): Element[] {
  // The cells the cascade draws (`tableAsDrawn`), so find reads what the slide shows. A blank
  // cell that the save adds to a short row has no text, and no node in the store.
  const saved = tableAsDrawn(frame, TABLE_MAX_COLUMNS);
  const out: Element[] = [];
  for (const b of saved?.grid.blocks ?? []) {
    const body = child(saved!.cells[b.row][b.col], 'a:txBody');
    if (body?.id) out.push(body);
  }
  return out;
}

/**
 * Shape text bodies under a shape tree or group, in tree order. `mc:AlternateContent` gives the
 * branch that the cascade draws (`alternateBranch`), so find reads what the slide shows.
 */
function shapeBodies(
  parent: Element | undefined,
  root: Element | undefined,
  depth: number,
  add: (nodeId: string, body: Element) => void
): void {
  for (const el of contentElements(parent, root)) {
    if (el.kind === 'sp') {
      const body = child(el, 'p:txBody');
      if (body) add(el.id, body);
    } else if (el.kind === 'grpSp' && depth < MAX_DEPTH) shapeBodies(el, root, depth + 1, add);
    else if (el.kind === 'graphicFrame') for (const body of tableBodies(el)) add(el.id, body);
  }
}

/** Every text body find reads with its element: each slide's shapes and tables, then notes. */
function bodies(pkg: PptxPackage): { ref: TextBodyRef; el: Element }[] {
  const out: { ref: TextBodyRef; el: Element }[] = [];
  presentationInfo(pkg).slides.forEach((s, slide) => {
    shapeBodies(shapeTreeOf(pkg, s.partName), xmlPartOf(pkg, s.partName)?.root, 0, (nodeId, el) =>
      out.push({ ref: { slide, where: 'slide', part: s.partName, nodeId, txBodyId: el.id }, el })
    );
    if (!s.notesPart) return;
    for (const sp of elements(shapeTreeOf(pkg, s.notesPart))) {
      const el = child(sp, 'p:txBody');
      if (el && attr(path(sp, 'p:nvSpPr', 'p:nvPr', 'p:ph'), 'type') === 'body')
        out.push({
          ref: { slide, where: 'notes', part: s.notesPart, nodeId: sp.id, txBodyId: el.id },
          el,
        });
    }
  });
  return out;
}

/** Every text body find reads, in deck order. */
export function textBodies(pkg: PptxPackage): TextBodyRef[] {
  return bodies(pkg).map((b) => b.ref);
}

/** The text of a paragraph as find reads it: fields are object characters. */
function searchText(p: Element, slideNumber: number | undefined): string {
  return slotsOf(p, slideNumber)
    .map((s) => (s.kind === 'field' ? FIELD_CHAR.repeat(s.text.length) : s.text))
    .join('');
}

/** Lower case one code point at a time, only where it keeps the length (offsets stay true). */
function fold(s: string): string {
  let out = '';
  for (const ch of s) {
    const lower = ch.toLowerCase();
    out += lower.length === ch.length ? lower : ch;
  }
  return out;
}

const WORD = /^[\p{L}\p{N}\p{M}_]$/u;

/** Whether the code point that ends at `end` (before) or starts at `start` (after) is a word character. */
function wordBefore(text: string, end: number): boolean {
  if (end <= 0) return false;
  const from = splitsPair(text, end - 1) ? end - 2 : end - 1;
  return WORD.test(text.slice(from, end));
}
function wordAfter(text: string, start: number): boolean {
  if (start >= text.length) return false;
  const to = splitsPair(text, start + 1) ? start + 2 : start + 1;
  return WORD.test(text.slice(start, to));
}

/** Non-overlapping matches of `query` in `text`, as `[start, end)` pairs. */
export function matchesIn(
  text: string,
  query: string,
  options: FindOptions = {}
): [number, number][] {
  if (!query || query.length > text.length) return [];
  const hay = options.matchCase ? text : fold(text);
  const needle = options.matchCase ? query : fold(query);
  const out: [number, number][] = [];
  let from = 0;
  for (;;) {
    const at = hay.indexOf(needle, from);
    if (at < 0) break;
    const end = at + needle.length;
    const ok =
      !splitsPair(text, at) &&
      !splitsPair(text, end) &&
      (!options.wholeWord || (!wordBefore(text, at) && !wordAfter(text, end)));
    if (ok) out.push([at, end]);
    from = ok ? end : at + 1;
  }
  return out;
}

/**
 * Every match of `query` in the deck, ordered by slide, then by position: the slide's shapes
 * in tree order (group children and table cells in place), then its speaker notes.
 */
export function findMatches(
  pkg: PptxPackage,
  query: string,
  options: FindOptions = {}
): FindMatch[] {
  if (!query || query.length > MAX_FIND_TEXT) return [];
  const out: FindMatch[] = [];
  for (const { ref, el } of bodies(pkg)) {
    // A `slidenum` field on a slide shows the slide number; in notes, its stored text.
    const n = ref.where === 'slide' ? ref.slide + 1 : undefined;
    paragraphsOf(el).forEach((p, paragraph) => {
      const text = searchText(p, n);
      for (const [start, end] of matchesIn(text, query, options))
        out.push({ ...ref, paragraph, start, end, text: text.slice(start, end) });
    });
  }
  return out;
}

/** A replacement as one line of text a slide can show. */
export function cleanReplacement(text: string): string {
  if (text.length > MAX_FIND_TEXT) throw new CommandRefusal('errors.findTooLong');
  return cleanText(text).replace(/[\n\t\u000B]/g, ' ');
}

/**
 * Replace one match. The new text goes into the run of the first replaced character, so it
 * keeps that character's formatting; then the matched text is deleted. Refuses a match whose
 * text is no longer there.
 */
function replaceOne(batch: OpBatch, m: FindMatch, replacement: string): void {
  const p = paragraphsOf(batch.element(m.part, m.txBodyId))[m.paragraph];
  if (!p) throw new CommandRefusal('errors.findStale');
  const n = slideNumberOf(batch, m.part);
  if (searchText(p, n).slice(m.start, m.end) !== m.text || m.end <= m.start)
    throw new CommandRefusal('errors.findStale');
  let shift = 0;
  if (replacement) {
    const host = slotsOf(p, n).find(
      (s) => s.kind === 'run' && s.start <= m.start && m.start < s.end
    );
    const t = host ? child(host.el, 'a:t') : undefined;
    if (host && t) {
      const local = m.start - host.start;
      batch.push({
        op: 'setText',
        part: m.part,
        node: t.id,
        text: host.text.slice(0, local) + replacement + host.text.slice(local),
      });
    } else
      insertText(
        batch,
        m.part,
        m.txBodyId,
        { paragraph: m.paragraph, offset: m.start },
        replacement
      );
    shift = replacement.length;
  }
  deleteRange(
    batch,
    m.part,
    m.txBodyId,
    { paragraph: m.paragraph, offset: m.start + shift },
    { paragraph: m.paragraph, offset: m.end + shift }
  );
}

/**
 * Replace matches in one batch. Later matches go first, so the offsets of earlier matches in
 * the same paragraph stay true.
 */
export function replaceMatches(
  batch: OpBatch,
  matches: readonly FindMatch[],
  replacement: string
): void {
  const text = cleanReplacement(replacement);
  const order = [...matches].sort(
    (a, b) =>
      a.part.localeCompare(b.part) ||
      a.txBodyId.localeCompare(b.txBodyId) ||
      b.paragraph - a.paragraph ||
      b.start - a.start
  );
  // Only a cell with more than one text body can need a join. Check each text body once, not
  // each match: the check reads the whole part, and a long run can hold thousands of matches.
  const inSplitCell = new Map<string, boolean>();
  const runs = createRunEdits(batch);
  for (const m of order) {
    const key = `${m.part}\u0000${m.txBodyId}`;
    let split = inSplitCell.get(key);
    if (split === undefined) inSplitCell.set(key, (split = isSplitCellBody(batch, m)));
    if (split) {
      runs.flush();
      joinCellBodies(batch, m);
    }
    if (runs.replace(m, text)) continue;
    runs.flush();
    replaceOne(batch, m, text);
  }
  runs.flush();
}

/** A run slot of a paragraph: its text node, and where its text starts and ends. */
interface RunSlot {
  readonly textId: string;
  readonly start: number;
  readonly end: number;
  readonly text: string;
}

/**
 * Replacements inside one run, written once for the run: a run with thousands of matches is
 * written one time, not one time for each match, which made a long run quadratic. Matches
 * come latest first, so the slots before the run being written keep their offsets. A match
 * that crosses runs, or that would leave its run empty, is not taken (`replace` answers false),
 * and `replaceOne` handles it after a `flush`.
 */
function createRunEdits(batch: OpBatch) {
  let paragraph: { key: string; slots: readonly RunSlot[] } | null = null;
  let open: { part: string; slot: RunSlot; text: string } | null = null;
  const slotsAt = (m: FindMatch): readonly RunSlot[] => {
    const key = `${m.part}\u0000${m.txBodyId}\u0000${m.paragraph}`;
    if (paragraph?.key === key) return paragraph.slots;
    const p = paragraphsOf(batch.element(m.part, m.txBodyId))[m.paragraph];
    const slots: RunSlot[] = [];
    for (const s of p ? slotsOf(p, slideNumberOf(batch, m.part)) : []) {
      const t = s.kind === 'run' ? child(s.el, 'a:t') : undefined;
      if (t) slots.push({ textId: t.id, start: s.start, end: s.end, text: s.text });
    }
    paragraph = { key, slots };
    return slots;
  };
  /** The run slot that holds all of `[start, end)`, by binary search over the sorted slots. */
  const slotOf = (slots: readonly RunSlot[], start: number, end: number): RunSlot | undefined => {
    let lo = 0;
    let hi = slots.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const s = slots[mid]!;
      if (s.end <= start) lo = mid + 1;
      else if (s.start > start) hi = mid - 1;
      else return end <= s.end ? s : undefined;
    }
    return undefined;
  };
  const flush = (): void => {
    if (open && open.text !== open.slot.text)
      batch.push({ op: 'setText', part: open.part, node: open.slot.textId, text: open.text });
    open = null;
    // A write outside this helper can change the paragraph: read it again next time.
    paragraph = null;
  };
  return {
    flush,
    replace(m: FindMatch, replacement: string): boolean {
      if (m.end <= m.start) return false;
      const slot = slotOf(slotsAt(m), m.start, m.end);
      if (!slot) return false;
      if (open && open.slot !== slot) {
        const keep = paragraph;
        flush();
        paragraph = keep;
      }
      const current = open?.text ?? slot.text;
      const from = m.start - slot.start;
      const to = m.end - slot.start;
      if (current.slice(from, to) !== m.text) throw new CommandRefusal('errors.findStale');
      const next = current.slice(0, from) + replacement + current.slice(to);
      if (next === '') return false;
      open = { part: m.part, slot, text: next };
      return true;
    },
  };
}

/** Whether the match's text body is the first of a cell with more than one text body. */
function isSplitCellBody(batch: OpBatch, m: FindMatch): boolean {
  const tc = ancestorsOf(xmlPartOf(batch.pkg, m.part)!, m.txBodyId).at(-1);
  const bodies = children(tc, 'a:txBody');
  return tc?.local === 'tc' && bodies[0]?.id === m.txBodyId && bodies.length > 1;
}

/**
 * Join the text bodies of the cell that holds a match past its first body's paragraphs: find
 * reads such a cell's bodies as one (`cellAsSaved`), and so does an edit of the cell.
 */
function joinCellBodies(batch: OpBatch, m: FindMatch): void {
  const tc = ancestorsOf(xmlPartOf(batch.pkg, m.part)!, m.txBodyId).at(-1);
  const [first, ...rest] = children(tc, 'a:txBody');
  if (tc?.local !== 'tc' || first?.id !== m.txBodyId || rest.length === 0) return;
  // A match in the first body's own paragraphs keeps its place, unless that body is blank and
  // gives way to the others (`joinedTextBodies`): then the join renumbers the paragraphs.
  const own = joinedTextBodies(children(tc, 'a:txBody')).includes(first);
  if (own && m.paragraph < paragraphsOf(first).length) return;
  joinTextBodies(batch, m.part, first, rest);
}
