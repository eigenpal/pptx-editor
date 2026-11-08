// Speaker notes: read the notes text of a slide, write a new text as the smallest edits to the
// notes body, and format a range of it. Unchanged text keeps its runs and their formatting; new
// text takes the formatting at the place it goes, as typing in a notes pane does.

import {
  attr,
  child,
  CONTENT_TYPE,
  elements,
  NS,
  path,
  presentationInfo,
  REL,
  shapeTreeOf,
  textOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import type { NotesFormat, NotesParagraph, NotesRun, TextPoint } from '../../contracts/neutral.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { addOverride, addRelationship, freePartName, XML_DECL } from './package-edits.ts';
import { NS_DECL, SP_TREE_HEAD, slideRef } from './slide-commands.ts';
import {
  deleteRange,
  insertText,
  paragraphLength,
  paragraphsOf,
  runsForFormatting,
  setRunAttr,
  slotsOf,
  splitParagraph,
} from './text-model.ts';
import { cleanText } from './text-units.ts';
import { diffHunks, type TextHunk } from '../../store/package/text-diff.ts';

/** The body placeholder of a notes slide. */
function notesBody(pkg: PptxPackage, notesPart: string): Element | undefined {
  for (const el of elements(shapeTreeOf(pkg, notesPart)))
    if (attr(path(el, 'p:nvSpPr', 'p:nvPr', 'p:ph'), 'type') === 'body')
      return child(el, 'p:txBody');
  return undefined;
}

/** Speaker notes of slide `index` as plain text (paragraphs and line breaks as newlines). */
export function notesText(pkg: PptxPackage, index: number): string {
  const ref = presentationInfo(pkg).slides[index];
  const body = ref?.notesPart ? notesBody(pkg, ref.notesPart) : undefined;
  return elements(body)
    .filter((p) => p.kind === 'paragraph')
    .map((p) =>
      elements(p)
        .map((r) => (r.kind === 'break' ? '\n' : textOf(child(r, 'a:t'))))
        .join('')
    )
    .join('\n');
}

/** The point of a plain-text offset in a notes body (paragraphs joined by one newline). */
function pointAt(paragraphs: readonly Element[], offset: number): TextPoint {
  let start = 0;
  for (let i = 0; i < paragraphs.length; i += 1) {
    const len = paragraphLength(paragraphs[i]);
    if (offset <= start + len) return { paragraph: i, offset: offset - start };
    start += len + 1;
  }
  const last = Math.max(0, paragraphs.length - 1);
  return { paragraph: last, offset: paragraphs.length ? paragraphLength(paragraphs[last]) : 0 };
}

const isLow = (c: number): boolean => c >= 0xdc00 && c <= 0xdfff;

/**
 * Widen each changed region so it never splits a surrogate pair. Null when widened regions
 * would overlap: the caller then replaces the whole changed middle.
 */
function widen(hunks: readonly TextHunk[], a: string, b: string): TextHunk[] | null {
  const out: TextHunk[] = [];
  for (const h of hunks) {
    let { aStart, aEnd, bStart, bEnd } = h;
    // The text around a region is the same in both, so both sides move together.
    if (aStart > 0 && (isLow(a.charCodeAt(aStart)) || isLow(b.charCodeAt(bStart)))) {
      aStart -= 1;
      bStart -= 1;
    }
    if (aEnd < a.length && (isLow(a.charCodeAt(aEnd)) || isLow(b.charCodeAt(bEnd)))) {
      aEnd += 1;
      bEnd += 1;
    }
    const last = out[out.length - 1];
    if (bStart < 0 || bEnd > b.length || (last && aStart < last.aEnd)) return null;
    out.push({ aStart, aEnd, bStart, bEnd });
  }
  return out;
}

/**
 * Turn a field into a run with the same text and the field's own `a:rPr`: an edited field is
 * static text, and it keeps its formatting.
 */
function fieldToRun(batch: OpBatch, part: string, paragraphId: string, field: Element): void {
  const at = batch.element(part, paragraphId).children.findIndex((c) => c.id === field.id);
  const run = batch.insert({
    op: 'insertNode',
    part,
    parent: paragraphId,
    index: at,
    spec: {
      ns: NS.a,
      local: 'r',
      prefix: 'a',
      children: [
        { ns: NS.a, local: 't', prefix: 'a', children: [{ text: textOf(child(field, 'a:t')) }] },
      ],
    },
  });
  const rPr = child(field, 'a:rPr');
  if (rPr) batch.push({ op: 'moveNode', part, node: rPr.id, parent: run, index: 0 });
  batch.push({ op: 'removeNode', part, node: field.id });
}

/**
 * Write `text` into a notes body as the smallest replacements of the parts that differ
 * (`diffHunks`), so every unchanged run keeps its formatting, also between two edits. A field
 * that an edit touches becomes a run first, so its text keeps the field's formatting.
 */
function editBody(batch: OpBatch, part: string, bodyId: string, before: string, text: string) {
  if (before === text) return;
  const hunks = widen(diffHunks(before, text), before, text) ??
    widen(diffHunks(before, text, 0), before, text) ?? [
      { aStart: 0, aEnd: before.length, bStart: 0, bEnd: text.length },
    ];
  let start = 0;
  for (const p of paragraphsOf(batch.element(part, bodyId))) {
    for (const s of slotsOf(p)) {
      const [fs, fe] = [start + s.start, start + s.end];
      const touched = hunks.some((h) =>
        h.aStart === h.aEnd ? h.aStart > fs && h.aStart < fe : h.aStart < fe && h.aEnd > fs
      );
      if (s.kind === 'field' && touched) fieldToRun(batch, part, p.id, s.el);
    }
    start += paragraphLength(p) + 1;
  }
  // From the last region back, so the offsets of the earlier ones stay valid.
  for (const h of [...hunks].reverse()) {
    const paragraphs = paragraphsOf(batch.element(part, bodyId));
    let at = deleteRange(
      batch,
      part,
      bodyId,
      pointAt(paragraphs, h.aStart),
      pointAt(paragraphs, h.aEnd)
    );
    text
      .slice(h.bStart, h.bEnd)
      .split('\n')
      .forEach((line, i) => {
        if (i > 0) at = splitParagraph(batch, part, bodyId, at);
        if (line) at = insertText(batch, part, bodyId, at, line);
      });
  }
}

/**
 * Set the speaker notes of slide `index`, creating the notes slide when it is missing. Only
 * the changed text is rewritten, so formatted runs keep their formatting. Line breaks in the
 * text become paragraphs; characters XML cannot hold are removed.
 */
export function setNotes(
  batch: OpBatch,
  index: number,
  raw: string,
  shown?: readonly NotesParagraph[]
): void {
  const text = cleanText(raw).replace(/\u000B/g, '\n');
  const ref = slideRef(batch.pkg, index);
  if (ref.notesPart && batch.pkg.parts.has(ref.notesPart)) {
    const body = notesBody(batch.pkg, ref.notesPart);
    if (!body) throw new CommandRefusal('errors.unsupported', 'notes without a body placeholder');
    editBody(batch, ref.notesPart, body.id, notesText(batch.pkg, index), text);
    if (shown) matchShownFormats(batch, index, text, shown);
    return;
  }
  if (!text) return;
  const info = presentationInfo(batch.pkg);
  if (!info.notesMasterPart) throw new CommandRefusal('errors.unsupported', 'no notes master');
  const notesName = freePartName(batch, '/ppt/notesSlides/notesSlide', '.xml');
  batch.push({
    op: 'addPart',
    part: notesName,
    xml:
      `${XML_DECL}<p:notes ${NS_DECL}><p:cSld><p:spTree>${SP_TREE_HEAD}` +
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>' +
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>' +
      '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>',
  });
  // The text goes in through the text model as nodes, never through the XML template.
  editBody(batch, notesName, notesBody(batch.pkg, notesName)!.id, '', text);
  addRelationship(batch, notesName, REL.notesMaster, info.notesMasterPart);
  addRelationship(batch, notesName, REL.slide, ref.partName);
  addOverride(batch, notesName, CONTENT_TYPE.notesSlide);
  addRelationship(batch, ref.partName, REL.notesSlide, notesName);
  if (shown) matchShownFormats(batch, index, text, shown);
}

/** The format of each UTF-16 unit of `paragraphs` (joined by `\n`), as `NotesRun` entries. */
function formatsPerUnit(paragraphs: readonly NotesParagraph[]): NotesRun[] {
  const out: NotesRun[] = [];
  paragraphs.forEach((p, i) => {
    if (i > 0) out.push(LINE_BREAK);
    for (const r of p.runs) for (let k = 0; k < r.text.length; k += 1) out.push(r);
  });
  return out;
}

/**
 * The `paragraphs` of a `setNotes` command when they are well formed (paragraphs of runs with
 * string text and boolean formats), else undefined: the text is written without them.
 */
export function notesShown(paragraphs: unknown): readonly NotesParagraph[] | undefined {
  if (!Array.isArray(paragraphs)) return undefined;
  const run = (r: unknown): boolean =>
    !!r &&
    typeof (r as NotesRun).text === 'string' &&
    (['bold', 'italic', 'underline', 'strike'] as const).every(
      (f) => typeof (r as NotesRun)[f] === 'boolean'
    );
  const ok = paragraphs.every(
    (p: unknown) =>
      !!p && Array.isArray((p as NotesParagraph).runs) && (p as NotesParagraph).runs.every(run)
  );
  return ok ? (paragraphs as readonly NotesParagraph[]) : undefined;
}

const LINE_BREAK: NotesRun = {
  text: '\n',
  bold: false,
  italic: false,
  underline: false,
  strike: false,
};

/**
 * Give each character of slide `index`'s notes the format it has in `shown`, the notes as a
 * field showed them while the user typed: text typed after a bold word, or after a format
 * armed at the caret, is stored as it was shown. Skipped when `shown` holds another text (a
 * character that `setNotes` cleaned). Line breaks carry no format.
 */
function matchShownFormats(
  batch: OpBatch,
  index: number,
  text: string,
  shown: readonly NotesParagraph[]
): void {
  if (shown.map((p) => p.runs.map((r) => r.text).join('')).join('\n') !== text) return;
  const want = formatsPerUnit(shown);
  const have = formatsPerUnit(notesParagraphs(batch.pkg, index));
  if (have.length !== text.length) return;
  for (const format of Object.keys(FORMATS) as NotesFormat[]) {
    let from = -1;
    for (let i = 0; i <= text.length; i += 1) {
      const differs = i < text.length && text[i] !== '\n' && want[i]![format] !== have[i]![format];
      // A range ends at a character that matches, at a break, or where the wanted value flips.
      if (from >= 0 && (!differs || want[i]![format] !== want[from]![format])) {
        formatNotes(batch, index, from, i, format, want[from]![format]);
        from = -1;
      }
      if (differs && from < 0) from = i;
    }
  }
}

/** The `a:rPr` attribute of each notes format, and its on and off values (§21.1.2.3.9). */
const FORMATS: Record<NotesFormat, readonly [name: string, on: string, off: string]> = {
  bold: ['b', '1', '0'],
  italic: ['i', '1', '0'],
  underline: ['u', 'sng', 'none'],
  strike: ['strike', 'sngStrike', 'noStrike'],
};

/** Whether an `a:rPr` or `a:defRPr` value turns a format on, or undefined when absent. */
function formatOf(props: Element | undefined, format: NotesFormat): boolean | undefined {
  const value = attr(props, FORMATS[format][0]);
  if (value === undefined) return undefined;
  if (format === 'bold' || format === 'italic') return value === '1' || value === 'true';
  return value !== FORMATS[format][2];
}

/**
 * The notes text style a paragraph starts from, in cascade order: its level of the notes
 * master's `p:notesStyle` (§19.3.1.28), the notes master's body placeholder list style, then
 * the body's own list style.
 */
function levelDefaults(pkg: PptxPackage, body: Element, p: Element): Element[] {
  const level = Math.max(0, Math.min(8, Number(attr(child(p, 'a:pPr'), 'lvl')) || 0));
  const name = `a:lvl${level + 1}pPr`;
  const master = presentationInfo(pkg).notesMasterPart;
  const root = master ? xmlPartOf(pkg, master)?.root : undefined;
  const masterBody = master ? notesBody(pkg, master) : undefined;
  return [
    path(child(root, 'p:notesStyle'), name, 'a:defRPr'),
    path(child(masterBody, 'a:lstStyle'), name, 'a:defRPr'),
    path(child(body, 'a:lstStyle'), name, 'a:defRPr'),
  ].filter((e): e is Element => !!e);
}

/**
 * The speaker notes of slide `index` as paragraphs of runs with their character formatting:
 * what a formatted notes view shows. The text joins (paragraphs by newlines) to `notesText`.
 */
export function notesParagraphs(pkg: PptxPackage, index: number): NotesParagraph[] {
  const ref = presentationInfo(pkg).slides[index];
  const body = ref?.notesPart ? notesBody(pkg, ref.notesPart) : undefined;
  if (!body) return [];
  return paragraphsOf(body).map((p) => {
    const defaults = levelDefaults(pkg, body, p);
    const runs: NotesRun[] = [];
    for (const s of slotsOf(p)) {
      const props = child(s.el, 'a:rPr');
      const read = (f: NotesFormat): boolean =>
        formatOf(props, f) ?? defaults.reduce<boolean>((on, d) => formatOf(d, f) ?? on, false);
      const run = {
        text: s.text,
        bold: read('bold'),
        italic: read('italic'),
        underline: read('underline'),
        strike: read('strike'),
      };
      const last = runs[runs.length - 1];
      // Neighbors with the same formatting join, so a view draws one span for them.
      if (
        last &&
        last.bold === run.bold &&
        last.italic === run.italic &&
        last.underline === run.underline &&
        last.strike === run.strike
      )
        runs[runs.length - 1] = { ...last, text: last.text + run.text };
      else runs.push(run);
    }
    return { runs };
  });
}

/**
 * Bold, italic, underline, or strike through the plain-text range `[start, end)` of slide
 * `index`'s notes (offsets as in `notesText`). `value` absent toggles: on, unless every
 * character of the range already has the format. Runs split at the range edges, so the text
 * around keeps its formatting.
 */
export function formatNotes(
  batch: OpBatch,
  index: number,
  start: number,
  end: number,
  format: NotesFormat,
  value?: boolean
): void {
  const ref = slideRef(batch.pkg, index);
  const part = ref.notesPart;
  const body = part && batch.pkg.parts.has(part) ? notesBody(batch.pkg, part) : undefined;
  if (!part || !body) throw new CommandRefusal('disabled.noNotes');
  if (!Object.hasOwn(FORMATS, format)) throw new CommandRefusal('errors.unsupported', 'format');
  const length = notesText(batch.pkg, index).length;
  const [from, to] = [Math.max(0, Math.min(start, end)), Math.min(length, Math.max(start, end))];
  if (!Number.isInteger(start) || !Number.isInteger(end) || from >= to)
    throw new CommandRefusal('disabled.selectNotesText');
  const paragraphs = paragraphsOf(body);
  const range = { from: pointAt(paragraphs, from), to: pointAt(paragraphs, to) };
  // What the range shows now, read before the runs split.
  const shown = notesParagraphs(batch.pkg, index);
  let at = 0;
  let all = true;
  shown.forEach((p) => {
    for (const r of p.runs) {
      // A line break carries no format of its own, as the notes field shows it.
      const text = r.text.replace(/\n/g, '');
      if (at < to && at + r.text.length > from && text && !r[format]) all = false;
      at += r.text.length;
    }
    at += 1;
  });
  const on = value ?? !all;
  const [name, yes, no] = FORMATS[format];
  // Fields are never split: a field the range touches takes the format whole.
  const fields: Element[] = [];
  let base = 0;
  for (const p of paragraphs) {
    for (const s of slotsOf(p))
      if (s.kind === 'field' && base + s.start < to && base + s.end > from) fields.push(s.el);
    base += paragraphLength(p) + 1;
  }
  for (const run of [...runsForFormatting(batch, part, body.id, range), ...fields])
    setRunAttr(batch, part, run, name, on ? yes : no);
}
