// Insert > Slide numbers, Insert > Footer, and Insert > Date and time.
//
// A slide shows its footer and its date through the footer (`p:ph type="ftr"`) and date
// (`p:ph type="dt"`) placeholders that its layout defines, or its master when the layout has
// none (ECMA-376 Part 1, 19.3.1.36 ph). Turning one on gives each target slide that placeholder,
// so position and style come from the layout. A slide that has one already keeps its own box and
// format; it keeps its text too, unless the command gives new text. Turning one off removes the
// slide's placeholders of that type. A slide whose layout and master define no such placeholder
// is skipped.

import {
  attr,
  elements,
  find,
  path,
  presentationInfo,
  specOf,
  textOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { fieldId, nextShapeId, spTreeOf, xmlAttr, xmlText } from './package-edits.ts';
import { specFromXml } from './shape-commands.ts';
import { removeShapes } from './shape-removal.ts';
import { setSlideNumbers } from './slide-numbers.ts';
import type {
  DateTimeState as DateShown,
  EditorCommand,
  HeaderFooterTargets,
} from '../../contracts/neutral.ts';

/** The header and footer placeholders that these commands write. */
export type HeaderFooterKind = 'ftr' | 'dt';

const NAMES: Record<HeaderFooterKind, string> = { ftr: 'Footer', dt: 'Date' };

const REFUSALS: Record<HeaderFooterKind, string> = {
  ftr: 'disabled.noFooterPlaceholder',
  dt: 'disabled.noDatePlaceholder',
};

/** The longest footer or date text, in characters. */
export const MAX_HEADER_FOOTER_TEXT = 1024;

const isKind =
  (kind: HeaderFooterKind) =>
  (el: Element): boolean =>
    attr(path(el, 'p:nvSpPr', 'p:nvPr', 'p:ph'), 'type') === kind;

function shapesOfKind(pkg: PptxPackage, part: string | undefined, kind: HeaderFooterKind) {
  const tree = part ? find(xmlPartOf(pkg, part)?.root, 'p:spTree') : undefined;
  return tree ? elements(tree).filter(isKind(kind)) : [];
}

/** What one placeholder shows: its text, and the field type of its first field, if any. */
interface Shown {
  readonly text: string;
  readonly fieldType: string | undefined;
}

/** A placeholder without text shows nothing, so it does not count as a footer or a date. */
function shownBy(shape: Element | undefined): Shown | null {
  if (!shape) return null;
  const body = find(shape, 'p:txBody');
  const text = textOf(body);
  if (text.trim() === '') return null;
  const field = body ? find(body, 'a:fld') : undefined;
  return { text, fieldType: field ? (attr(field, 'type') ?? '') : undefined };
}

/** The footer and date of one parsed slide part. A part keeps its object until it changes. */
const cache = new WeakMap<object, { readonly ftr: Shown | null; readonly dt: Shown | null }>();

function shownOn(pkg: PptxPackage, part: string | undefined, kind: HeaderFooterKind) {
  const xml = part ? xmlPartOf(pkg, part) : undefined;
  if (!xml) return null;
  let got = cache.get(xml);
  if (!got) {
    got = {
      ftr: shownBy(shapesOfKind(pkg, part, 'ftr')[0]),
      dt: shownBy(shapesOfKind(pkg, part, 'dt')[0]),
    };
    cache.set(xml, got);
  }
  return got[kind];
}

/**
 * Which slides a deck-level read may parse. A progressive open passes the slides that are open
 * already, so a read does not parse the slides that wait. Default: all.
 */
export type SlideReady = (index: number) => boolean;

/** What the placeholder of `kind` shows: on slide `index`, else on the first slide with one. */
function deckShown(
  pkg: PptxPackage,
  kind: HeaderFooterKind,
  index: number,
  ready: SlideReady
): Shown | null {
  const slides = presentationInfo(pkg).slides;
  const own = shownOn(pkg, slides[index]?.partName, kind);
  if (own) return own;
  for (const [i, slide] of slides.entries()) {
    const got = ready(i) ? shownOn(pkg, slide.partName, kind) : null;
    if (got) return got;
  }
  return null;
}

/**
 * The footer text that the deck shows: slide `index`'s, else that of the first slide with a
 * footer. Null when no slide shows one. Footers apply to all slides, so this is the menu state.
 */
export function footerText(
  pkg: PptxPackage,
  index: number,
  ready: SlideReady = () => true
): string | null {
  return deckShown(pkg, 'ftr', index, ready)?.text ?? null;
}

/** The date that the deck shows, read as `footerText` reads the footer. Null when none. */
export function dateShown(
  pkg: PptxPackage,
  index: number,
  ready: SlideReady = () => true
): DateShown | null {
  const got = deckShown(pkg, 'dt', index, ready);
  return got ? { text: got.text, fixed: got.fieldType === undefined } : null;
}

/** The paragraph content of a footer or date: plain text, or a date field. */
function contentXml(text: string, what: string, field: string | undefined, rPr: boolean): string {
  const t = xmlText(text, what);
  const format = rPr ? '<a:rPr lang="en-US"/>' : '';
  // A field type copied from another slide is file content: escape it.
  const type = field ? ` type="${xmlAttr(field, 'field type')}"` : '';
  return field !== undefined
    ? `<a:fld id="${fieldId()}"${type}>${format}<a:t>${t}</a:t></a:fld>`
    : `<a:r>${format}<a:t>${t}</a:t></a:r>`;
}

function paragraphXml(text: string, what: string, field: string | undefined): string {
  return `<a:p>${contentXml(text, what, field, true)}<a:endParaRPr lang="en-US"/></a:p>`;
}

const TEXT_CONTENT = new Set(['r', 'fld', 'br']);

/**
 * Show `text` in an existing placeholder. The first paragraph and its first run or field stay,
 * so the placeholder keeps its paragraph and character format; other text is removed. A date
 * that changes between a field and fixed text gets a new run in place of the old one, with the
 * old character format.
 */
function replaceText(
  batch: OpBatch,
  part: string,
  shape: Element,
  what: string,
  text: string,
  field: string | undefined
): void {
  const body = find(shape, 'p:txBody');
  if (!body) {
    // `p:txBody` follows `p:spPr` and `p:style` (`CT_Shape`).
    let after = -1;
    shape.children.forEach((c, i) => {
      if (c.type === 'element' && ['nvSpPr', 'spPr', 'style'].includes(c.local)) after = i;
    });
    batch.insert({
      op: 'insertNode',
      part,
      parent: shape.id,
      index: after + 1,
      spec: specFromXml(
        `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphXml(text, what, field)}</p:txBody>`
      ),
    });
    return;
  }
  const [first, ...rest] = elements(body).filter((c) => c.local === 'p');
  for (const p of rest) batch.push({ op: 'removeNode', part, node: p.id });
  if (!first) {
    batch.insert({
      op: 'insertNode',
      part,
      parent: body.id,
      index: body.children.length,
      spec: specFromXml(paragraphXml(text, what, field)),
    });
    return;
  }
  const content = elements(first).filter((c) => TEXT_CONTENT.has(c.local));
  const run = content.find((c) => c.local === 'r' || c.local === 'fld');
  const wanted = field !== undefined ? 'fld' : 'r';
  const t = run?.local === wanted ? find(run, 'a:t') : undefined;
  if (run && t) {
    xmlText(text, what);
    batch.push({ op: 'setText', part, node: t.id, text });
    // The new text has the field's format, so a field of another format takes that format.
    if (field !== undefined && attr(run, 'type') !== field)
      batch.push({
        op: 'setAttribute',
        part,
        node: run.id,
        name: { ns: '', local: 'type' },
        value: field,
      });
  } else {
    const rPr = run ? find(run, 'a:rPr') : undefined;
    const end = first.children.findIndex((c) => c.type === 'element' && c.local === 'endParaRPr');
    const endPr = end < 0 ? undefined : (first.children[end] as Element);
    const created = batch.insert({
      op: 'insertNode',
      part,
      parent: first.id,
      index: end < 0 ? first.children.length : end,
      spec: specFromXml(contentXml(text, what, field, !rPr && !endPr)),
    });
    // The new run or field keeps the character format of the old one; in an empty paragraph, it
    // takes the end-of-paragraph format (`a:endParaRPr` is a `CT_TextCharacterProperties` too).
    if (rPr) batch.push({ op: 'moveNode', part, node: rPr.id, parent: created, index: 0 });
    else if (endPr)
      batch.insert({
        op: 'insertNode',
        part,
        parent: created,
        index: 0,
        spec: { ...specOf(endPr), local: 'rPr' },
      });
  }
  for (const c of content) if (c !== run || !t) batch.push({ op: 'removeNode', part, node: c.id });
}

/** The target slides, split into those that show and those that hide. */
export function headerFooterTargets(
  pkg: PptxPackage,
  show: boolean,
  options: HeaderFooterTargets
): { readonly show: readonly number[]; readonly hide: readonly number[] } {
  const slides = presentationInfo(pkg).slides;
  const targets = (options.indices ?? slides.map((_, i) => i)).filter((i) => slides[i]);
  if (!show) return { show: [], hide: targets };
  if (!options.hideOnTitleSlides) return { show: targets, hide: [] };
  const isTitle = (i: number): boolean => {
    const layout = slides[i]!.layoutPart;
    return !!layout && attr(xmlPartOf(pkg, layout)?.root, 'type') === 'title';
  };
  return { show: targets.filter((i) => !isTitle(i)), hide: targets.filter(isTitle) };
}

/** What to show: new text, or (`text` absent) the slide's own, else the deck's, else `fallback`. */
interface Showing {
  readonly text: string | undefined;
  /** The field type for new text; undefined for fixed text. */
  readonly field: string | undefined;
  readonly fallback: (() => Shown) | undefined;
  /** Without text: turn the date fields that slides keep into fixed text. */
  readonly toFixed?: boolean;
}

/**
 * Show the placeholder of `kind` on the `show` slides and hide it on the `hide` slides. Returns
 * the number of slides changed. Refuses (`errors.unsupported`) blank, multi-line, or too long
 * text and an empty target list, and refuses when none of the targets can show the placeholder.
 */
function applyKind(
  batch: OpBatch,
  kind: HeaderFooterKind,
  targets: { readonly show: readonly number[]; readonly hide: readonly number[] },
  showing: Showing
): number {
  const what = NAMES[kind];
  // One line of text: a line break inside `a:t` is not a line break (`a:br` is).
  if (
    showing.text !== undefined &&
    targets.show.length > 0 &&
    (showing.text.trim() === '' ||
      showing.text.length > MAX_HEADER_FOOTER_TEXT ||
      /[\r\n\u000b\u0085\u2028\u2029]/.test(showing.text))
  )
    throw new CommandRefusal('errors.unsupported', what);
  if (targets.show.length + targets.hide.length === 0)
    throw new CommandRefusal('errors.unsupported', 'slides');
  const pkg = batch.pkg;
  const slides = presentationInfo(pkg).slides;
  let changed = 0;
  let possible = 0;
  // Text absent: a slide without the placeholder shows what the first slide with one shows.
  const deck =
    showing.text === undefined && targets.show.length > 0
      ? (deckShown(pkg, kind, targets.show[0]!, () => true) ?? showing.fallback?.() ?? null)
      : null;
  for (const index of [...targets.show, ...targets.hide]) {
    const slide = slides[index]!;
    const tree = spTreeOf(batch, slide.partName);
    const existing = elements(tree).filter(isKind(kind));
    const template =
      shapesOfKind(pkg, slide.layoutPart, kind)[0] ?? shapesOfKind(pkg, slide.masterPart, kind)[0];
    if (!template && existing.length === 0) continue;
    possible += 1;
    if (!targets.show.includes(index)) {
      removeShapes(
        batch,
        slide.partName,
        existing.map((el) => el.id)
      );
      changed += existing.length > 0 ? 1 : 0;
      continue;
    }
    if (existing.length > 0) {
      if (showing.text === undefined) {
        // Keeps its own text; an empty one takes the deck's, and `fixed` turns a date field
        // into its text.
        let edited = false;
        for (const shape of existing) {
          const own = shownBy(shape);
          if (!own && deck) {
            const field = showing.toFixed ? undefined : deck.fieldType;
            replaceText(batch, slide.partName, shape, what, deck.text, field);
          } else if (own && showing.toFixed && own.fieldType !== undefined) {
            replaceText(batch, slide.partName, shape, what, own.text, undefined);
          } else continue;
          edited = true;
        }
        changed += edited ? 1 : 0;
        continue;
      }
      for (const shape of existing)
        replaceText(batch, slide.partName, shape, what, showing.text, showing.field);
      changed += 1;
      continue;
    }
    const text = showing.text ?? deck?.text;
    if (text === undefined) throw new CommandRefusal('disabled.noFooterText');
    const field = showing.text !== undefined || showing.toFixed ? showing.field : deck?.fieldType;
    const ph = path(template!, 'p:nvSpPr', 'p:nvPr', 'p:ph');
    const phAttrs = ['type', 'orient', 'sz', 'idx']
      .map((a) => {
        const value = attr(ph, a);
        return value !== undefined ? ` ${a}="${xmlAttr(value, 'placeholder')}"` : '';
      })
      .join('');
    const id = nextShapeId(batch, tree);
    const xml =
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${what} Placeholder ${id - 1}"/>` +
      `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${phAttrs}/></p:nvPr></p:nvSpPr>` +
      `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphXml(text, what, field)}` +
      '</p:txBody></p:sp>';
    const ext = tree.children.findIndex((c) => c.type === 'element' && c.local === 'extLst');
    batch.insert({
      op: 'insertNode',
      part: slide.partName,
      parent: tree.id,
      index: ext < 0 ? tree.children.length : ext,
      spec: specFromXml(xml),
    });
    changed += 1;
  }
  if (possible === 0) throw new CommandRefusal(REFUSALS[kind]);
  return changed;
}

/**
 * Show or hide the footer. With `text`, every slide that shows it shows `text`; without it, a
 * slide keeps its footer, and a slide without one takes the deck's footer text. Refuses without
 * text when no slide has a footer (`disabled.noFooterText`), and for blank or multi-line text.
 */
export function setFooter(
  batch: OpBatch,
  show: boolean,
  text?: string,
  options: HeaderFooterTargets = {}
): number {
  const targets = headerFooterTargets(batch.pkg, show, options);
  return applyKind(batch, 'ftr', targets, { text, field: undefined, fallback: undefined });
}

/**
 * Show or hide the date. With `text`, every slide that shows it shows `text`: a `datetime1` field
 * that a presentation application can update when it opens the file, or fixed text with `fixed`.
 * Without `text`, a slide keeps its date, and a slide without one takes the deck's date, else
 * today's date as a field; `fixed` then turns the dates into fixed text.
 */
export function setDateTime(
  batch: OpBatch,
  show: boolean,
  text?: string,
  fixed = false,
  options: HeaderFooterTargets = {}
): number {
  const targets = headerFooterTargets(batch.pkg, show, options);
  return applyKind(batch, 'dt', targets, {
    text,
    field: fixed ? undefined : 'datetime1',
    fallback: () => ({ text: shortDate(), fieldType: 'datetime1' }),
    toFixed: fixed && text === undefined,
  });
}

/** Today's date as the `datetime1` field shows it in English: month/day/year. */
export function shortDate(now: Date = new Date()): string {
  return `${now.getMonth() + 1}/${now.getDate()}/${now.getFullYear()}`;
}

/** The Insert menu commands for slide numbers, footers, and dates. */
export type HeaderFooterCommand = Extract<
  EditorCommand,
  { readonly type: 'setSlideNumbers' | 'setFooter' | 'setDateTime' }
>;

/** Run a slide-number, footer, or date command. Returns its history label. */
export function applyHeaderFooter(batch: OpBatch, command: HeaderFooterCommand): string {
  const options = { indices: command.indices, hideOnTitleSlides: command.hideOnTitleSlides };
  switch (command.type) {
    case 'setSlideNumbers': {
      const targets = headerFooterTargets(batch.pkg, command.show, options);
      if (targets.show.length + targets.hide.length === 0)
        throw new CommandRefusal('errors.unsupported', 'slides');
      // Refused only when neither the slides that show nor those that hide can have a number.
      let refusal: unknown = new CommandRefusal('disabled.noSlideNumberPlaceholder');
      let done = false;
      for (const [show, indices] of [
        [true, targets.show],
        [false, targets.hide],
      ] as const) {
        if (indices.length === 0) continue;
        try {
          setSlideNumbers(batch, show, indices);
          done = true;
        } catch (error) {
          if (
            !(error instanceof CommandRefusal) ||
            error.reason !== 'disabled.noSlideNumberPlaceholder'
          )
            throw error;
          refusal = error;
        }
      }
      if (!done) throw refusal;
      return 'slideNumbers';
    }
    case 'setFooter':
      setFooter(batch, command.show, command.text, options);
      return 'footer';
    case 'setDateTime':
      setDateTime(batch, command.show, command.text, command.fixed, options);
      return 'dateTime';
  }
}
