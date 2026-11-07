// Insert > Footer and Insert > Date and time, applied to all slides, as a Header and Footer
// dialog applies them.
//
// A slide shows its footer and its date through the footer (`p:ph type="ftr"`) and date
// (`p:ph type="dt"`) placeholders that its layout defines, or its master when the layout has
// none (ECMA-376 Part 1, 19.3.1.36 ph). Turning one on gives each slide that placeholder, so
// position and style come from the layout. A slide that has one already gets the new text, and
// keeps its own box and format. Turning one off removes the slide's placeholders of that type.
// A slide whose layout and master define no such placeholder is skipped.

import {
  attr,
  elements,
  find,
  path,
  presentationInfo,
  textOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { nextShapeId, spTreeOf, xmlAttr, xmlText } from './package-edits.ts';
import { specFromXml } from './shape-commands.ts';
import { removeShapes } from './shape-removal.ts';
import { setSlideNumbers } from './slide-numbers.ts';
import type { EditorCommand } from '../../contracts/neutral.ts';

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

/** The text of the footer of slide `index`, or null when it shows none. */
export function footerText(pkg: PptxPackage, index: number): string | null {
  const shape = shapesOfKind(pkg, presentationInfo(pkg).slides[index]?.partName, 'ftr')[0];
  return shape ? textOf(find(shape, 'p:txBody')) : null;
}

/** Whether slide `index` shows a date. */
export function hasDate(pkg: PptxPackage, index: number): boolean {
  return shapesOfKind(pkg, presentationInfo(pkg).slides[index]?.partName, 'dt').length > 0;
}

/** A field GUID in braces (`a:fld/@id`, `ST_Guid`). */
function fieldId(): string {
  const hex = (n: number) =>
    Array.from({ length: n }, () => Math.floor(Math.random() * 16).toString(16)).join('');
  return `{${hex(8)}-${hex(4)}-4${hex(3)}-${'89AB'[Math.floor(Math.random() * 4)]}${hex(3)}-${hex(12)}}`.toUpperCase();
}

/** The paragraph of a footer or date: plain text, or a date field that updates on open. */
function paragraphXml(kind: HeaderFooterKind, text: string, field: boolean): string {
  const t = xmlText(text, NAMES[kind]);
  const content =
    kind === 'dt' && field
      ? `<a:fld id="${fieldId()}" type="datetime1"><a:rPr lang="en-US"/><a:t>${t}</a:t></a:fld>`
      : `<a:r><a:rPr lang="en-US"/><a:t>${t}</a:t></a:r>`;
  return `<a:p>${content}<a:endParaRPr lang="en-US"/></a:p>`;
}

const TEXT_CONTENT = new Set(['r', 'fld', 'br']);

/**
 * Show `text` in an existing placeholder. The first paragraph and its first run or field stay,
 * so the placeholder keeps its paragraph and character format; other text is removed. A date
 * that changes between a field and fixed text gets a new run in place of the old one.
 */
function replaceText(
  batch: OpBatch,
  part: string,
  shape: Element,
  kind: HeaderFooterKind,
  text: string,
  field: boolean
): void {
  const body = find(shape, 'p:txBody');
  if (!body) return;
  const [first, ...rest] = elements(body).filter((c) => c.local === 'p');
  for (const p of rest) batch.push({ op: 'removeNode', part, node: p.id });
  if (!first) {
    batch.insert({
      op: 'insertNode',
      part,
      parent: body.id,
      index: body.children.length,
      spec: specFromXml(paragraphXml(kind, text, field)),
    });
    return;
  }
  const content = elements(first).filter((c) => TEXT_CONTENT.has(c.local));
  const run = content.find((c) => c.local === 'r' || c.local === 'fld');
  const wanted = kind === 'dt' && field ? 'fld' : 'r';
  const t = run?.local === wanted ? find(run, 'a:t') : undefined;
  if (run && t) {
    xmlText(text, NAMES[kind]);
    batch.push({ op: 'setText', part, node: t.id, text });
  } else {
    const xml = paragraphXml(kind, text, field);
    const inner = /<a:p>(.*)<a:endParaRPr/.exec(xml)![1]!;
    const end = first.children.findIndex((c) => c.type === 'element' && c.local === 'endParaRPr');
    batch.insert({
      op: 'insertNode',
      part,
      parent: first.id,
      index: end < 0 ? first.children.length : end,
      spec: specFromXml(inner),
    });
  }
  for (const c of content) if (c !== run || !t) batch.push({ op: 'removeNode', part, node: c.id });
}

/**
 * Show `text` in the placeholder of `kind` on the given slides (all slides by default), or hide
 * it with null. A date is a field unless `fixed`. Returns the number of slides changed. Refuses
 * when no slide has such a placeholder on its layout or master.
 */
export function setHeaderFooter(
  batch: OpBatch,
  kind: HeaderFooterKind,
  text: string | null,
  options: { readonly fixed?: boolean; readonly indices?: readonly number[] } = {}
): number {
  if (text !== null && text.length > MAX_HEADER_FOOTER_TEXT)
    throw new CommandRefusal('errors.unsupported', NAMES[kind]);
  const slides = presentationInfo(batch.pkg).slides;
  const targets = options.indices ?? slides.map((_, i) => i);
  let changed = 0;
  let possible = 0;
  for (const index of targets) {
    const slide = slides[index];
    if (!slide) continue;
    const tree = spTreeOf(batch, slide.partName);
    const existing = elements(tree).filter(isKind(kind));
    if (text === null) {
      removeShapes(
        batch,
        slide.partName,
        existing.map((el) => el.id)
      );
      changed += existing.length > 0 ? 1 : 0;
      possible += 1;
      continue;
    }
    const template =
      shapesOfKind(batch.pkg, slide.layoutPart, kind)[0] ??
      shapesOfKind(batch.pkg, slide.masterPart, kind)[0];
    if (!template && existing.length === 0) continue;
    possible += 1;
    const field = !options.fixed;
    if (existing.length > 0) {
      for (const shape of existing) replaceText(batch, slide.partName, shape, kind, text, field);
      changed += 1;
      continue;
    }
    const ph = path(template!, 'p:nvSpPr', 'p:nvPr', 'p:ph');
    const phAttrs = ['type', 'orient', 'sz', 'idx']
      .map((a) => {
        const value = attr(ph, a);
        return value !== undefined ? ` ${a}="${xmlAttr(value, 'placeholder')}"` : '';
      })
      .join('');
    const id = nextShapeId(batch, tree);
    const xml =
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${NAMES[kind]} Placeholder ${id - 1}"/>` +
      `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${phAttrs}/></p:nvPr></p:nvSpPr>` +
      `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphXml(kind, text, field)}` +
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

/** Today's date as the `datetime1` field shows it in English: month/day/year. */
export function shortDate(now: Date = new Date()): string {
  return `${now.getMonth() + 1}/${now.getDate()}/${now.getFullYear()}`;
}

/** The Insert menu commands for slide numbers, footers, and dates. */
export type HeaderFooterCommand = Extract<
  EditorCommand,
  { readonly type: 'setSlideNumbers' | 'setFooter' | 'setDateTime' }
>;

/** Run a slide-number, footer, or date command on all slides. Returns its history label. */
export function applyHeaderFooter(batch: OpBatch, command: HeaderFooterCommand): string {
  switch (command.type) {
    case 'setSlideNumbers':
      setSlideNumbers(batch, command.show);
      return 'slideNumbers';
    case 'setFooter':
      setHeaderFooter(batch, 'ftr', command.text);
      return 'footer';
    case 'setDateTime': {
      const text = command.show ? (command.text ?? shortDate()) : null;
      setHeaderFooter(batch, 'dt', text, { fixed: command.fixed });
      return 'dateTime';
    }
  }
}
