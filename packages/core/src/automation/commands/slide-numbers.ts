// Insert > Slide numbers, applied to all slides, as a Header and Footer dialog applies them.
//
// A slide shows its number through the slide-number placeholder (`p:ph type="sldNum"`) that
// its layout defines (or its master, when the layout has none). Turning numbers on gives each
// slide that placeholder, holding a `slidenum` field, so position and style come from the
// layout; turning them off removes the slide's slide-number placeholders. A
// slide whose layout and master define no slide-number placeholder is skipped.

import {
  attr,
  elements,
  find,
  path,
  presentationInfo,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { fieldId, nextShapeId, spTreeOf, xmlAttr } from './package-edits.ts';
import { specFromXml } from './shape-commands.ts';
import { removeShapes } from './shape-removal.ts';

const isSlideNumber = (el: Element): boolean =>
  attr(path(el, 'p:nvSpPr', 'p:nvPr', 'p:ph'), 'type') === 'sldNum';

/** The slide-number placeholder a layout or master defines. */
function templateIn(pkg: PptxPackage, part: string | undefined): Element | undefined {
  const tree = part ? find(xmlPartOf(pkg, part)?.root, 'p:spTree') : undefined;
  return tree ? elements(tree).find(isSlideNumber) : undefined;
}

/** Whether slide `index` shows its number. */
export function hasSlideNumber(pkg: PptxPackage, index: number): boolean {
  const part = presentationInfo(pkg).slides[index]?.partName;
  const tree = part ? find(xmlPartOf(pkg, part)?.root, 'p:spTree') : undefined;
  return !!tree && elements(tree).some(isSlideNumber);
}

/**
 * Show or hide slide numbers on the given slides (all slides by default). Returns the number of
 * slides changed. Refuses when no slide can show a number.
 */
export function setSlideNumbers(
  batch: OpBatch,
  show: boolean,
  indices?: readonly number[]
): number {
  const slides = presentationInfo(batch.pkg).slides;
  const targets = indices ?? slides.map((_, i) => i);
  let changed = 0;
  let possible = 0;
  for (const index of targets) {
    const slide = slides[index];
    if (!slide) continue;
    const tree = spTreeOf(batch, slide.partName);
    const existing = elements(tree).filter(isSlideNumber);
    const template =
      templateIn(batch.pkg, slide.layoutPart) ?? templateIn(batch.pkg, slide.masterPart);
    if (!template && existing.length === 0) continue;
    possible += 1;
    if (!show) {
      removeShapes(
        batch,
        slide.partName,
        existing.map((el) => el.id)
      );
      changed += existing.length > 0 ? 1 : 0;
      continue;
    }
    if (existing.length > 0 || !template) continue;
    const ph = path(template, 'p:nvSpPr', 'p:nvPr', 'p:ph');
    const phAttrs = ['type', 'orient', 'sz', 'idx']
      .map((a) => {
        const value = attr(ph, a);
        return value !== undefined ? ` ${a}="${xmlAttr(value, 'placeholder')}"` : '';
      })
      .join('');
    const id = nextShapeId(batch, tree);
    const xml =
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Slide Number Placeholder ${id - 1}"/>` +
      `<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${phAttrs}/></p:nvPr></p:nvSpPr>` +
      '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>' +
      `<a:fld id="${fieldId()}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld>` +
      '<a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>';
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
  if (possible === 0) throw new CommandRefusal('disabled.noSlideNumberPlaceholder');
  return changed;
}
