// Equation commands: insert an equation from its linear form, and replace the math of an
// equation with a new linear form. An equation is Office Math (`a14:m`, ECMA-376 Part 1, 22.1)
// in a text box, in `mc:AlternateContent` (ECMA-376 Part 3, 10.2.1): a `mc:Choice` that
// requires `a14` holds the math, and the `mc:Fallback` holds the same text box with the
// equation as plain text, for a reader that does not understand Office Math. An edit of the
// math removes the fallback copy, as every edit of a drawn `mc:Choice` copy does
// (`alternate-copies.ts`).

import {
  child,
  elements,
  escapeXmlAttribute,
  NS,
  specOf,
  type Element,
  type NodeSpec,
} from '../../store/index.ts';
import { isMathZone, isSupportedMath, MATH_NS } from '../../cascade/math.ts';
import { oMathParaOf, ommlOf, parseLinear, plainTextOf } from './equation-linear.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { nextShapeId, parentOf, spTreeOf } from './package-edits.ts';
import { specFromXml } from './shape-commands.ts';

/** The font of the math runs of a new equation: the math font that decks name most. */
export const DEFAULT_MATH_FONT = 'Cambria Math';

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

const esc = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** `Equation n`, from `n` up, that no shape of `tree` uses as its name. */
function freeName(tree: Element, n: number): string {
  const names = new Set<string>();
  const walk = (el: Element): void => {
    if (el.local === 'cNvPr') {
      const name = el.attrs.find((a) => a.local === 'name')?.value;
      if (name !== undefined) names.add(name);
    }
    for (const c of elements(el)) walk(c);
  };
  walk(tree);
  while (names.has(`Equation ${n}`)) n++;
  return `Equation ${n}`;
}

/**
 * Give each math run (`m:r/a:rPr`) of `spec` the attributes and children of `template`: the
 * font, size, and fill of the equation. The run keeps its own `lang` and `i`.
 */
function withRunFormat(spec: NodeSpec, template: NodeSpec | undefined, inRun = false): NodeSpec {
  const kids = spec.children?.map((c) =>
    'text' in c ? c : withRunFormat(c, template, spec.ns === MATH_NS && spec.local === 'r')
  );
  if (inRun && spec.ns === NS.a && spec.local === 'rPr' && template) {
    const own = new Set((spec.attrs ?? []).map((a) => a.local));
    const attrs = [
      ...(spec.attrs ?? []),
      ...(template.attrs ?? []).filter(
        (a) => !own.has(a.local) && a.local !== 'i' && a.local !== 'b'
      ),
    ];
    return { ...spec, attrs, children: template.children ?? [] };
  }
  return kids ? { ...spec, children: kids } : spec;
}

/** The run format of a new equation: its math font. */
function defaultRunFormat(font: string): NodeSpec {
  return specFromXml(
    `<a:rPr lang="en-US"><a:latin typeface="${escapeXmlAttribute(font)}"/></a:rPr>`
  );
}

/**
 * Insert an equation from its linear form (`equation-linear.ts`) in a new text box centered on
 * the slide. Returns the node id of the text box that the editor draws (in `mc:Choice`).
 */
export function insertEquation(
  batch: OpBatch,
  part: string,
  linear: string,
  slideW: number,
  slideH: number,
  font = DEFAULT_MATH_FONT
): string {
  if (!font.trim() || font.length > 64) throw new CommandRefusal('errors.unsupported', 'font');
  const row = parseLinear(linear);
  const tree = spTreeOf(batch, part);
  const id = nextShapeId(batch, tree);
  const cx = Math.round(slideW * 0.4);
  const cy = 731_520;
  const name = freeName(tree, id - 1);
  const nv = `<p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`;
  const spPr =
    `<p:spPr><a:xfrm><a:off x="${Math.round((slideW - cx) / 2)}" y="${Math.round((slideH - cy) / 2)}"/>` +
    `<a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>`;
  const body = '<a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>';
  const xml =
    `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:a14="${NS.a14}" Requires="a14">` +
    `<p:sp>${nv}${spPr}<p:txBody>${body}<a:p><a14:m>${oMathParaOf(linear)}</a14:m>` +
    '<a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp></mc:Choice>' +
    `<mc:Fallback><p:sp>${nv}${spPr}<p:txBody>${body}<a:p><a:pPr algn="ctr"/>` +
    `<a:r><a:rPr lang="en-US" dirty="0"/><a:t>${esc(plainTextOf(row))}</a:t></a:r>` +
    '</a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent>';
  const spec = withRunFormat(specFromXml(xml), defaultRunFormat(font));
  const ext = tree.children.findIndex((c) => c.type === 'element' && c.local === 'extLst');
  const acId = batch.insert({
    op: 'insertNode',
    part,
    parent: tree.id,
    index: ext < 0 ? tree.children.length : ext,
    spec,
  });
  const choice = elements(batch.element(part, acId))[0];
  const sp = choice ? child(choice, 'p:sp') : undefined;
  if (!sp) throw new CommandRefusal('errors.unsupported', 'equation');
  return sp.id;
}

/**
 * Give the fallback copy of a new equation the box of its math copy, after the math copy took
 * the size of its text (`spAutoFit`).
 */
export function matchFallbackBox(batch: OpBatch, part: string, spId: string): void {
  const sp = batch.element(part, spId);
  const ext = child(child(child(sp, 'p:spPr'), 'a:xfrm'), 'a:ext');
  const choice = parentOf(batch, part, spId);
  if (choice.local !== 'Choice') return;
  const fallback = elements(parentOf(batch, part, choice.id)).find((c) => c.local === 'Fallback');
  const other = child(child(child(child(fallback, 'p:sp'), 'p:spPr'), 'a:xfrm'), 'a:ext');
  if (!ext || !other) return;
  for (const local of ['cx', 'cy']) {
    const value = ext.attrs.find((a) => a.local === local)?.value;
    if (value !== undefined)
      batch.push({ op: 'setAttribute', part, node: other.id, name: { ns: '', local }, value });
  }
}

/** The `m:oMath` of a zone with one equation, or undefined. */
function onlyEquation(zone: Element): Element | undefined {
  const top = elements(zone);
  if (top.length !== 1) return undefined;
  const first = top[0]!;
  if (first.ns === MATH_NS && first.local === 'oMath') return first;
  if (first.ns !== MATH_NS || first.local !== 'oMathPara') return undefined;
  const lines = elements(first).filter((c) => c.ns === MATH_NS && c.local === 'oMath');
  return lines.length === 1 ? lines[0] : undefined;
}

/** The format of the first math run of `el`, as a spec, without its italic. */
function firstRunFormat(el: Element): NodeSpec | undefined {
  for (const c of elements(el)) {
    if (c.ns === MATH_NS && c.local === 'r') {
      const rPr = child(c, 'a:rPr');
      return rPr ? specOf(rPr) : undefined;
    }
    const inner = firstRunFormat(c);
    if (inner) return inner;
  }
  return undefined;
}

/**
 * Replace the math of the equation `zoneId` (an `a14:m`) with a linear form. The zone keeps its
 * paragraph properties (`m:oMathParaPr`) and the format of its first math run. Returns the node
 * id of the shape (`p:sp`) that holds the equation, to fit again.
 */
export function setEquation(batch: OpBatch, part: string, zoneId: string, linear: string): string {
  const zone = batch.element(part, zoneId);
  if (!isMathZone(zone) || !isSupportedMath(zone))
    throw new CommandRefusal('errors.unsupported', 'not an equation');
  const math = onlyEquation(zone);
  if (!math) throw new CommandRefusal('errors.unsupported', 'equation lines');
  const template = firstRunFormat(zone) ?? defaultRunFormat(DEFAULT_MATH_FONT);
  const row = parseLinear(linear);
  const spec = withRunFormat(
    specFromXml(`<m:oMath xmlns:m="${MATH_NS}">${ommlOf(row)}</m:oMath>`),
    template
  );
  for (const c of math.children) batch.push({ op: 'removeNode', part, node: c.id });
  (spec.children ?? []).forEach((c, index) => {
    if ('text' in c) return;
    batch.insert({ op: 'insertNode', part, parent: math.id, index, spec: c });
  });
  // a14:m → a:p → p:txBody → p:sp.
  let el = zone;
  for (let i = 0; i < 3; i++) el = parentOf(batch, part, el.id);
  if (el.ns !== NS.p || el.local !== 'sp')
    throw new CommandRefusal('errors.unsupported', 'not in a shape');
  return el.id;
}
