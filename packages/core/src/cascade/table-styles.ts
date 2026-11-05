// Table style lookup (design D1 of `table-rendering`). A saved deck holds the definition of every
// table style a table uses in `ppt/tableStyles.xml`, so styles come from that part. One
// default style is built in, because other tools reference it without writing it.

import {
  attr,
  elements,
  parseXmlPart,
  REL,
  relationshipsOf,
  textOf,
  xmlPartOf,
  child,
  type Element,
  type PptxPackage,
} from '../store/index.ts';

/** Medium Style 2 - Accent 1: the default table style of a new table. */
export const DEFAULT_TABLE_STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';

const lt1Line = (w: number): string =>
  `<a:ln w="${w}" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln>`;
const accent1 = (tint?: number): string =>
  `<a:fill><a:solidFill><a:schemeClr val="accent1">${tint ? `<a:tint val="${tint}"/>` : ''}</a:schemeClr></a:solidFill></a:fill>`;
const emphasis =
  '<a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle>';

/**
 * The built-in definition of the default style, Medium Style 2 - Accent 1. An APPROXIMATION
 * until a reference PDF confirms it (listed on `docs/site/content/pptx-fidelity.mdx`).
 */
const DEFAULT_STYLE_XML =
  `<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" styleId="${DEFAULT_TABLE_STYLE_ID}" styleName="Medium Style 2 - Accent 1">` +
  '<a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="dk1"/></a:tcTxStyle>' +
  `<a:tcStyle><a:tcBdr>${['left', 'right', 'top', 'bottom', 'insideH', 'insideV']
    .map((side) => `<a:${side}>${lt1Line(12700)}</a:${side}>`)
    .join('')}</a:tcBdr>${accent1(20000)}</a:tcStyle></a:wholeTbl>` +
  `<a:band1H><a:tcStyle><a:tcBdr/>${accent1(40000)}</a:tcStyle></a:band1H>` +
  '<a:band2H><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2H>' +
  `<a:band1V><a:tcStyle><a:tcBdr/>${accent1(40000)}</a:tcStyle></a:band1V>` +
  '<a:band2V><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2V>' +
  `<a:lastCol>${emphasis}<a:tcStyle><a:tcBdr/>${accent1()}</a:tcStyle></a:lastCol>` +
  `<a:firstCol>${emphasis}<a:tcStyle><a:tcBdr/>${accent1()}</a:tcStyle></a:firstCol>` +
  `<a:lastRow>${emphasis}<a:tcStyle><a:tcBdr><a:top>${lt1Line(38100)}</a:top></a:tcBdr>${accent1()}</a:tcStyle></a:lastRow>` +
  `<a:firstRow>${emphasis}<a:tcStyle><a:tcBdr><a:bottom>${lt1Line(38100)}</a:bottom></a:tcBdr>${accent1()}</a:tcStyle></a:firstRow>` +
  '</a:tblStyle>';

let builtIn: Element | undefined;
function defaultStyle(): Element | undefined {
  if (!builtIn) {
    const r = parseXmlPart('/builtin/tableStyle.xml', DEFAULT_STYLE_XML);
    builtIn = r.ok ? r.part.root : undefined;
  }
  return builtIn;
}

const cache = new WeakMap<Element, Map<string, Element>>();

/** The table styles of a package, by `styleId`. */
function stylesOf(pkg: PptxPackage): Map<string, Element> {
  const root = tableStylesRoot(pkg);
  if (!root) return new Map();
  const hit = cache.get(root);
  if (hit) return hit;
  const map = new Map<string, Element>();
  for (const style of elements(root))
    if (style.local === 'tblStyle') {
      const id = attr(style, 'styleId');
      if (id && !map.has(id)) map.set(id, style);
    }
  cache.set(root, map);
  return map;
}

/** The `ppt/tableStyles.xml` root of a package, when the presentation has one. */
function tableStylesRoot(pkg: PptxPackage): Element | undefined {
  const rel = relationshipsOf(pkg, '/ppt/presentation.xml').find(
    (r) => r.type === REL.tableStyles && !r.external
  );
  return rel?.target ? xmlPartOf(pkg, rel.target)?.root : undefined;
}

/**
 * The style that a new table takes: `a:tblStyleLst/@def` (ECMA-376 Part 1, 20.1.4.2.27),
 * else Medium Style 2 - Accent 1. A `def` that the package does not define is a built-in style
 * the editor cannot draw (the definitions are not in the file or the specification), so a new
 * table takes Medium Style 2 - Accent 1, which the editor has a built-in definition for.
 */
export function defaultTableStyleId(pkg: PptxPackage): string {
  const def = attr(tableStylesRoot(pkg), 'def')?.trim();
  return def && STYLE_ID.test(def) && stylesOf(pkg).has(def) ? def : DEFAULT_TABLE_STYLE_ID;
}

/** A table style id: a GUID in braces, upper case as `ST_Guid` allows. @internal */
export const STYLE_ID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/;

/**
 * The styles a table can use: the deck's own, in file order, and the built-in default. A deck
 * style carries its `styleName`; the built-in default is marked `builtIn` and has no name, so
 * the chrome names it in the user's language.
 */
export function tableStyleList(
  pkg: PptxPackage
): readonly { id: string; name: string; builtIn?: true }[] {
  const out: { id: string; name: string; builtIn?: true }[] = [...stylesOf(pkg).values()].map(
    (s) => ({ id: attr(s, 'styleId') ?? '', name: attr(s, 'styleName') ?? '' })
  );
  if (!out.some((s) => s.id === DEFAULT_TABLE_STYLE_ID))
    out.unshift({ id: DEFAULT_TABLE_STYLE_ID, name: '', builtIn: true });
  return out.filter((s) => STYLE_ID.test(s.id));
}

/** The style a table uses: inline `a:tableStyle`, else `a:tableStyleId` looked up. */
export function tableStyleFor(pkg: PptxPackage, tblPr: Element | undefined): Element | undefined {
  const inline = child(tblPr, 'a:tableStyle');
  if (inline) return inline;
  const id = textOf(child(tblPr, 'a:tableStyleId')).trim();
  if (!id) return undefined;
  return stylesOf(pkg).get(id) ?? (id === DEFAULT_TABLE_STYLE_ID ? defaultStyle() : undefined);
}
