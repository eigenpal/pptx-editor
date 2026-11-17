// A shape with `useBgFill` (ECMA-376 Part 1, 19.3.1.43) fills with the slide background, not
// with its style fill, and the attribute survives save and reopen.

import { describe, expect, test } from 'bun:test';
import { buildPptx, type FixtureSlide } from '../../../../../scripts/fixtures/builder.ts';
import { PptxDocumentStore, readPptx, writePptx, xmlPartOf } from '../../store/index.ts';
import { resolveSlide, type ResolvedElement } from '../index.ts';

const STYLE =
  '<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr>' +
  '</a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
  '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>' +
  '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>';

function shape(opts: { x?: number; cx?: number; fill?: string; rot?: number } = {}): string {
  const rot = opts.rot ? ` rot="${opts.rot}"` : '';
  return (
    '<p:sp useBgFill="1"><p:nvSpPr><p:cNvPr id="2056" name="Rectangle 74"/><p:cNvSpPr/>' +
    `<p:nvPr/></p:nvSpPr><p:spPr><a:xfrm${rot}><a:off x="${opts.x ?? 0}" y="0"/>` +
    `<a:ext cx="${opts.cx ?? 12192000}" cy="6858000"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${opts.fill ?? ''}` +
    `<a:ln><a:noFill/></a:ln></p:spPr>${STYLE}</p:sp>`
  );
}

const WHITE_BG =
  '<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>';
const GRADIENT_BG =
  '<p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"/></a:gs>' +
  '<a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/>' +
  '</a:gradFill></p:bgPr></p:bg>';

function open(slide: FixtureSlide) {
  const opened = readPptx(buildPptx({ width: 12192000, height: 6858000, slides: [slide] }));
  if (!opened.ok) throw new Error('open failed');
  return opened.pkg;
}

function fillOf(el: ResolvedElement | undefined) {
  if (!el || !('fill' in el)) throw new Error('no fill');
  return el.fill;
}

interface Node {
  readonly id?: string;
  readonly local?: string;
  readonly attrs?: readonly { readonly local: string; readonly value: string }[];
  readonly children?: readonly Node[];
}

function find(el: Node | undefined, local: string): Node | undefined {
  if (!el || typeof el !== 'object') return undefined;
  if (el.local === local) return el;
  for (const c of el.children ?? []) {
    const found = find(c, local);
    if (found) return found;
  }
  return undefined;
}

const lastShape = (pkg: ReturnType<typeof open>) => resolveSlide(pkg, 0)!.elements.at(-1);

describe('useBgFill', () => {
  test('a shape fills with a solid slide background, not its style fill', () => {
    const pkg = open({ shapes: shape(), layout: 'blank', background: WHITE_BG });
    expect(fillOf(lastShape(pkg))).toEqual({
      type: 'solid',
      color: { r: 255, g: 255, b: 255, a: 1 },
    });
  });

  test('a shape fills with the background that the slide inherits', () => {
    // The fixture master's background is `bg1` through `p:bgRef`: white.
    const pkg = open({ shapes: shape(), layout: 'blank' });
    expect(fillOf(lastShape(pkg))).toEqual(resolveSlide(pkg, 0)!.background);
  });

  test("a fill in the shape's own properties still applies", () => {
    const fill = '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>';
    const pkg = open({ shapes: shape({ fill }), layout: 'blank', background: WHITE_BG });
    expect(fillOf(lastShape(pkg))).toEqual({ type: 'solid', color: { r: 0, g: 255, b: 0, a: 1 } });
  });

  test('a gradient background keeps its place on the slide', () => {
    // The shape covers the right half of the slide: the gradient box extends one shape width
    // to the left.
    const pkg = open({
      shapes: shape({ x: 6096000, cx: 6096000 }),
      layout: 'blank',
      background: GRADIENT_BG,
    });
    const fill = fillOf(lastShape(pkg));
    expect(fill.type).toBe('linear');
    if (fill.type !== 'linear') return;
    expect(fill.tileRect).toEqual({ l: -1, t: 0, r: 0, b: 0 });
  });

  test('a turned shape over a gradient background shows the background through it', () => {
    const pkg = open({
      shapes: shape({ cx: 6096000, rot: 5400000 }),
      layout: 'blank',
      background: GRADIENT_BG,
    });
    expect(fillOf(lastShape(pkg))).toEqual({ type: 'none' });
  });

  test('a background edit reaches the cached shape', () => {
    const store = new PptxDocumentStore(
      open({ shapes: shape(), layout: 'blank', background: WHITE_BG })
    );
    const white = fillOf(lastShape(store.package));
    const part = resolveSlide(store.package, 0)!.partName;
    const color = find(xmlPartOf(store.package, part)!.root as Node, 'schemeClr')!;
    const name = { ns: '', local: 'val' };
    expect(
      store.transact([{ op: 'setAttribute', part, node: color.id!, name, value: 'accent2' }]).ok
    ).toBe(true);
    const after = resolveSlide(store.package, 0)!;
    expect(fillOf(after.elements.at(-1))).toEqual(after.background);
    expect(after.background).not.toEqual(white);
  });

  test('save keeps useBgFill', () => {
    const pkg = open({ shapes: shape(), layout: 'blank', background: WHITE_BG });
    const reopened = readPptx(writePptx(pkg));
    if (!reopened.ok) throw new Error('reopen failed');
    const part = resolveSlide(reopened.pkg, 0)!.partName;
    const sp = find(xmlPartOf(reopened.pkg, part)!.root as Node, 'sp')!;
    expect(sp.attrs?.some((a) => a.local === 'useBgFill' && a.value === '1')).toBe(true);
    expect(fillOf(lastShape(reopened.pkg))).toMatchObject({ color: { r: 255, g: 255, b: 255 } });
  });
});
