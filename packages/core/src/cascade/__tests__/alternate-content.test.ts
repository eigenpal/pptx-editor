// The branch of `mc:AlternateContent` that the engine reads (ECMA-376 Part 3, 10.2.1): the first
// `mc:Choice` whose `Requires` namespaces it understands and whose content it draws, else
// `mc:Fallback`. Both branches survive save unchanged.

import { describe, expect, test } from 'bun:test';
import {
  buildPptx,
  NS_DECL,
  para,
  run,
  shape,
  solid,
} from '../../../../../scripts/fixtures/builder.ts';
import {
  attr,
  canonicalPackageFingerprint,
  child,
  elements,
  find,
  parseXmlPart,
  readPptx,
  serializeXmlPart,
  writePptx,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import { alternateBranch, resolveSlide } from '../index.ts';

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const A14 = 'http://schemas.microsoft.com/office/drawing/2010/main';
const VML = 'urn:schemas-microsoft-com:vml';
const IN = 914_400;

const box = (id: number, color: string, text: string): string =>
  shape({ id, x: IN, y: IN, cx: 3 * IN, cy: IN, fill: solid(color), text: para(run(text)) });

const choice = (requires: string, decls: string, inner: string): string =>
  `<mc:Choice ${decls} Requires="${requires}">${inner}</mc:Choice>`;

const alternate = (branches: string): string =>
  `<mc:AlternateContent xmlns:mc="${MC}">${branches}</mc:AlternateContent>`;

/** The `mc:AlternateContent` of a slide whose root has `rootDecls`, and that root. */
function parse(branches: string, rootDecls = ''): { ac: Element; root: Element } {
  const xml = `<p:sld ${NS_DECL} ${rootDecls}><p:cSld><p:spTree>${alternate(branches)}</p:spTree></p:cSld></p:sld>`;
  const r = parseXmlPart('/ppt/slides/slide1.xml', xml);
  if (!r.ok) throw new Error('parse failed');
  return { ac: find(r.part.root, 'mc:AlternateContent')!, root: r.part.root };
}

/** The text of the shape in the branch that the engine reads. */
function readText(branches: string, rootDecls = '', withRoot = true): string | undefined {
  const { ac, root } = parse(branches, rootDecls);
  const sp = elements(alternateBranch(ac, withRoot ? root : undefined))[0];
  return sp ? (find(sp, 'a:t')?.children[0] as { value?: string } | undefined)?.value : undefined;
}

const fallback = `<mc:Fallback>${box(3, '70AD47', 'fallback')}</mc:Fallback>`;

describe('alternateBranch', () => {
  test('reads a Choice that requires p14 and holds shapes', () => {
    expect(readText(choice('p14', `xmlns:p14="${P14}"`, box(3, 'C00000', 'p14')) + fallback)).toBe(
      'p14'
    );
  });

  test('resolves a prefix on the part root, and not without it', () => {
    const branches = choice('p14', '', box(3, 'C00000', 'p14')) + fallback;
    expect(readText(branches, `xmlns:p14="${P14}"`)).toBe('p14');
    expect(readText(branches, `xmlns:p14="${P14}"`, false)).toBe('fallback');
    expect(readText(branches)).toBe('fallback');
  });

  test('resolves a prefix on the nearest element above that declares it', () => {
    const branches = choice('p14', '', box(3, 'C00000', 'p14')) + fallback;
    const inGroup = (decl: string): string => {
      const xml =
        `<p:sld ${NS_DECL} xmlns:p14="${P14}"><p:cSld><p:spTree>` +
        `<p:grpSp ${decl}>${alternate(branches)}</p:grpSp></p:spTree></p:cSld></p:sld>`;
      const r = parseXmlPart('/ppt/slides/slide1.xml', xml);
      if (!r.ok) throw new Error('parse failed');
      const ac = find(r.part.root, 'mc:AlternateContent')!;
      const sp = elements(alternateBranch(ac, r.part.root))[0];
      return (find(sp, 'a:t')?.children[0] as { value: string }).value;
    };
    // The root binds `p14`; no element between it and the wrapper rebinds it.
    expect(inGroup('')).toBe('p14');
    // A group that binds `p14` to another namespace hides the root's binding.
    expect(inGroup('xmlns:p14="urn:example"')).toBe('fallback');
  });

  test('reads the Fallback for a namespace it does not understand', () => {
    expect(readText(choice('v', `xmlns:v="${VML}"`, box(3, 'C00000', 'v')) + fallback)).toBe(
      'fallback'
    );
    expect(readText(choice('a14', `xmlns:a14="${A14}"`, box(3, 'C00000', 'a14')) + fallback)).toBe(
      'fallback'
    );
    // A prefix bound to an unknown namespace is not p14, whatever its name.
    expect(
      readText(choice('p14', 'xmlns:p14="urn:example"', box(3, 'C00000', 'x')) + fallback)
    ).toBe('fallback');
  });

  test('every required prefix must be understood', () => {
    const decls = `xmlns:p14="${P14}" xmlns:v="${VML}"`;
    expect(readText(choice('p14 v', decls, box(3, 'C00000', 'both')) + fallback)).toBe('fallback');
  });

  test('takes the first Choice that qualifies, in order', () => {
    const branches =
      choice('v', `xmlns:v="${VML}"`, box(3, '000000', 'v')) +
      choice('p14', `xmlns:p14="${P14}"`, box(3, 'C00000', 'second')) +
      fallback;
    expect(readText(branches)).toBe('second');
  });

  test('reads the Fallback when the Choice holds content the cascade does not draw', () => {
    const ink =
      '<p14:contentPart xmlns:p14="' +
      P14 +
      '" r:id="rId9"><p14:nvContentPartPr><p14:cNvPr id="3" name="Ink 2"/><p14:cNvContentPartPr/><p14:nvPr/></p14:nvContentPartPr></p14:contentPart>';
    expect(readText(choice('p14', `xmlns:p14="${P14}"`, ink) + fallback)).toBe('fallback');
  });

  test('a Choice without Requires is never selected', () => {
    const branches = `<mc:Choice>${box(3, 'C00000', 'none')}</mc:Choice>${fallback}`;
    expect(readText(branches)).toBe('fallback');
  });

  test('with no Fallback, the first Choice is read', () => {
    expect(readText(choice('v', `xmlns:v="${VML}"`, box(3, 'C00000', 'only')))).toBe('only');
  });
});

describe('a slide with mc:AlternateContent', () => {
  const deck = (): Uint8Array =>
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: alternate(
            choice('p14', `xmlns:p14="${P14}"`, box(3, 'C00000', 'Choice branch')) +
              `<mc:Fallback>${box(3, '70AD47', 'Fallback branch')}</mc:Fallback>`
          ),
        },
      ],
    });

  test('draws the Choice shape', () => {
    const r = readPptx(deck());
    if (!r.ok) throw new Error('open failed');
    const slide = resolveSlide(r.pkg, 0)!;
    expect(slide.elements).toHaveLength(1);
    const el = slide.elements[0];
    if (el.kind !== 'shape') throw new Error('not a shape');
    expect(el.fill).toMatchObject({ type: 'solid', color: { r: 192, g: 0, b: 0 } });
    const choiceSp = find(find(xmlPartOf(r.pkg, slide.partName)!.root, 'mc:Choice')!, 'p:sp')!;
    expect(el.nodeId).toBe(choiceSp.id);
  });

  test('keeps both branches through save and reopen', () => {
    const r = readPptx(deck());
    if (!r.ok) throw new Error('open failed');
    const again = readPptx(writePptx(r.pkg));
    if (!again.ok) throw new Error('reopen failed');
    expect(canonicalPackageFingerprint(again.pkg)).toBe(canonicalPackageFingerprint(r.pkg));
    const part = '/ppt/slides/slide1.xml';
    const xml = serializeXmlPart(xmlPartOf(again.pkg, part)!);
    expect(xml).toContain('Choice branch');
    expect(xml).toContain('Fallback branch');
    const ac = find(xmlPartOf(again.pkg, part)!.root, 'mc:AlternateContent')!;
    expect(attr(child(ac, { ns: MC, local: 'Choice' }), 'Requires')).toBe('p14');
  });
});
